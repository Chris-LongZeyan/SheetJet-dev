import pytest

from sheetjet import SheetJetError, Workbook


def test_cache_paths_with_quotes_keep_exact_read_allowlist(model, tmp_path):
    import duckdb

    cache = tmp_path / "analyst's cache"
    secret = cache / "other.txt"
    with Workbook(model, cache) as book:
        engine = book.query_engine
        engine.load("Sales", columns=["Revenue"], schema={"Revenue": "BIGINT"})
        secret.write_text("not a projection")
        assert engine.query("SELECT sum(Revenue) FROM Sales")["rows"] == [[51500]]
        with pytest.raises(duckdb.Error, match="disabled"):
            engine.query("SELECT * FROM read_text(?)", [str(secret)])
        assert engine.query("SELECT current_setting('lock_configuration')")["rows"] == [[True]]


def test_scalar_query_does_not_import_optional_dataframe_stack(model, cache):
    import subprocess
    import sys

    script = """
import sys
from sheetjet import Workbook
with Workbook(sys.argv[1], sys.argv[2]) as book:
    engine = book.query_engine
    engine.load('Sales', columns=['Revenue'], schema={'Revenue': 'BIGINT'})
    assert engine.query('SELECT sum(Revenue) FROM Sales')['rows'] == [[51500]]
    engine.materialize('Sales')
    assert engine.query('SELECT count(*) FROM Sales')['rows'] == [[100]]
    assert engine.query('SELECT ? AS value', ['APAC'])["rows"] == [['APAC']]
    assert 'pandas' not in sys.modules
"""
    result = subprocess.run(
        [sys.executable, "-c", script, str(model), str(cache)], capture_output=True, text=True
    )
    assert result.returncode == 0, result.stdout + result.stderr


@pytest.mark.parametrize(
    "value",
    ["", "APAC", "O'Brien", "日本語 😀", "a\\b\n\t? $1", "\\'); DROP TABLE Sales; --", "nul\0text"],
)
def test_text_parameters_match_native_binding(model, cache, value):
    import duckdb

    sql = "SELECT ? AS value, '?' AS literal, length(?) AS size"
    with duckdb.connect() as native, Workbook(model, cache) as book:
        expected = [list(row) for row in native.execute(sql, [value, value]).fetchall()]
        assert book.query_engine.query(sql, [value, value])["rows"] == expected
        assert book.query_engine.query("SELECT $2, $1", [value, "second"])["rows"] == [
            ["second", value]
        ]


def test_text_parameter_errors_leave_connection_reusable(model, cache):
    import duckdb

    with Workbook(model, cache) as book:
        engine = book.query_engine
        for sql, params in [("SELECT ?, ?", ["only one"]), ("SELECT ?::INTEGER", ["invalid"])]:
            with pytest.raises(duckdb.Error):
                engine.query(sql, params)
            assert engine.query("SELECT ?", ["recovered"])["rows"] == [["recovered"]]
        assert engine.query("SELECT $region", {"region": "named"})["rows"] == [["named"]]
        assert engine.query("SELECT ?, ?", ["mixed", 42])["rows"] == [["mixed", 42]]


@pytest.mark.parametrize("value", ["2", "invalid"])
@pytest.mark.parametrize(
    "sql",
    [
        "SELECT ? + 1",
        "SELECT coalesce(?,1)",
        "SELECT typeof(?)",
        "SELECT ? = 2",
        "SELECT greatest(?,2)",
    ],
)
def test_text_parameter_coercion_matches_native_binding(model, cache, sql, value):
    import duckdb

    with duckdb.connect() as native, Workbook(model, cache) as book:
        try:
            expected = [list(row) for row in native.execute(sql, [value]).fetchall()]
        except duckdb.Error as exc:
            with pytest.raises(type(exc)):
                book.query_engine.query(sql, [value])
        else:
            assert book.query_engine.query(sql, [value])["rows"] == expected


def test_prepared_text_query_rebinds_and_survives_relation_changes(model, cache):
    with Workbook(model, cache) as book:
        engine = book.query_engine
        args = {
            "columns": ["Region", "Revenue"],
            "schema": {"Region": "VARCHAR", "Revenue": "BIGINT"},
        }
        engine.load("Sales", **args)
        sql = 'SELECT SUM("Revenue") FROM Sales WHERE "Region" = ?'
        for region, total in [("APAC", 25500), ("EMEA", 26000), ("APAC", 25500)]:
            assert engine.query(sql, [region])["rows"] == [[total]]
        engine.materialize("Sales")
        assert engine.query(sql, ["EMEA"])["rows"] == [[26000]]
        # Staging another relation reopens the writable connection, then query
        # must prepare again on the new restricted connection.
        engine.load("Other", sheet="Transactions", ref="A1:E101", **args)
        assert engine.query(sql, ["APAC"])["rows"] == [[25500]]
        assert engine.query("SELECT ?", ["different query"])["rows"] == [["different query"]]
        assert engine.query(sql, ["EMEA"])["rows"] == [[26000]]


def test_lazy_cache_views_and_query_session_restrictions(model, cache, tmp_path):
    import duckdb

    secret = tmp_path / "unrelated.csv"
    secret.write_text("secret\nnot workbook data\n")
    args = {"columns": ["Revenue"], "schema": {"Revenue": "BIGINT"}}
    for _ in range(2):
        with Workbook(model, cache) as book:
            engine = book.query_engine
            engine.load("Sales", **args)
            assert engine.query("SELECT count(*) FROM duckdb_tables() WHERE table_name='Sales'")[
                "rows"
            ] == [[0]]
            assert engine.query("SELECT sum(Revenue) FROM Sales")["rows"] == [[51500]]
            with pytest.raises(duckdb.Error, match="disabled"):
                engine.query("SELECT * FROM read_csv(?)", [str(secret)])
            assert engine.query("SELECT current_setting('enable_external_access')")["rows"] == [
                [False]
            ]
            # A denied query must leave subsequent queries restricted and usable.
            assert engine.query("SELECT count(*) FROM Sales")["rows"] == [[100]]
            assert not list(cache.glob("index-*.sqlite"))
            engine.load("Sales", columns=["Region"], schema={"Region": "VARCHAR"})
            assert engine.query("SELECT count(DISTINCT Region) FROM Sales")["rows"] == [[2]]


def test_loaded_projection_change_fails_closed(model, cache):
    with Workbook(model, cache) as book:
        book.query_engine.load("Sales", columns=["Revenue"], schema={"Revenue": "BIGINT"})
        path = next(cache.glob("projection-*.parquet"))
        path.write_bytes(b"changed")
        with pytest.raises(SheetJetError, match="Projection cache changed"):
            book.query_engine.query("SELECT sum(Revenue) FROM Sales")


def test_failed_projection_reload_keeps_published_data(model, cache):
    with Workbook(model, cache) as book:
        engine = book.query_engine
        engine.load("Sales", columns=["Revenue"], schema={"Revenue": "BIGINT"})
        with pytest.raises(Exception, match="Conversion|convert"):
            engine.load("Sales", columns=["Region"], schema={"Region": "BIGINT"})
        assert engine.query("SELECT sum(Revenue) FROM Sales")["rows"] == [[51500]]


def test_empty_projection_and_cache_opt_out(model, cache):
    with Workbook(model, cache) as book:
        engine = book.query_engine
        engine.load(
            "Empty",
            sheet="Transactions",
            ref="A1:A1",
            schema={"Region": "VARCHAR"},
            persistent_cache=False,
        )
        assert engine.query("SELECT count(*) FROM Empty")["rows"] == [[0]]
        assert not list(cache.glob("projection-*.parquet"))


def test_user_relation_name_cannot_collide_with_staging_view(model, cache):
    with Workbook(model, cache) as book:
        engine = book.query_engine
        engine.load(
            "_sheetjet_stage",
            sheet="Transactions",
            ref="A1:C101",
            columns=["Revenue"],
            schema={"Revenue": "BIGINT"},
        )
        assert engine.query("SELECT sum(Revenue) FROM _sheetjet_stage")["rows"] == [[51500]]


def test_materialize_reload_and_parameterized_queries(model, cache):
    import duckdb

    with Workbook(model, cache) as book:
        engine = book.query_engine
        engine.load("Sales", columns=["Revenue"], schema={"Revenue": "BIGINT"})
        assert not engine.materialize("Sales")["cached"]
        assert engine.materialize("Sales")["cached"]
        assert engine.query("SELECT sum(Revenue) FROM Sales WHERE Revenue > ?", [1000])["rows"] == [
            [1010]
        ]
        with pytest.raises(duckdb.Error, match="disabled"):
            engine.query("SELECT * FROM read_text('unrelated.txt')")
        assert engine.query("SELECT count(*) FROM Sales")["rows"] == [[100]]
        # Replacing a materialized relation with a different projection remains atomic.
        engine.load("Sales", columns=["Customer"], schema={"Customer": "VARCHAR"})
        assert engine.query("SELECT count(DISTINCT Customer) FROM Sales")["rows"] == [[5]]
        engine.materialize("Sales")
        with pytest.raises(SheetJetError, match="already loaded"):
            engine.materialize("Missing")


def test_materialized_union_can_reload_without_stale_schema(model, cache):
    with Workbook(model, cache) as book:
        engine = book.query_engine
        ranges = {"Transactions": "A1:C101"}
        engine.load_sheets("Combined", ranges, {"Revenue": "BIGINT"})
        engine.materialize("Combined")
        assert engine.query("SELECT sum(Revenue) FROM Combined")["rows"] == [[51500]]
        with pytest.raises(SheetJetError, match="Formula inputs"):
            engine.load_sheets("Combined", {"Transactions": "A1:E101"}, {"Profit": "BIGINT"})
        assert engine.query("SELECT sum(Revenue) FROM Combined")["rows"] == [[51500]]
        engine.load_sheets("Combined", ranges, {"Region": "VARCHAR"})
        assert engine.query("SELECT DISTINCT _sheet FROM Combined")["rows"] == [["Transactions"]]
