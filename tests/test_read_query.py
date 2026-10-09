import json
import os

import pytest

from sheetjet import BudgetExceeded, SheetJetError, Workbook
from sheetjet.core import Budget, bounds, compact


def test_fast_inspection_no_cell_scan(model, cache):
    with Workbook(model, cache) as w:
        meta = w.inspect_workbook()
        assert meta["sheet_count"] == 3
        assert meta["features"]["charts"] == 1
        assert w.metrics.counts.get("cells_scanned", 0) == 0
        assert w.find_names("Growth")[0]["formula"] == "'Assumptions'!$B$1"
        assert w.find_table("sal")[0]["range"] == "A1:E101"


def test_index_patterns_and_persistent_cache(model, cache):
    with Workbook(model, cache) as w:
        deep = w.inspect_sheet("Transactions", deep=True)
        assert deep["formula_count"] == 100
        assert deep["nonempty_cells"] == 505
        assert deep["hidden_rows_count"] == 1
        assert w.formula_patterns("Transactions")[0]["count"] == 100
        result = w.find_text("China revenue growth assumption", ["Assumptions"])
        assert result["matches"][0]["cell"] == "A1"
        assert not w.find_text("Need full search", ["Assumptions"])["matches"]
        assert (
            w.find_text("Need full search", ["Assumptions"], full=True)["matches"][0]["cell"]
            == "G21"
        )
    with Workbook(model, cache) as w:
        assert w.inspect_sheet("Transactions", deep=True)["formula_count"] == 100
        assert w.metrics.counts["index_cache_hits"] == 1
        assert w.metrics.counts.get("cells_scanned", 0) == 0


def test_range_budget_and_cache(model, cache):
    with Workbook(model, cache) as w:
        r = w.read_range("Transactions", "C2:E3")
        assert r["rows"] == [[20, 12, None], [30, 18, None]]
        assert r["formulas"]["E2"]["text"] == "=C2-D2"
        w.read_range("Transactions", "C2:E3")
        assert w.metrics.counts["range_cache_hits"] == 1
        assert w.get_dependencies("Transactions", "E2")[0]["references"] == ["C2", "D2"]
        assert w.get_styles("Transactions", "C2")["styles"]
        with pytest.raises(BudgetExceeded):
            w.read_range("Transactions", "A1:XFD1048576")
    with pytest.raises(BudgetExceeded):
        compact({"long": "x" * 100}, Budget(max_chars=20))
    with pytest.raises(SheetJetError):
        bounds("B2:A1")


def test_stale_source_detection(model, cache):
    with Workbook(model, cache) as w:
        os.utime(model, ns=(model.stat().st_atime_ns, model.stat().st_mtime_ns + 100000))
        with pytest.raises(SheetJetError, match="changed"):
            w.inspect_workbook()


def test_query_aggregation_join_and_budgets(model, cache):
    schema = {"Region": "VARCHAR", "Revenue": "BIGINT"}
    with Workbook(model, cache) as w:
        e = w.query_engine
        staged = e.load("Sales", schema=schema, columns=list(schema))
        assert staged["rows"] == 100
        answer = e.query(
            'SELECT "Region",sum("Revenue") AS revenue FROM Sales GROUP BY 1 ORDER BY 1'
        )
        assert answer["rows"] == [["APAC", 25500], ["EMEA", 26000]]
        assert e.load("Sales", schema=schema, columns=list(schema))["cached"]
        assert e.query("SELECT count(*) FROM Sales a JOIN Sales b ON a.Revenue=b.Revenue")[
            "rows"
        ] == [[100]]
        with pytest.raises(BudgetExceeded):
            e.query("SELECT * FROM Sales", budget=Budget(max_rows=2))
        with pytest.raises(SheetJetError):
            e.query("DROP TABLE Sales")
        import duckdb

        with pytest.raises(duckdb.Error, match="disabled"):
            e.query("SELECT * FROM read_csv('missing.csv')")
        # Staging more data after a restricted query still works.
        e.load("Sales", schema={"Customer": "VARCHAR"}, columns=["Customer"])
        assert e.query("SELECT count(*) FROM Sales")["rows"] == [[100]]


def test_formulas_require_explicit_valid_cache(model, cache):
    with Workbook(model, cache) as w:
        with pytest.raises(SheetJetError, match="Formula inputs"):
            w.query_engine.load("Sales")
        with pytest.raises(SheetJetError, match="Formula inputs"):
            w.query_engine.load("Sales", allow_cached_formulas=True)


def test_read_table_and_profile(model, cache):
    with Workbook(model, cache) as w:
        result = w.read_table(
            "Sales",
            columns=["Region", "Revenue"],
            filters={"Region": "APAC"},
            limit=3,
            schema={"Region": "VARCHAR", "Revenue": "BIGINT"},
        )
        assert len(result["rows"]) == 3
        assert all(row[0] == "APAC" for row in result["rows"])


def test_cli_bounded_output(model, cache, capsys):
    from sheetjet.cli import main

    assert (
        main(
            [
                "--workspace",
                str(model.parent),
                "--cache-dir",
                str(cache),
                "read",
                str(model),
                "Assumptions",
                "A1:B2",
            ]
        )
        == 0
    )
    assert json.loads(capsys.readouterr().out)["rows"][0][1] == 0.08
    assert (
        main(
            [
                "--workspace",
                str(model.parent),
                "--cache-dir",
                str(cache),
                "read",
                str(model),
                "Assumptions",
                "A1:ZZ10000",
            ]
        )
        == 2
    )
    assert json.loads(capsys.readouterr().err)["error"] == "BudgetExceeded"
