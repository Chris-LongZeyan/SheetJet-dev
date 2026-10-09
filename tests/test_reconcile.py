import hashlib
import json

import pytest
from openpyxl import Workbook as Excel

from sheetjet import BudgetExceeded, SheetJetError, reconcile
from sheetjet.cli import main
from sheetjet.core import Budget


def write_book(path, sheets):
    book = Excel()
    book.remove(book.active)
    for name, rows in sheets.items():
        sheet = book.create_sheet(name)
        for row in rows:
            sheet.append(row)
    book.save(path)
    return path


@pytest.fixture
def pair(tmp_path):
    before = write_book(
        tmp_path / "before.xlsx",
        {
            "North'Q": [
                ["ID", "Amount", "Label"],
                [1, "1.10", "old"],
                [2, "2.20", "same"],
                [3, None, None],
            ],
            "South": [["Label", "Amount", "ID"], ["removed", "3.30", 4]],
        },
    )
    after = write_book(
        tmp_path / "after.xlsx",
        {
            "North'Q": [["Label", "ID", "Amount"], ["same", 2, "2.20"], ["new", 1, "1.11"]],
            "Moved": [["ID", "Amount", "Label"], [8, "8.88", "added"], [3, None, None]],
        },
    )
    spec = {
        "before_ranges": {"North'Q": "A1:C4", "South": "A1:C2"},
        "after_ranges": {"North'Q": "A1:C3", "Moved": "A1:C3"},
        "schema": {"ID": "BIGINT", "Amount": "DECIMAL(20,4)", "Label": "VARCHAR"},
        "keys": ["ID"],
    }
    return before, after, spec


def test_reconcile_reordered_multisheet_records_and_complete_export(pair, tmp_path):
    before, after, spec = pair
    hashes = [hashlib.sha256(p.read_bytes()).hexdigest() for p in (before, after)]
    output = tmp_path / "changes.jsonl"
    result = reconcile(
        before, after, **spec, sample_limit=2, output=output, cache_dir=tmp_path / "cache"
    )
    assert result["counts"] == {"added": 1, "removed": 1, "changed": 1, "unchanged": 2}
    assert result["changed_columns"] == {"Amount": 1, "Label": 1}
    assert result["rows"] == {"before": 4, "after": 4}
    assert result["sample_truncated"]
    records = [json.loads(line) for line in output.read_text(encoding="utf-8").splitlines()]
    assert len(records) == 3
    assert result["sample"] == records[:2]
    assert [record["key"]["ID"] for record in records] == [1, 4, 8]
    assert records[0]["before"]["values"]["Amount"] == "1.1000"
    assert records[0]["after"]["values"]["Amount"] == "1.1100"
    assert records[1]["after"] is None and records[2]["before"] is None
    assert hashes == [hashlib.sha256(p.read_bytes()).hexdigest() for p in (before, after)]
    again = reconcile(before, after, **spec, sample_limit=0, cache_dir=tmp_path / "cache")
    assert again["cache_hits"] == {"before": 2, "after": 2}
    assert again["sample"] == [] and again["counts"] == result["counts"]


@pytest.mark.parametrize("rows,message", [([[1, 2], [1, 3]], "duplicate"), ([[None, 2]], "null")])
def test_invalid_business_keys_fail_without_publishing(tmp_path, rows, message):
    source = write_book(tmp_path / "in.xlsx", {"Data": [["ID", "Value"], *rows]})
    output = tmp_path / "out.jsonl"
    with pytest.raises(SheetJetError, match=message):
        reconcile(
            source,
            source,
            before_ranges={"Data": f"A1:B{len(rows) + 1}"},
            after_ranges={"Data": f"A1:B{len(rows) + 1}"},
            schema={"ID": "BIGINT", "Value": "BIGINT"},
            keys=["ID"],
            output=output,
            cache_dir=tmp_path / "cache",
        )
    assert not output.exists()


def test_composite_keys_null_changes_and_ignored_formulas(tmp_path):
    before = write_book(
        tmp_path / "b.xlsx",
        {
            "Data": [
                ["ID", "Region", "Amount", "Ignored"],
                [1, "A", None, "=1/0"],
                [1, "B", "0.10000000000000000001", "=1/0"],
            ]
        },
    )
    after = write_book(
        tmp_path / "a.xlsx",
        {
            "Data": [
                ["ID", "Region", "Amount", "Ignored"],
                [1, "B", "0.10000000000000000002", "=2/0"],
                [1, "A", "0", "=2/0"],
            ]
        },
    )
    result = reconcile(
        before,
        after,
        before_ranges={"Data": "A1:D3"},
        after_ranges={"Data": "A1:D3"},
        schema={
            "ID": "BIGINT",
            "Region": "VARCHAR",
            "Amount": "DECIMAL(38,20)",
            "Ignored": "DOUBLE",
        },
        keys=["ID", "Region"],
        compare_columns=["Amount"],
        cache_dir=tmp_path / "cache",
    )
    assert result["counts"]["changed"] == 2
    assert result["sample"][1]["before"]["values"]["Amount"] == "0.10000000000000000001"
    assert result["sample"][1]["after"]["values"]["Amount"] == "0.10000000000000000002"
    assert result["sample"][0]["before"]["values"]["Amount"] is None
    assert result["changed_columns"] == {"Amount": 2}


def test_budget_failure_does_not_publish_and_existing_output_is_protected(pair, tmp_path):
    before, after, spec = pair
    output = tmp_path / "changes.jsonl"
    with pytest.raises(BudgetExceeded):
        reconcile(
            before,
            after,
            **spec,
            budget=Budget(max_chars=100),
            output=output,
            cache_dir=tmp_path / "cache",
        )
    assert not output.exists()
    output.write_text("keep")
    with pytest.raises(SheetJetError, match="exists"):
        reconcile(before, after, **spec, output=output)
    assert output.read_text() == "keep"
    with pytest.raises(SheetJetError, match="input"):
        reconcile(before, after, **spec, output=before, overwrite=True)
    reconcile(before, after, **spec, output=output, overwrite=True, cache_dir=tmp_path / "cache")
    assert len(output.read_text().splitlines()) == 3


def test_empty_key_only_tables_and_cache_opt_out(tmp_path):
    before = write_book(tmp_path / "b.xlsx", {"Data": [["ID"]]})
    after = write_book(tmp_path / "a.xlsx", {"Data": [["ID"], [2], [1]]})
    result = reconcile(
        before,
        after,
        before_ranges={"Data": "A1:A1"},
        after_ranges={"Data": "A1:A3"},
        schema={"ID": "BIGINT"},
        keys=["ID"],
        persistent_cache=False,
        cache_dir=tmp_path / "cache",
    )
    assert result["counts"] == {"added": 2, "removed": 0, "changed": 0, "unchanged": 0}
    assert [r["key"] for r in result["sample"]] == [{"ID": 1}, {"ID": 2}]
    assert not list((tmp_path / "cache").glob("projection-*.parquet"))


def test_reconcile_cli(pair, tmp_path, capsys):
    before, after, spec = pair
    path = tmp_path / "spec.json"
    path.write_text(json.dumps(spec))
    assert (
        main(
            [
                "--cache-dir",
                str(tmp_path / "cache"),
                "reconcile",
                str(before),
                str(after),
                "--spec",
                str(path),
                "--sample-limit",
                "0",
            ]
        )
        == 0
    )
    result = json.loads(capsys.readouterr().out)
    assert result["counts"]["changed"] == 1 and result["sample"] == []


def test_selected_formula_policy_and_missing_cache(tmp_path):
    source = write_book(tmp_path / "formula.xlsx", {"Data": [["ID", "Value"], [1, "=1+1"]]})
    spec = {
        "before_ranges": {"Data": "A1:B2"},
        "after_ranges": {"Data": "A1:B2"},
        "schema": {"ID": "BIGINT", "Value": "BIGINT"},
        "keys": ["ID"],
    }
    for accept_cache in [False, True]:
        with pytest.raises(SheetJetError, match="Formula inputs"):
            reconcile(
                source,
                source,
                **spec,
                allow_cached_formulas=accept_cache,
                cache_dir=tmp_path / "cache",
            )


def test_valid_formula_cache_requires_opt_in_and_is_disclosed(tmp_path):
    from zipfile import ZipFile

    base = write_book(tmp_path / "base.xlsx", {"Data": [["ID", "Value"], [1, "=1+1"]]})
    source = tmp_path / "cached.xlsx"
    with ZipFile(base) as original, ZipFile(source, "w") as target:
        for info in original.infolist():
            content = original.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                assert b"<f>1+1</f><v></v>" in content
                content = content.replace(b"<f>1+1</f><v></v>", b"<f>1+1</f><v>2</v>")
            target.writestr(info, content)
    spec = {
        "before_ranges": {"Data": "A1:B2"},
        "after_ranges": {"Data": "A1:B2"},
        "schema": {"ID": "BIGINT", "Value": "BIGINT"},
        "keys": ["ID"],
        "cache_dir": tmp_path / "cache",
    }
    output = tmp_path / "unchanged.jsonl"
    result = reconcile(source, source, **spec, allow_cached_formulas=True, output=output)
    assert result["counts"] == {"added": 0, "removed": 0, "changed": 0, "unchanged": 1}
    assert result["formula_policy"] == "stored caches accepted; not recalculated"
    assert output.read_bytes() == b""
    with pytest.raises(SheetJetError, match="Formula inputs"):
        reconcile(source, source, **spec)


def test_export_race_keeps_other_writers_file_and_cleans_temporary(pair, tmp_path, monkeypatch):
    import importlib

    module = importlib.import_module("sheetjet.reconcile")
    operation = "rename" if module.os.name == "nt" else "link"
    original = getattr(module.os, operation)
    output = tmp_path / "race.jsonl"

    def race(source, destination):
        output.write_text("other writer")
        return original(source, destination)

    monkeypatch.setattr(module.os, operation, race)
    before, after, spec = pair
    with pytest.raises(FileExistsError):
        reconcile(before, after, **spec, output=output, cache_dir=tmp_path / "cache")
    assert output.read_text() == "other writer"
    assert not list(tmp_path.glob(".sheetjet-reconcile-*"))


def test_duplicate_keys_across_sheets_are_rejected(pair, tmp_path):
    before, after, spec = pair
    duplicate = write_book(
        tmp_path / "duplicate.xlsx",
        {
            "A": [["ID", "Amount", "Label"], [1, 10, "one"]],
            "B": [["ID", "Amount", "Label"], [1, 20, "two"]],
        },
    )
    spec["before_ranges"] = {"A": "A1:C2", "B": "A1:C2"}
    with pytest.raises(SheetJetError, match="duplicate"):
        reconcile(duplicate, after, **spec, cache_dir=tmp_path / "cache")


def test_quoted_columns_and_source_name_collision(tmp_path):
    columns = ['key"id', "_sheetjet_source"]
    before = write_book(tmp_path / "b.xlsx", {"Data": [columns, [1, "old"]]})
    after = write_book(tmp_path / "a.xlsx", {"Data": [columns, [1, "new"]]})
    result = reconcile(
        before,
        after,
        before_ranges={"Data": "A1:B2"},
        after_ranges={"Data": "A1:B2"},
        schema={columns[0]: "BIGINT", columns[1]: "VARCHAR"},
        keys=[columns[0]],
        cache_dir=tmp_path / "cache",
    )
    assert result["counts"]["changed"] == 1
    assert result["sample"][0]["changed_columns"] == [columns[1]]


def test_failed_export_preserves_existing_output(pair, tmp_path, monkeypatch):
    import importlib

    module = importlib.import_module("sheetjet.reconcile")
    original = module._record

    def fail_on_second_record(row, keys, columns):
        if row[0] == "removed":
            raise OSError("simulated export failure")
        return original(row, keys, columns)

    monkeypatch.setattr(module, "_record", fail_on_second_record)
    output = tmp_path / "keep.jsonl"
    output.write_text("previous export")
    before, after, spec = pair
    with pytest.raises(OSError, match="simulated"):
        reconcile(
            before,
            after,
            **spec,
            sample_limit=0,
            output=output,
            overwrite=True,
            cache_dir=tmp_path / "cache",
        )
    assert output.read_text() == "previous export"
    assert not list(tmp_path.glob(".sheetjet-reconcile-*"))


@pytest.mark.parametrize(
    "override", [{"keys": "ID"}, {"compare_columns": "Label"}, {"schema": {1: "BIGINT"}}]
)
def test_invalid_column_specification_is_rejected(pair, override):
    before, after, spec = pair
    with pytest.raises(SheetJetError):
        reconcile(before, after, **(spec | override))
