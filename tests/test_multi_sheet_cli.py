import json

from openpyxl import Workbook as Excel

from sheetjet.cli import main


def test_query_sheets_cli(tmp_path, capsys):
    path = tmp_path / "book.xlsx"
    b = Excel()
    b.active.append(["Region", "Amount"])
    b.active.append(["APAC", 10])
    b.save(path)
    ranges = tmp_path / "ranges.json"
    schema = tmp_path / "columns.json"
    ranges.write_text(json.dumps({"Sheet": "A1:B2"}))
    schema.write_text(json.dumps({"Region": "VARCHAR", "Amount": "BIGINT"}))
    assert (
        main(
            [
                "--workspace",
                str(tmp_path),
                "--cache-dir",
                str(tmp_path / "cache"),
                "query-sheets",
                str(path),
                "SELECT _sheet,sum(Amount) FROM Sales GROUP BY 1",
                "--ranges",
                str(ranges),
                "--schema",
                str(schema),
                "--no-persistent-cache",
            ]
        )
        == 0
    )
    assert json.loads(capsys.readouterr().out)["rows"] == [["Sheet", 10]]
    assert not list((tmp_path / "cache").glob("*.parquet"))
