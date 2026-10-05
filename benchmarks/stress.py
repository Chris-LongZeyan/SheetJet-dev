"""Exercise discovery, computation and fidelity across all ten fixture families."""

import argparse
import json
import time
from pathlib import Path
from zipfile import ZipFile

from sheetjet import Workbook

from .fixtures import suite


def run(directory, rows):
    manifest = suite(directory / "fixtures", rows)
    reports = []
    for entry in manifest:
        path = directory / "fixtures" / entry["file"]
        start = time.perf_counter()
        with Workbook(path, directory / "cache") as w:
            meta = w.inspect_workbook()
            kind = entry["kind"]
            if kind == "transactions":
                e = w.query_engine
                e.load(
                    "Transactions",
                    columns=["Region", "Revenue"],
                    schema={"Region": "VARCHAR", "Revenue": "BIGINT"},
                )
                result = e.query(
                    'SELECT "Region",SUM("Revenue") FROM Transactions GROUP BY 1 ORDER BY 1'
                )
                patch = {
                    "operation": "SET_VALUE",
                    "sheet": "Assumptions",
                    "cell": "B1",
                    "value": 0.12,
                }
            elif kind in {"financial-model", "formula-heavy"}:
                result = w.formula_patterns("Model")
                assert any(p["count"] > 1 for p in result)
                patch = {
                    "operation": "SET_FORMULA",
                    "sheet": "Model",
                    "cell": "C3",
                    "formula": "=B3*1.10",
                }
            elif kind == "formatting-heavy":
                result = w.get_styles("Model", "A2:C3")
                patch = {
                    "operation": "COPY_STYLE",
                    "sheet": "Model",
                    "source_cell": "B2",
                    "cell": "B3",
                }
            elif kind == "hidden-names":
                result = w.find_names("Input")
                assert len(result) == 10
                patch = {"operation": "SET_VALUE", "sheet": "Hidden0", "cell": "B2", "value": 3}
            elif kind == "merged":
                result = w.inspect_sheet("Model", deep=True)
                assert result["merged_ranges_count"] > 0
                patch = {
                    "operation": "SET_VALUE",
                    "sheet": "Model",
                    "cell": "A3",
                    "value": "Revised section",
                }
            elif kind == "messy":
                e = w.query_engine
                e.load("Sales", schema={"Customer": "VARCHAR", "Sales": "BIGINT"})
                e.load("Costs", schema={"Customer": "VARCHAR", "Costs": "BIGINT"})
                result = e.query(
                    'SELECT SUM(s.amount-c.amount) AS margin FROM (SELECT "Customer",SUM("Sales") amount FROM Sales GROUP BY 1) s JOIN (SELECT "Customer",SUM("Costs") amount FROM Costs GROUP BY 1) c USING ("Customer")'
                )
                assert result["rows"] == [[sum(range(3, 100)) * 4]]
                patch = {"operation": "SET_VALUE", "sheet": "Model", "cell": "B3", "value": 30}
            elif kind == "interconnected":
                result = w.get_dependencies("Schedule49", "B2")
                assert result[0]["references"] == ["'Schedule48'!B2"]
                patch = {"operation": "SET_VALUE", "sheet": "Model", "cell": "B2", "value": 200}
            else:
                result = meta["features"]
                patch = {"operation": "SET_VALUE", "sheet": "Model", "cell": "B2", "value": 200}
            output = directory / "edited" / path.name
            report = w.patch_cells([patch], output, overwrite=True)
            with ZipFile(path) as a, ZipFile(output) as b:
                untouched = [name for name in a.namelist() if name not in report["changed_parts"]]
                assert all(a.read(name) == b.read(name) for name in untouched)
            reports.append(
                {
                    "fixture": entry,
                    "seconds": time.perf_counter() - start,
                    "response_characters": len(w.serialize(result)),
                    "untouched_parts_byte_equal": len(untouched),
                    "validation": report["validation"],
                    "metrics": w.metrics.snapshot(),
                }
            )
    result = {
        "rows_parameter": rows,
        "fixture_count": len(reports),
        "passed": True,
        "results": reports,
        "scope": "Generated OOXML and opaque-byte preservation; no native Excel rendering, pivot refresh or VBA execution",
    }
    directory.mkdir(parents=True, exist_ok=True)
    (directory / "stress-results.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                "passed": True,
                "fixtures": len(reports),
                "output": str(directory / "stress-results.json"),
            }
        )
    )


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--output", type=Path, default=Path("benchmark-output/stress"))
    p.add_argument("--rows", type=int, default=10000)
    a = p.parse_args()
    run(a.output, a.rows)
