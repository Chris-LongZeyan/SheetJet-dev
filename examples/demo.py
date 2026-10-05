"""Offline, reproducible end-to-end demo. Run from the repository root."""

import argparse
import json
from pathlib import Path

from benchmarks.fixtures import transactions
from sheetjet import Workbook


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, default=10000)
    parser.add_argument("--output", type=Path, default=Path("benchmark-output/demo"))
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    path = transactions(args.output / "transactions.xlsx", args.rows)
    with Workbook(path, args.output / "cache") as book:
        print("Inspect:", book.serialize(book.inspect_workbook()))
        print("Find:", book.serialize(book.find_text("China growth", ["Assumptions"])))
        book.query_engine.load(
            "Transactions",
            columns=["Region", "Revenue"],
            schema={"Region": "VARCHAR", "Revenue": "BIGINT"},
        )
        answer = book.query_engine.query(
            'SELECT "Region",SUM("Revenue") AS revenue FROM Transactions GROUP BY 1 ORDER BY 1'
        )
        print("Answer:", book.serialize(answer))
        patch = [
            {
                "operation": "SET_VALUE",
                "sheet": "Assumptions",
                "cell": "B1",
                "expected": 0.08,
                "value": 0.12,
            }
        ]
        print(
            "Edit:",
            book.serialize(book.patch_cells(patch, args.output / "revised.xlsx", overwrite=True)),
        )
        print("Metrics:", json.dumps(book.metrics.snapshot()))


if __name__ == "__main__":
    main()
