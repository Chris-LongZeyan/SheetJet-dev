"""Generate two snapshots and reconcile them without depending on row positions."""

import argparse
import json
from pathlib import Path

from openpyxl import Workbook as Excel

from sheetjet import reconcile


def fixtures(directory, rows=10000):
    if rows < 1:
        raise ValueError("rows must be positive")
    directory.mkdir(parents=True, exist_ok=True)
    paths = [directory / "before.xlsx", directory / "after.xlsx"]
    if any(path.exists() for path in paths):
        raise FileExistsError("Choose a fresh directory for generated reconciliation fixtures")
    all_ranges = []
    added = max(1, rows // 100)
    for revision, path in enumerate(paths):
        book = Excel(write_only=True)
        sheets = [book.create_sheet(f"Partition{i + 1}") for i in range(4)]
        headers = [["ID", "Region", "Amount"], ["Amount", "ID", "Region"]]
        counts = [0] * 4
        for i, sheet in enumerate(sheets):
            sheet.append(headers[(i + revision) % 2])
        ids = range(1, rows + 1) if revision == 0 else range(rows + added, 0, -1)
        for key in ids:
            if revision and key <= rows and key % 97 == 0:
                continue
            partition = (key + revision) % 4
            cents = key * 37 + (5 if revision and key <= rows and key % 101 == 0 else 0)
            values = {
                "ID": key,
                "Region": ["APAC", "EMEA", "Americas", "Other"][key % 4],
                "Amount": f"{cents // 100}.{cents % 100:02d}",
            }
            sheets[partition].append(
                [values[column] for column in headers[(partition + revision) % 2]]
            )
            counts[partition] += 1
        notes = book.create_sheet("Excluded")
        notes.append(["Not part of reconciliation", "=1/0"])
        notes.sheet_state = "hidden"
        book.save(path)
        all_ranges.append(
            {s.title: f"A1:C{count + 1}" for s, count in zip(sheets, counts, strict=True)}
        )
    spec = {
        "before_ranges": all_ranges[0],
        "after_ranges": all_ranges[1],
        "schema": {"ID": "BIGINT", "Region": "VARCHAR", "Amount": "DECIMAL(18,2)"},
        "keys": ["ID"],
    }
    (directory / "spec.json").write_text(json.dumps(spec, indent=2), encoding="utf-8")
    expected = {"added": added, "removed": rows // 97, "changed": rows // 101 - rows // (97 * 101)}
    expected["unchanged"] = rows - expected["removed"] - expected["changed"]
    return paths, spec, expected


def run(directory, rows):
    paths, spec, expected = fixtures(directory, rows)
    result = reconcile(
        *paths,
        **spec,
        sample_limit=3,
        output=directory / "changes.jsonl",
        cache_dir=directory / "cache",
    )
    assert result["counts"] == expected
    (directory / "summary.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=Path("benchmark-output/reconciliation"))
    parser.add_argument("--rows", type=int, default=10000)
    args = parser.parse_args()
    run(args.output, args.rows)
