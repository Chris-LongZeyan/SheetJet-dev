# Reconcile workbook snapshots

Available in v0.6. Compare business records across two XLSX/XLSM files, even when
headers and rows reorder or records move between selected sheets. Both inputs stay
unchanged. Use this for period snapshots, ledger extracts, inventory versions, and
other datasets with a stable business key.

## Run the example

```console
python -m examples.reconcile --rows 10000 --output benchmark-output/reconciliation
```

The example generates four selected sheets per workbook, changes their header order,
reverses record order, moves records between sheets, and introduces known additions,
removals, and amount changes. An unrelated hidden sheet is excluded. The script checks
the summary and saves both `summary.json` and complete `changes.jsonl`.
Choose a fresh output directory for another run.

## CLI

Create `spec.json` using the exact ranges and column names in your files:

```json
{
  "before_ranges": {"North": "A1:C10001", "South": "A1:C10001"},
  "after_ranges": {"North": "A1:C10021", "South": "A1:C9991"},
  "schema": {"ID": "BIGINT", "Region": "VARCHAR", "Amount": "DECIMAL(18,2)"},
  "keys": ["ID"],
  "compare_columns": ["Region", "Amount"]
}
```

```console
sheetjet --cache-dir .sheetjet-cache --metrics reconcile before.xlsx after.xlsx --spec spec.json --sample-limit 5 --output changes.jsonl
```

Standard output contains one bounded JSON summary. `--metrics` writes elapsed time
and projection-cache hits to standard error. Omit `--output` for a summary alone.
`--sample-limit 0` returns counts without example records. `--overwrite` explicitly
permits replacing an existing export; either input path is always protected.
From 0.6.1, all CLI paths must lie inside `--workspace` (current directory by default).
Place the global option before `reconcile` to choose a different trusted root.

## Python

```python
import json
from pathlib import Path
from sheetjet import reconcile

spec = json.loads(Path("spec.json").read_text(encoding="utf-8"))
summary = reconcile(
    "before.xlsx", "after.xlsx", **spec,
    sample_limit=5, output="changes.jsonl", cache_dir=".sheetjet-cache",
)
print(summary["counts"])
```

For reusable execution settings, 0.6.1 also provides an immutable options object:

```python
from sheetjet import ReconcileOptions, reconcile

options = ReconcileOptions(sample_limit=5, cache_dir=".sheetjet-cache")
summary = reconcile("before.xlsx", "after.xlsx", **spec, options=options)
```

The execution arguments below remain accepted individually for v0.6 compatibility.
Do not combine individual execution settings with `options`; ambiguous settings fail.

| Argument | Meaning |
|---|---|
| `before_ranges`, `after_ranges` | Nonempty sheet-name → rectangular A1-range dictionaries; first row contains headers |
| `schema` | Header-name → SQL-type dictionary; choose compatible meanings and units on both sides |
| `keys` | Nonempty list of unique column names; composite keys are supported |
| `compare_columns` | Optional list; defaults to all schema columns except keys; `[]` compares membership only |
| `sample_limit` | Default 20; from zero through `budget.max_rows` |
| `output`, `overwrite` | Optional complete JSONL destination; replacement defaults to false |
| `cache_dir`, `persistent_cache` | Local cache location and projection persistence; persistence defaults to true |
| `allow_cached_formulas` | Defaults to false; true accepts stored caches without recalculation |
| `budget` | Optional `sheetjet.core.Budget`; bounds summary rows, cells and JSON characters |

Supported scalar types are `VARCHAR`, `BIGINT`, `DOUBLE`, `BOOLEAN`, `DATE`,
`TIMESTAMP`, and `DECIMAL(precision,scale)` within DuckDB's supported bounds. Excel
numeric dates remain serial values; a date style does not convert them to dates.
Choose a shared representation explicitly. Use `VARCHAR` for identifiers with
significant leading zeros and an appropriate DECIMAL scale for financial amounts.
Comparison happens **after conversion**: a scale of two decimal places can erase
smaller differences. Invalid typed values fail rather than being silently dropped.

## Matching and output contract

Keys must be non-null and unique across **all selected sheets on each side**. A
duplicate aborts the operation; choose a composite key when an ID repeats by region,
period, or another business dimension. Sheet placement is provenance, not identity.
If sheet identity matters, include a business column representing it in the key.
Blank rows included in a range fail the non-null key check; choose exact endpoints.

Matched records use SQL `IS DISTINCT FROM` for every selected comparison column:
null equals null, and null differs from zero or an empty string. Matching is exact,
case-sensitive for text, and has no fuzzy matching or numeric tolerance. Reordering
rows, moving unchanged records across sheets, and changing ignored columns do not
count as changes. Only keys and comparison columns are projected.

The summary includes:

- Input paths, selected ranges, effective schema, keys, and comparison columns.
- Input record counts and added, removed, changed, and unchanged counts.
- Per-column counts of differences among matched changed records.
- A deterministic sample sorted by typed business key, with `sample_truncated`.
- The export path, formula policy, and per-input projection-cache hit counts.

Each JSONL line has this structure:

```json
{"kind":"changed","key":{"ID":101},"before":{"sheet":"North","values":{"Region":"APAC","Amount":"37.37"}},"after":{"sheet":"South","values":{"Region":"APAC","Amount":"37.42"}},"changed_columns":["Amount"]}
```

An addition has `before: null`; a removal has `after: null`. Both have an empty
`changed_columns` list. DECIMAL values are JSON strings to preserve precision;
DATE/TIMESTAMP values are strings too. Non-finite floating-point values cannot be
serialized. The export contains all changed records, independent of sample size,
and is empty when there are no changes. Unchanged records are never exported.

Summary budget overflow raises an error before publication. Reduce `sample_limit`,
narrow the comparison, or explicitly supply an appropriate larger budget. The
complete export stays local and is not constrained by the model response budget.

## Execution, caching and publication

Each selected sheet becomes a typed Parquet projection. A disk-backed DuckDB join
matches records, checks keys, and materializes only the differences. Exports fetch
512 rows at a time. Each query session uses DuckDB's 512 MB memory setting; this is
not a total process memory cap. Large comparisons need temporary disk space for
staging, joins, sorting, and the complete export. The two workbook sessions coexist.

Repeated runs reuse unchanged sheet projections, including after an unrelated sheet
is edited. Cache identity includes selected columns, schema, range, and formula
policy. Projection reuse does not prove formula-cache freshness. Set
`persistent_cache=False` / `--no-persistent-cache` to avoid persistent projections;
the shared-string store may still persist. Choose a temporary cache directory when
all local cache data should be removed after use.

The export is written and flushed to a temporary sibling before publication. A
failed write leaves the previous destination intact. Publication without overwrite
also refuses a destination created by another writer during processing. Source and
projection changes detected during the operation fail closed. Concurrent source
writers are unsupported; filesystem publication guarantees depend on the filesystem.

This is record reconciliation, not a workbook merge or an edit plan. It does not
compare formulas as text, styles, charts, layout, or omitted sheets; produce cell
coordinates; apply changes; or recalculate formulas. Accept stored formula caches
only when their provenance fits the task; missing caches still fail.

See [measurements and reproduction](performance-v0.6.md) and the
[architecture contract](architecture.md).
