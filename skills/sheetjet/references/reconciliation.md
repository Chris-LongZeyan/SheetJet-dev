# Reconciliation reference

Requires SheetJet 0.6+. Match records by business key across selected sheet ranges
in two workbooks. Inspect names and small header samples first; do not assume that
equal headers establish compatible units or meanings. Exact ranges include one
header row and exclude unrelated notes, subtotals, and blank trailing records.

```python
from sheetjet import reconcile

summary = reconcile(
    "before.xlsx", "after.xlsx",
    before_ranges={"North": "A1:C10001", "South": "A1:C10001"},
    after_ranges={"North": "A1:C10021", "South": "A1:C9991"},
    schema={"ID": "BIGINT", "Region": "VARCHAR", "Amount": "DECIMAL(18,2)"},
    keys=["ID"],
    compare_columns=["Region", "Amount"],
    sample_limit=5,
    output="changes.jsonl",
    cache_dir=".sheetjet-cache",
)
print(summary)
```

For the CLI, put `before_ranges`, `after_ranges`, `schema`, `keys`, and optional
`compare_columns` in `spec.json`:

```console
sheetjet --cache-dir .sheetjet-cache reconcile before.xlsx after.xlsx --spec spec.json --sample-limit 5 --output changes.jsonl
```

Keys must be unique and non-null across all selected sheets on each side. Use a
composite key such as `["ID", "Region"]` when that reflects record identity.
Changing a key appears as one removal and one addition. Sheet names do not enter
the key; sheet movement alone is unchanged. Header and row order may differ.

Comparison is exact after type conversion, case-sensitive for text, and null-safe.
Choose `VARCHAR` for identifiers with significant leading zeros; choose a DECIMAL
scale fine enough to retain meaningful differences. Numeric Excel dates remain
serials. No fuzzy matching, tolerance, or implicit date/locale conversion applies.
Only key and comparison columns are projected; omitted schema columns are ignored.
Without `compare_columns`, all non-key schema columns are compared. An empty list
compares record membership only.

The summary has `counts` (added/removed/changed/unchanged), `rows`, per-column
`changed_columns` counts, `sample`, `sample_truncated`, `cache_hits`, and scope.
`sample_limit=0` requests counts without examples. Default sample limit is 20;
normal response budgets still apply. Budget overflow raises before export. Reduce
the sample before raising budgets. Never treat the sample as the full change set.

The optional JSONL exports every changed record sorted by typed business key. Each
has `kind`, `key`, `before`, `after`, and `changed_columns`. Present sides contain
`sheet` and `values`; additions have no before side and removals have no after side.
DECIMAL values are strings, preserving precision. Zero changes produce an empty file.
Publication is atomic; existing files require `overwrite=True` / `--overwrite`.
Neither input can be used as the export destination.

Selected formulas fail unless `allow_cached_formulas=True` / `--cached-formulas`
is appropriate and their caches exist. This never recalculates or establishes
freshness. Repeated runs reuse local per-sheet projections; opt out with
`persistent_cache=False` / `--no-persistent-cache`. Shared-string data can still
persist in the cache directory. Both caches and exports contain source-derived data.

Do not describe this as a complete workbook diff: styles, formula text, layout,
charts, and omitted sheets are not compared. It does not emit cell coordinates,
apply updates, or merge workbooks. A later edit requires a separate plan identifying
exact cells and suitable preconditions.
