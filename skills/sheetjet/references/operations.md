# Execution reference

## Table analytics

Keep the workbook open and reuse a staged table. Supply explicit SQL types to avoid
sample-based coercion of identifiers and mixed columns. Without a schema, all
selected columns are VARCHAR. Excel dates stay as serials; numeric financial amounts
can use `DECIMAL(18,2)` rather than floating point.

```python
from sheetjet import Workbook

with Workbook("book.xlsx", cache_dir=".sheetjet-cache") as book:
    print(book.find_table("Sales"))
    book.query_engine.load(
        "Sales", columns=["Region", "Revenue"],
        schema={"Region": "VARCHAR", "Revenue": "DECIMAL(18,2)"},
    )
    result = book.query_engine.query('''
        SELECT "Region", SUM("Revenue") AS revenue
        FROM Sales GROUP BY 1 ORDER BY revenue DESC LIMIT 10
    ''')
    print(book.serialize(result))
```

For an unnamed region, pass an alias, sheet and exact range to `load`. Its first row
must contain unique, nonblank string headers. Load a second table to use SQL joins,
window functions or reconciliation. Identifiers are double quoted and values use
`?` parameters. Query execution accepts one SELECT and disables external file/network
access in its read-only DuckDB connection. This is a local query interface, not a
resource sandbox for arbitrary hostile SQL. Keep queries task-directed.

CLI schemas are JSON files mapping table names to selected columns and SQL types:

```json
{"Sales":{"Region":"VARCHAR","Revenue":"DECIMAL(18,2)"}}
```

```console
sheetjet query book.xlsx 'SELECT "Region", SUM("Revenue") FROM Sales GROUP BY 1' --schema schema.json
```

Shell quoting varies; use the Python API for complex SQL on Windows.

## Multiple sheets and persistent projections

```python
with Workbook("regional.xlsx", cache_dir=".sheetjet-cache") as book:
    book.query_engine.load_sheets(
        "Sales", {"January": "A1:H250001", "February": "A1:H250001"},
        {"Region": "VARCHAR", "Revenue": "DECIMAL(18,2)"},
    )
    print(book.query_engine.query('''
        SELECT _sheet, "Region", SUM("Revenue") AS revenue
        FROM Sales GROUP BY 1, 2 ORDER BY 1, 2
    '''))
```

Headers may be in a different order on different sheets. Explicit source ranges
avoid trusting stale worksheet dimensions. Persistent projections contain only the
selected columns; numerical XML values reach DuckDB without a Python float round
trip, preserving the source's precision for DECIMAL schemas. Subsequent sessions
reuse Parquet rather than scanning unchanged data. Cache keys include formula policy,
so accepting cached formulas in one request cannot bypass the default refusal later.
`persistent_cache=False` disables Parquet use and creation for a load. Choose a
temporary cache directory if data must not persist. Cache files have no automatic
retention policy; remove obsolete files when no session uses them.

CLI: `sheetjet query-sheets book.xlsx "SELECT _sheet, count(*) FROM Sales GROUP BY 1" --ranges ranges.json --schema columns.json`.
`ranges.json` maps sheet names to ranges, and `columns.json` maps column names to SQL
types. `--name` changes the SQL relation name and `--no-persistent-cache` opts out.

## Patches

```json
[
  {"operation":"SET_VALUE","sheet":"Assumptions","cell":"B1","expected":0.08,"value":0.12},
  {"operation":"COPY_FORMULA","sheet":"Model","source_cell":"D4","cell":"E4"},
  {"operation":"COPY_STYLE","sheet":"Model","source_cell":"D4","cell":"E4"}
]
```

```console
sheetjet patch book.xlsx patch.json revised.xlsx
sheetjet read revised.xlsx Assumptions A1:B2
sheetjet validate revised.xlsx Model D4:E4
```

Copy sources refer to the original workbook snapshot, not earlier operations in the
batch. One content operation and one style operation per target cell are allowed.
The source is never overwritten. Existing output/audit paths require `--overwrite`.
The audit contains old/new values and must be treated as workbook-sensitive data.

## Formula and style evidence

```console
sheetjet patterns book.xlsx Model --limit 10
sheetjet formulas book.xlsx Model D4:F8
sheetjet dependencies book.xlsx Model D4:F8
sheetjet styles book.xlsx Model D4:F8
sheetjet inspect book.xlsx --sheet Model --deep --section merged_ranges
```

Patterns compact repeated vertical formulas. Dependencies are lexical precedents,
not a complete Excel dependency graph: defined names, structured references and
external references remain symbolic; INDIRECT/OFFSET report incomplete dependencies.
Shared formula followers may have empty stored formula text with a shared ID; use
patterns for translated examples, or inspect the anchor. Style definitions are
returned once per ID. `validate` reports stored errors and missing caches, and never
claims that formulas were recalculated.

## Performance and limits

An XLSX worksheet is a compressed XML stream. A late-row read still scans preceding
XML; small returned ranges do not imply constant-time random access. First full-text
indexing is linear. Targeted edits spool touched worksheets to disk and stream-copy
the ZIP; they avoid loading every worksheet as Python cell objects. Temporary disk
space must accommodate expanded touched sheets, the output workbook, staged tables,
and SQLite/Parquet caches. In v0.2, untouched ZIP local records are copied compressed
without inflate/deflate. Their compressed bytes, data descriptors and member metadata
survive; central-directory offsets are relocated. Touched sheets are still spooled,
parsed, edited and recompressed. Prefixed/SFX, multi-disk and unusual directory layouts
are rejected instead of silently rewritten. Copied members are not fully decompressed
for a pre-existing corruption audit; input health and native Excel compatibility are
separate from unchanged-byte preservation.

Use `--metrics` for timings, scanned/decoded cells, range/cache counts and serializer
size estimates. Estimated tokens use characters divided by four, not a tokenizer or
actual API billing. Do not claim universal speed superiority from synthetic results.
