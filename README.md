# SheetJet

**Keep the workbook local. Give the agent the answer.**

[![Tests](https://github.com/Chris-LongZeyan/SheetJet-dev/actions/workflows/ci.yml/badge.svg)](https://github.com/Chris-LongZeyan/SheetJet-dev/actions/workflows/ci.yml)
[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](https://www.python.org/downloads/)
[![MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)

SheetJet is a Python engine, CLI, and reusable agent skill for large Excel workbooks.
Search locally, query selected columns with SQL, and return a few useful answers.
Edit exact cells without round-tripping the rest of the workbook through a cell-object
library. No LLM API key, embeddings service, or Excel installation is required.

```text
1,000,000 transactions → local SQL → 4 regional totals
1 assumption edit     → XML patch → untouched workbook parts preserved
```

**New in v0.4:** project from 2,500-column sheets without exposing their headers or
raising the response budget, and avoid optional dataframe imports for queries
without user parameters. Includes independent same-model trials using the official
Anthropic XLSX and installed OpenAI Spreadsheets skills: **72 successful executions
across two rounds**, with frozen builders, independent saved-file checks, and raw results.

In the final round, SheetJet's single-cell edit was **3.53× faster than the Anthropic
submission and 7.41× faster than the OpenAI submission**. Selected multi-sheet
aggregation plus writeback was **1.53× and 2.70× faster**. These are three-trial
medians for specific generated workflows; Anthropic won the small wide-sheet task
and used less query memory. [Full skill comparison and reproduction](docs/skill-workflows-v0.4.md).

In the current four-sheet, 100k-row comparison, changing one assumption took a median
**0.105 s and 38.5 MiB peak RSS**, versus **7.247 s and 258.8 MiB** for openpyxl's
load/save path. Both use three synthetic trials and edit a small assumptions sheet.
At one million rows, cache reopening plus aggregation fell from **0.311 s to 0.124 s**
against a fresh v0.2 rerun. Native readers still win cold aggregation.
[Full results, raw trials and limits](docs/performance-v0.3.md).

## The useful difference

**Context limits are enforced in code.** Responses default to at most 2,000 cells,
100 query rows, and 12,000 JSON characters. An oversized answer fails with guidance
to narrow or aggregate it. It is never silently truncated.

**Edits have a preservation contract.** SheetJet changes authorized XML spans and
copies untouched compressed ZIP records exactly. That includes opaque chart, pivot,
VBA and extension parts. A saved output comes with a machine-readable change log,
changed-cell checks and package structure checks. Only changed parts are recompressed.

**Unchanged sheets stay reusable.** A per-sheet Parquet cache retains only selected
columns, keyed by their source, range, schema and formula policy. Change one sheet,
and other sheets' projections remain reusable—even in a renamed output workbook.
Numeric XML values reach DECIMAL conversion without passing through a Python float.

**Performance claims come with a script.** Compare cold metadata/lookup/aggregation,
warm SQL, memory use and edits against openpyxl, pandas with two reader engines,
Calamine, and Polars/Calamine. A Polars-Parquet baseline also tests cached queries.
See [current peer measurements](docs/performance-v0.3.md), the
[v0.2 peer comparison](docs/peer-benchmarks.md), and the
[original v0.1 measurements](docs/benchmarks.md).
Ordinary deterministic analysis also avoids an LLM context dump; SheetJet packages
that approach with enforceable budgets, discovery tools and audited edits.
The [skill workflow benchmark](docs/skill-workflows-v0.4.md) additionally evaluates
independent agent-authored solutions, including preservation-aware peer fallbacks.

## Try it in a minute

Python 3.11 or newer:

```console
git clone https://github.com/Chris-LongZeyan/SheetJet-dev.git
cd SheetJet-dev
python -m pip install -e ".[test]"
python -m examples.demo --rows 10000
```

The offline demo creates a transaction workbook, finds an assumption, computes
regional totals, changes one assumption and verifies the saved output. Generated
files stay under `benchmark-output/demo/`. Use `--rows 1000000` for the large version.

Run `python -m examples.multisheet` to see reordered headers align across sheets,
exact decimal totals, optional materialization, and both period caches reused after
an assumption edit. [Read the runnable example](examples/multisheet.py).

```python
from sheetjet import Workbook

with Workbook("book.xlsx") as book:
    print(book.inspect_workbook())
    print(book.find_text("China growth", sheets=["Assumptions"]))

    book.query_engine.load(
        "Sales",
        columns=["Region", "Revenue"],
        schema={"Region": "VARCHAR", "Revenue": "DECIMAL(18,2)"},
    )
    answer = book.query_engine.query('''
        SELECT "Region", SUM("Revenue") AS revenue
        FROM Sales GROUP BY 1 ORDER BY revenue DESC LIMIT 10
    ''')
    print(book.serialize(answer))

    report = book.patch_cells([
        {"operation": "SET_VALUE", "sheet": "Assumptions", "cell": "B1",
         "expected": 0.08, "value": 0.12}
    ], output="revised.xlsx")
    print(report)
```

The example expects a table named `Sales` with those columns. For an unnamed table,
pass `sheet=` and `ref=` to `query_engine.load`. Its first row supplies the headers.

For a series of queries that repeatedly scan the same data, call
`book.query_engine.materialize("Sales")` once after loading. The copy is local to
that session and consumes temporary storage; cache loading stays lazy by default.
Use the same open `Workbook` for related queries to reuse its restricted connection.

## CLI and agent skill

For multiple sheets:

```python
with Workbook("regional.xlsx", cache_dir=".sheetjet-cache") as book:
    book.query_engine.load_sheets(
        "Sales",
        {"January": "A1:H250001", "February": "A1:H250001"},
        {"Region": "VARCHAR", "Revenue": "DECIMAL(18,2)"},
    )
    print(book.query_engine.query('''
        SELECT _sheet, "Region", SUM("Revenue") AS revenue
        FROM Sales GROUP BY 1, 2 ORDER BY 1, 2
    '''))
```

Column order may differ; headers must identify the same concepts and units. Ranges
are explicit because worksheet dimensions can be stale. Persistent caches are local
data files; opt out with `persistent_cache=False` or `--no-persistent-cache`.

```console
sheetjet inspect book.xlsx
sheetjet find book.xlsx "China growth" --sheet Assumptions
sheetjet read book.xlsx Assumptions A1:C8
sheetjet patterns book.xlsx Model --limit 10
sheetjet patch book.xlsx patch.json revised.xlsx
sheetjet --metrics inspect revised.xlsx
```

Copy [`skills/sheetjet/`](skills/sheetjet/) into your agent's skill directory after
installing the Python package. For Codex, the usual location is
`~/.codex/skills/sheetjet/`. The [skill instructions](skills/sheetjet/SKILL.md) guide
progressive discovery, ambiguity resolution, query planning, and bounded disclosure.
The Python engine and CLI can also be used without an agent.

## What works today

| Capability | Current behavior |
|---|---|
| XLSX / XLSM | Streaming inspection and targeted editing |
| Workbook map | Sheet states, tables, names, panes, calculation settings and feature counts |
| Local search | Persistent SQLite FTS5 index; scoped label search or explicit full-text scan |
| Large-table analysis | DuckDB filters, groups, joins, windows, profiles and exact decimal schemas |
| Multiple sheets | Explicit selected ranges, header alignment and source-sheet provenance |
| Persistent projections | Typed Parquet, per-sheet invalidation, corruption recovery and opt-out |
| Formula inspection | Formula locations, repeated vertical patterns and lexical precedents |
| Cell edits | Set values; set/copy formulas and existing styles; write rectangular value ranges |
| Preservation | Untouched compressed ZIP records and worksheet XML outside edited spans retained |
| Validation | Expected-value preconditions, saved-cell checks, structure checks and full audit |
| Tests and stress data | Ten fixture families, an offline demo and process-isolated benchmarks |

**Know the boundary:** SheetJet does not calculate Excel formulas, refresh pivots,
execute macros, or verify native Excel rendering. Content edits request recalculation;
dependent stored caches may remain stale until Excel recalculates. SQL refuses
formula inputs unless cached values are explicitly accepted, and missing formula
caches still fail.

Dependency-sensitive structural changes—row/column shifts, sheet renames, table
creation/resizing and chart creation—are explicitly unsupported in this release.
Shared/array formula groups, table headers/totals and merged non-anchor cells are
protected. XLS/XLSB, encrypted workbooks and strict OOXML are also unsupported.
Read the [full architecture and fidelity contract](docs/architecture.md).

## Reproduce the evidence

```console
python -m pytest -q
python -m pip install -e ".[benchmark]"
python -m benchmarks.peers --rows 25000 --sheets 4 --repeats 3
python -m benchmarks.peers --rows 250000 --sheets 4 --repeats 1 --shared-strings --output benchmark-output/million-peers/results.json
python -m benchmarks.fixtures benchmark-output/suite --rows 10000
python -m benchmarks.run --rows 100000 --repeats 3 --output benchmark-output/100k/results.json
python -m benchmarks.run --rows 1000000 --repeats 3 --output benchmark-output/million/results.json
```

Large runs take several minutes and need temporary disk space. Benchmarks verify
aggregation answers against independent fixture arithmetic. Results include raw
trials, versions, file hashes, peak RSS and clearly labeled context-token estimates.
The VBA fixture is an opaque-byte canary, **not a functioning macro project**; real
Excel-authored fixtures and native verification are welcome.

## Help improve it

Useful contributions are small workbooks that expose a bug, fidelity regressions,
and reproducible performance improvements. See [CONTRIBUTING.md](CONTRIBUTING.md).
The next milestones are native Excel fidelity tests, table-aware structural edits,
and native parsing with the same value/error/formula semantics. SheetJet does not
claim to beat every library in every workload: specialized native readers currently
win cold tabular reads, and Polars-Parquet wins the measured cached aggregation.

MIT licensed. The original design brief is preserved in
[`Excel-skill-instruction.md`](Excel-skill-instruction.md).
