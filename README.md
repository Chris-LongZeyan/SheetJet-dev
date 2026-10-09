# SheetJet

**Keep the workbook local. Give the agent the answer.**

[![Tests](https://github.com/Chris-LongZeyan/SheetJet-dev/actions/workflows/ci.yml/badge.svg)](https://github.com/Chris-LongZeyan/SheetJet-dev/actions/workflows/ci.yml)
[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](https://www.python.org/downloads/)
[![MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)

SheetJet is a Python library, CLI, and agent skill for working locally with large
Excel workbooks. Query selected columns, reconcile records across snapshots, and
patch exact cells while preserving untouched workbook parts. No LLM API key,
embeddings service, or Excel installation is required.

```text
1,000,000 transactions → local SQL → 4 regional totals
Two workbook snapshots → business-key matching → complete change file
One assumption edit    → XML patch → untouched workbook parts preserved
```

**New in v0.6:** cross-workbook reconciliation across selected sheets, with composite
keys, exact decimal comparison, duplicate-key checks, reusable projections, and
atomic JSONL exports. [Usage and semantics](docs/reconciliation.md) ·
[Release notes](CHANGELOG.md) · [Roadmap and release gates](docs/roadmap.md)

## Get started

Python 3.11 or newer:

```console
git clone https://github.com/Chris-LongZeyan/SheetJet-dev.git
cd SheetJet-dev
python -m pip install -e .
python -m examples.demo --rows 10000
python -m examples.reconcile --rows 10000
```

These offline examples generate synthetic workbooks, execute the workflow, and check
the result. Outputs stay under `benchmark-output/`. Reconciliation requires a fresh
output directory; use `--output benchmark-output/another-run` when repeating it.
Run `python -m examples.multisheet` for header alignment and cache reuse after an edit.

## Analyze selected sheets

```python
from sheetjet import Workbook

with Workbook("regional.xlsx", cache_dir=".sheetjet-cache") as book:
    book.query_engine.load_sheets(
        "Sales",
        {"January": "A1:H250001", "February": "A1:H250001"},
        {"Region": "VARCHAR", "Revenue": "DECIMAL(18,2)"},
    )
    result = book.query_engine.query('''
        SELECT "Region", SUM("Revenue") AS revenue
        FROM Sales GROUP BY 1 ORDER BY 1
    ''')
    print(book.serialize(result))
```

The first row in each range supplies headers; column order may differ. Select sheets
with compatible units and meanings. Explicit ranges avoid stale worksheet dimensions.
Use one open `Workbook` for related queries. Optional `materialize("Sales")` pays for
one session-local copy when many repeated scans justify it.

## Reconcile workbook snapshots

```python
from sheetjet import reconcile

summary = reconcile(
    "before.xlsx", "after.xlsx",
    before_ranges={"North": "A1:C10001", "South": "A1:C10001"},
    after_ranges={"North": "A1:C10021", "South": "A1:C9991"},
    schema={"ID": "BIGINT", "Region": "VARCHAR", "Amount": "DECIMAL(18,2)"},
    keys=["ID"],
    sample_limit=5,
    output="changes.jsonl",
    cache_dir=".sheetjet-cache",
)
print(summary["counts"])
```

Records match by key across all selected sheets. Moving or reordering an unchanged
record does not create a false change. The summary reports added, removed, changed,
and unchanged counts; the optional JSONL contains **every change**, sorted by key.
Duplicate or null keys fail before export. [Full Python and CLI reference](docs/reconciliation.md).

## Inspect and edit exact cells

```console
sheetjet inspect book.xlsx
sheetjet find book.xlsx "China growth" --sheet Assumptions
sheetjet read book.xlsx Assumptions A1:C8
sheetjet patch book.xlsx patch.json revised.xlsx
```

An example `patch.json`:

```json
[{"operation":"SET_VALUE","sheet":"Assumptions","cell":"B1","expected":0.08,"value":0.12}]
```

The patch changes authorized XML spans, copies untouched compressed ZIP records,
and verifies saved cells and package structure. A full audit accompanies the output.
See the [operation reference](skills/sheetjet/references/operations.md).

## Design strengths

| Capability | Contract |
|---|---|
| Bounded answers | Defaults: 2,000 cells, 100 query rows, 12,000 JSON characters; overflow raises an error |
| Selected-column analytics | Streaming projection into typed Parquet; SQL filters, groups, joins and windows |
| Exact decimals | Numeric XML text reaches DECIMAL conversion without a Python float round trip |
| Multi-sheet reconciliation | Explicit ranges, shared schema, global key checks, null-safe comparison, full change export |
| Reusable local data | Per-sheet projections invalidate by source, range, schema and formula policy |
| Precise edits | Value/formula/existing-style changes and rectangular value writes with preconditions |
| Preservation | Untouched compressed parts and worksheet XML outside edited spans retained |
| Discovery | Package metadata, local label search, formula patterns and lexical precedents |

Caches contain workbook data. Choose an appropriate local cache directory or pass
`persistent_cache=False` / `--no-persistent-cache` for query projections. Search
indexes have separate persistence. Read the [architecture and fidelity contract](docs/architecture.md).

## Agent skill

After installing the package, copy [`skills/sheetjet/`](skills/sheetjet/) to your
agent's skill directory (for Codex, usually `~/.codex/skills/sheetjet/`). The
[skill instructions](skills/sheetjet/SKILL.md) cover discovery, query planning,
reconciliation, ambiguity resolution, bounded disclosure, and audited edits.
The library and CLI also work without an agent.

## Evidence and limits

The [v0.5 frozen skill comparison](docs/performance-v0.5.md) measured faster execution
than the Anthropic submission on three existing-workbook tasks: **1.39× wide
projection, 2.55× single-cell editing, and 1.55× multi-sheet aggregation/writeback**.
All 60 final executions passed independent output checks. Anthropic retained lower
query memory and faster new-workbook creation. Those measurements concern the frozen
submissions and fixtures; they are not a universal skill ranking or a v0.6 rerun.

- [v0.6 reconciliation measurements](docs/performance-v0.6.md): fresh-process cold and cached runs, with complete export verification.
- [v0.4 three-skill evaluation](docs/skill-workflows-v0.4.md): includes the installed OpenAI Spreadsheets skill.
- [v0.3 library benchmarks](docs/performance-v0.3.md): native readers win cold aggregation; Polars-Parquet wins the measured cached aggregation.
- [Original measurements](docs/benchmarks.md) and [v0.2 comparison](docs/peer-benchmarks.md).

SheetJet does **not** calculate formulas, refresh pivots, execute macros, or verify
native Excel rendering. Content edits request recalculation; stored formula caches
may remain stale. Query and reconciliation inputs reject formulas unless cached
values are explicitly accepted, and missing caches still fail.

Structural changes (row/column shifts, sheet renames, table creation/resizing, chart
creation) are unsupported. Shared/array formula groups, table headers/totals, and
merged non-anchor cells are protected. XLS/XLSB, encrypted workbooks, and strict
OOXML are unsupported. Reconciliation compares selected typed values, not styles,
formula text, or whole-workbook fidelity.

## Development

```console
python -m pip install -e ".[test]"
python -m pytest -q
python -m ruff check src tests benchmarks examples
python -m ruff format --check src tests benchmarks examples
python -m benchmarks.reconciliation --rows 500000 --repeats 3
```

Large benchmarks require temporary disk space. Fixture generation and verification
are documented separately from timed execution. For library peer comparisons, install
`.[benchmark]` and run `python -m benchmarks.peers`. See [CONTRIBUTING.md](CONTRIBUTING.md)
for correctness and evidence requirements.

MIT licensed. The original design brief is preserved in
[Excel-skill-instruction.md](Excel-skill-instruction.md).
