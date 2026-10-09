# Roadmap and release gates

SheetJet versions describe shipped capabilities. v0.6 delivers cross-workbook
reconciliation. The milestones below are planned work, not available features or
promised dates. Major versions require evidence of broader capability and a clear
migration contract; version numbers alone are not a measure of quality.

| Milestone | Intended capability | Gate before release |
|---|---|---|
| **1.0 — Stable core** | Stable discovery, analytics, reconciliation, and cell-patch APIs | Documented compatibility policy and error categories; clean Windows/Linux CI on supported Python versions; installed-package tests; public reproducible benchmarks; a legally shareable Excel-authored fidelity corpus with native opening/recalculation evidence |
| **2.0 — Structural editing** | Controlled table resizing and row/column changes | Reference rewriting across formulas, names, tables, charts, and links; transactional preflight and change plans; fail-closed handling of unsupported dependencies; native Excel round-trip tests and explicit migration notes |
| **2.1 — Batch reconciliation** | Declarative comparisons across folders, schema-drift diagnostics, resumable runs | Stable run manifests, input fingerprints, interruption/restart tests, duplicate handling across files, deterministic complete outputs, and incremental-work benchmarks that include preparation cost |
| **3.0 — Interchangeable execution engines** | Optional native parsing/query backends for faster cold processing | Differential tests preserving numeric lexemes, nulls, dates, errors, shared strings, and formula policies; reproducible gains across diverse workloads; bounded resource behavior; portable fallback and packaging |

## Current release: v0.6

- Cross-workbook matching by unique or composite business keys across selected sheets.
- Header alignment, exact DECIMAL comparison, null-safe matching, and duplicate checks.
- Bounded summaries plus atomic complete JSONL change exports.
- Reusable selected-column projections, runnable examples, and independently checked benchmarks.

## How priorities are chosen

Prefer a demonstrated user problem with a small reproducible workbook over a broad
feature count. Preservation failures and incorrect answers take priority over speed.
Performance work must include cold and cached phases, memory, independent output
checks, and strong peer implementations. Agent evaluations must keep the model,
task, available tools, and correctness rubric comparable.

The current benchmark evidence identifies remaining gaps: specialized native
readers win cold tabular reads, the measured Polars-Parquet baseline wins cached
aggregation, and the frozen Anthropic submission wins new-workbook creation and
query memory on tested tasks. Future releases should target those gaps without
weakening validation or changing the workload to favor SheetJet.

Formula calculation is not implied by any milestone. Supporting a calculation
engine would require a separate compatibility and verification design. Likewise,
preserving opaque workbook parts does not establish native visual fidelity.

Contributions should state the intended milestone and its measurable acceptance
criteria. See [CONTRIBUTING.md](../CONTRIBUTING.md) and [release notes](../CHANGELOG.md).
