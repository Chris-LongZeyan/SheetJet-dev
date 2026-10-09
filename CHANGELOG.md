# Changelog

## 0.6.1

- Confine all CLI paths to an explicit workspace (current directory by default),
  rejecting traversal, existing symlink escapes, Windows device names, and alternate
  streams. **CLI migration:** select `--workspace` when files live elsewhere.
- Validate reconciliation spec fields before execution. Apply the same workspace
  checks to the repeated-text-query benchmark and protect its existing outputs.
- Split reconciliation, CLI setup, fixture generation, and benchmark judging into
  focused functions. Add `ReconcileOptions`; existing Python keyword settings remain
  supported and cannot be mixed with an options object.
- Add traversal and compatibility regression tests, branch coverage in CI, and
  SonarQube automatic-analysis scope with an explicit Python version.
- Preserve four historical benchmark submissions byte-for-byte and exclude only
  those archived implementations from analysis. Maintained harnesses remain scanned.

## 0.6.0

- Add `sheetjet.reconcile` and the `sheetjet reconcile` CLI for matching records
  across two workbooks and multiple selected sheets.
- Support composite business keys, exact typed comparison, duplicate/null-key
  rejection, deterministic samples, per-column change counts, and complete JSONL.
- Protect input files and existing exports; publish exports atomically after
  validation and clean up incomplete writes.
- Reuse per-sheet projections; expose formula policy and cache-hit counts.
- Add a runnable reconciliation example, fresh-process benchmark with independent
  export checks, regression tests, and CI example coverage.
- Reorganize the README around workflows; document current limits and future
  release gates. Existing query and patch APIs remain available.

## 0.5.0

- Avoid optional dataframe imports for positional text query parameters using
  native prepared statements, reusing one plan per query session.
- Publish repeated frozen-submission skill measurements and warm-query checks,
  including workloads and memory measurements where peers still win.

## 0.4.0

- Project requested columns from wide worksheets without exposing all headers
  through the response budget.
- Add isolated skill workflow evaluation, frozen submissions, and independent
  saved-workbook validation.

Earlier architecture and measurements are documented in the
[v0.3 report](docs/performance-v0.3.md), [v0.2 comparison](docs/peer-benchmarks.md),
and [original benchmark report](docs/benchmarks.md).
