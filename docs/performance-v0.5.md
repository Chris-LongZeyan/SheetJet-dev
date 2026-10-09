# v0.5: fixing the wide-sheet loss

The v0.4 skill comparison showed a concrete weakness: SheetJet took 1.323 seconds
on a 2,500-column sheet where the frozen Anthropic submission took 0.754 seconds.
That task selects two columns, filters 40 rows, and writes one total. Treating the
result as inevitable SQL overhead would have missed the main cause.

**The loss is now reversed.** In the fresh five-trial comparison, SheetJet v0.5
completed the wide-sheet task in a median **0.633 s**, versus **0.883 s** for the
frozen Anthropic submission. The slowest v0.5 trial (0.729 s) was faster than the
fastest Anthropic trial (0.808 s). All 60 final executions passed independent
saved-workbook checks.

## Final measurements

Median wall time in seconds, including process startup and writeback:

| Task | Fresh v0.4 | v0.5 | Frozen Anthropic submission |
|---|---:|---:|---:|
| Wide projection + edit | 1.083 | **0.633** | 0.883 |
| Single-cell edit | 0.270 | **0.285** | 0.726 |
| Selected multi-sheet aggregation + edit | 1.254 | **1.240** | 1.925 |
| New budget + chart | 0.516 | 0.559 | **0.465** |

Bold marks the faster current SheetJet/Anthropic workflow. v0.4 is a version-control
baseline, not another peer. The new engine wins all three existing-workbook tasks:
**1.39×**, **2.55×**, and **1.55×** faster respectively. Every v0.5 trial was faster
than every Anthropic trial for those three tasks in this run. Creation still favors
Anthropic's median, with overlapping ranges, and both workflows use openpyxl there.
It is a fallback comparison, not native SheetJet chart creation.

Median sampled peak process-tree RSS, MiB:

| Task | Fresh v0.4 | v0.5 | Frozen Anthropic submission |
|---|---:|---:|---:|
| Wide projection + edit | 106.8 | 54.2 | 47.4 |
| Single-cell edit | 35.6 | 35.8 | 45.1 |
| Selected multi-sheet aggregation + edit | 58.2 | 58.1 | 48.6 |
| New budget + chart | 40.3 | 40.2 | 40.5 |

The wide-sheet change reduces time by **41.5%** and RSS by **49.3%** relative to the
fresh v0.4 baseline. It does not erase the remaining query-memory gap to Anthropic.
Other paths are unchanged; small before/after shifts there are not attributed to
this optimization. These are local, synthetic workload results, not a universal
claim about either skill.

[Complete final trials, ranges, checks, dependencies, and hashes](benchmarks/skill-workflows-v0.5.json).

## What was unnecessarily slow

The earlier optimization removed pandas imports from internal connection settings
and projection writes. It stopped short of the user query. The wide task passes
`["APAC"]` to a parameterized SQL filter; DuckDB's Python value-conversion path
imports its optional dataframe stack for that simple string too. The extra startup
and memory cost was large compared with the useful work in this small projection.

v0.5 uses DuckDB's native [PREPARE, EXECUTE, and DEALLOCATE](https://duckdb.org/docs/current/sql/statements/prepare)
for positional parameter lists containing ordinary text. It prepares the original
bounded query, binds escaped text arguments through EXECUTE, fetches the result,
and retains one prepared query for reuse. It deallocates on an execution failure or
query change, and invalidates the plan when the connection is replaced. It never substitutes
placeholders inside user SQL. Dictionary parameters, mixed types, and strings with
NUL characters retain the original DB-API path.

The optimization is in the engine. **Both SheetJet versions execute the exact same
frozen builder**, including the same parameterized filter. The benchmark has not
been rewritten to use a hard-coded SQL literal or a precomputed answer.

Tests compare text values and coercion behavior against native Python parameter
binding, including quotes, Unicode, backslashes, embedded SQL syntax, positional
reordering, incorrect argument counts, and conversion failures. Existing restricted
file-access tests still apply. A fresh-process check verifies that the text-query
path does not import pandas.

## Controlled rerun

Measured on **9 October 2026**, using the exact fixture bytes and archived builders
from the [v0.4 comparison](skill-workflows-v0.4.md). No new agents generated solutions
for this rerun. The Anthropic implementation remains the competent preservation-aware
submission, with openpyxl streaming reads and targeted XML patches.

Three configurations run serially in randomized order: installed v0.4, candidate
v0.5, and the frozen Anthropic submission. Each executes all four tasks five times,
for 60 fresh-process trials. Final seed: `2026100902`. Data caches and output directories
are fresh; OS caches are not flushed. All correctness judging remains outside the
timer, including values, formulas, existing formula caches, styles, hidden sheets,
custom XML relationships, and new-workbook chart references. Memory is sampled
process-tree RSS at 10 ms intervals, with the same limitations as v0.4.

The v0.4 baseline is the previously built, isolated wheel with SHA-256
`4c13f48195bf99e9ca0d7645156a4d5e98d8d7280d34ff07301dd903734bb2bd`.
Its `query.py` hash is
`c6c5f6f0540b8aa91448606ef0115bbfd557f929af6d8308e12ce5ee99018691`, matching
the v0.4 release measurements. Candidate source hashes and builder hashes are in the
new raw report. Python and library versions are recorded there too.

## Reproduction

Use the [submission binder and fixture instructions](../benchmarks/submissions/README.md).
Install v0.4 and v0.5 into separate environments, bind the same archived SheetJet
builder once to each environment's Python, and bind the Anthropic builder to an
environment with the same openpyxl version. Label those registries `sheetjet_v04`,
`sheetjet_v05`, and `anthropic`:

```console
python -m benchmarks.skill_trials --run benchmark-output/skill-inputs --registry sheetjet_v04=PATH_TO_V04_REGISTRY --registry sheetjet_v05=PATH_TO_V05_REGISTRY --registry anthropic=PATH_TO_ANTHROPIC_REGISTRY --output benchmark-output/v05-results --repeats 5 --seed 2026100902
```

Use the original 25,000 rows per period for these frozen builders. Do not overwrite
an earlier results directory or reuse a data cache. These execution trials do not
measure agent reasoning time or model-token consumption, and they do not establish
that one skill is best at every spreadsheet task.

## Repeated-query check and development evidence

An initial implementation prepared and deallocated the statement on every call.
Its [60-trial workflow run](benchmarks/v0.5-before-plan-reuse.json) passed, but a
separate warm-query probe found a 4.40 ms median per query, versus 3.60 ms for v0.4.
Those were single-process block medians with overlapping ranges, not a formal
regression estimate, but they exposed unnecessary repeated preparation work.
The final implementation retains one plan and rebinds new text values. The complete
peer comparison was rerun after that change; preliminary results were not substituted
for final measurements.

The [warm-query script](../benchmarks/text_queries.py) runs five blocks of 100 queries
after the first query has completed. Startup is reported separately. Run it under
each installed version:

```console
python -m benchmarks.text_queries benchmark-output/skill-inputs/inputs/wide.xlsx benchmark-output/warm-text/report.json
```

Initial probes are retained for review: [v0.4](benchmarks/v0.4-warm-text-initial.json)
and [v0.5 before plan reuse](benchmarks/v0.5-warm-text-before-reuse.json).

The final warm check alternated three processes per version, each with five blocks
of 100 queries. Median process-level block medians were **3.09 ms for v0.4** and
**3.16 ms for v0.5**. Process medians ranged 2.31–3.27 ms and 2.92–3.28 ms respectively.
The ranges overlap; this does not establish a warm-query speedup or statistical
non-inferiority. The demonstrated win is startup and complete wide-sheet execution.
[All six warm-process results](benchmarks/v0.5-warm-text-final.json).

## Release checks

All 91 regression tests, Ruff checks, formatting checks, and the skill validator
passed. The v0.5 wheel was built and installed into an isolated directory; both
examples, a parameterized-query smoke check, and all [10 preservation stress
fixtures](benchmarks/stress-v5.json) passed from that installation. Source hashes
match the final benchmarked engine. Remote GitHub CI and native Excel execution
are not asserted by these local checks.
