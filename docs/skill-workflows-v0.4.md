# v0.4: independent skill workflow trials

This comparison uses **agent-authored workflows**, rather than equating each skill
with a hand-written library wrapper. Three isolated agents inherited the same model
and received the same task manifest, fixtures, and one assigned skill. They inspected
and repaired their own saved outputs before freezing their builders. No agent saw
competitors' implementations or performance results. The parent corrected one
SheetJet data-cache isolation issue before measurement.

Sources are the repository's SheetJet skill, the official Anthropic XLSX skill at
commit `8a1541c4a3ffa5a20a5a91de0dcf3f0bab1d1ef4`, and the official installed OpenAI
Spreadsheets skill bundle `26.921.10847`. These are **same-model, different-skill**
trials; they do not compare OpenAI models against Anthropic models.
[Source hashes, frozen builders, and reproduction instructions](../benchmarks/submissions/README.md).

All **36 final executions passed**, as did all 36 in the initial round. SheetJet led
these peers on preserving single-cell edits and selected multi-sheet aggregation.
Anthropic's submission won the small wide-sheet task and used less query memory.
There is no evidence of superiority in every aspect.

## Final results

Median wall time in seconds; three fresh processes per task and skill:

| Workflow | SheetJet | Anthropic XLSX submission | OpenAI Spreadsheets submission |
|---|---:|---:|---:|
| Single-cell edit | **0.230** | 0.814 | 1.707 |
| Selected multi-sheet aggregation + edit | **1.038** | 1.590 | 2.801 |
| Wide projection + edit | 1.323 | **0.754** | 1.777 |
| New budget + chart | 0.496 | **0.451** | 1.171 |

For these frozen workflows, SheetJet's median edit time was **3.53× faster** than
Anthropic's and **7.41× faster** than OpenAI's. Its multi-sheet task was **1.53×** and
**2.70× faster**, respectively. All three edit trials were faster than all peer edit
trials; the same held for the multi-sheet task. Creation used an openpyxl fallback
for SheetJet and has overlapping timing ranges with Anthropic; it is not evidence
of an engine capability advantage.

Median sampled peak process-tree RSS in MiB:

| Workflow | SheetJet | Anthropic XLSX submission | OpenAI Spreadsheets submission |
|---|---:|---:|---:|
| Single-cell edit | **35.0** | 45.2 | 245.7 |
| Multi-sheet aggregation + edit | 57.6 | **48.9** | 233.5 |
| Wide projection + edit | 106.9 | **47.5** | 258.6 |
| New budget + chart | 40.4 | 40.4 | 273.3 |

The wide task still triggers DuckDB's optional dataframe imports through its user
query parameter. Its 40-row input also pays SQL staging costs that the peer's simple
streaming reduction avoids. This is a real tradeoff, not an omitted losing case.

[Every final trial, check, hash, and peak](benchmarks/skill-workflows-v0.4.json).

## Workloads and independent checks

| Task | Input and required output |
|---|---|
| Edit | Four 25,000-row transaction sheets, assumptions, summary, and hidden notes. Change only Assumptions!B1 from 0.08 to 0.12. |
| Multi-sheet aggregation | Same workbook. Sum Revenue by Region using only Period01 and Period03: 50,000 selected records. Write four static totals to Summary. |
| Wide projection | 2,500 columns and 40 data rows. Sum Amount for APAC using the first and last columns; write one static total. |
| Create | New three-month budget, profit and total formulas, number formats, frozen header, readable widths, and a column chart excluding the total row. |

Existing-workbook tasks require preserving all unrequested values, formula text,
cell styles, sheet order, dimensions, hidden states, and custom XML metadata. Formula
recalculation is explicitly forbidden for existing inputs; cached formula freshness
is not assumed. New-workbook formula caches are optional. Every response must be a
compact JSON file of at most 6,000 bytes. Inputs are hashed before and after each trial.

The independent evaluator reopens saved files using openpyxl, compares cell/style
fingerprints, checks requested values against independent fixture arithmetic, and
checks custom XML bytes and relationships. Creation checks inspect actual formulas,
formatting, freeze panes, and native chart series references in the saved package.
An additional audit preserved Assumptions!C1's original cached value of 8 in every
edit/aggregate output in both rounds; it was added to the judge after timing.
Judging runs outside the timing. Package checks do not establish native Excel visual
fidelity, VBA execution, or calculation correctness for arbitrary formulas.

## Execution and interpretation

Each final cell in the results table represents three fresh subprocess executions
of one frozen agent submission. All methods run serially in a fixed randomized order
with new output directories and fresh SheetJet data caches. Timings include imports,
file reads/writes, and each builder's own validation. OS file caches are not flushed.
The agents' authoring and rendering QA are outside the measured execution.

Measured on 5 October 2026, Windows 11 build 26200, evaluator Python 3.13.7,
DuckDB 1.5.5, openpyxl 3.1.5, lxml 6.1.1, and artifact-tool 2.8.71 from the installed
runtime. RSS is the sampled sum across the builder process tree, every 10 ms;
it can miss brief peaks and double-count shared pages. Each method's runtime and
fallbacks are part of its measured workflow.

These trials measure execution of **one independently authored solution per skill**.
Three timing repeats are not three independent model generations. They do not
establish general agent success rates, skill breadth, actual model-token savings,
or a universal ranking. All submissions keep workbook contents out of their compact
results; SheetJet's enforced API budget is a product contract, not evidence that peer
agents must dump worksheets.

## Implementation changes prompted by the evaluation

Wide headers previously passed through the public response budget before column
projection. v0.4 resolves headers locally, so a 2,500-column input can produce a
two-cell answer under a two-cell budget. Numeric, duplicate, and blank header
validation remains in place.

Profiling the first complete matrix found optional pandas loading through internal
DuckDB parameter conversion. v0.4 uses connection configuration and escaped internal
file-path literals for engine setup and projection I/O. This avoids importing pandas
for queries without user parameters. User SQL parameters retain normal DuckDB
binding. Regression checks cover quoted cache paths, exact file allowlists, locked
configuration, unrelated-file denial, and the absence of the optional import in a
fresh process. The setup follows DuckDB's [connection configuration API](https://duckdb.org/docs/current/clients/python/overview.html).

The first full matrix is retained in [initial raw results](benchmarks/skill-workflows-initial.json).
The final matrix reruns **every peer**, with the same builders and input bytes, after
the engine change. Before/after rounds are separate batches and host conditions can
change; their timing differences must not be attributed entirely to the optimization.
In particular, all peers ran faster in the second round. The observed SheetJet
multi-sheet memory median fell from 115.3 to 57.6 MiB after avoiding the optional
import, but the wall-time comparison above uses only the contemporaneous final round.

## Scope of the comparison

The Anthropic agent used openpyxl for streaming reads and original targeted XML
updates, preserving opaque metadata. The OpenAI agent used artifact-tool authoring
plus an original package-preservation adapter for capabilities absent from its
documented authoring API. Its renderer wrote previews but exited with status 1
without diagnostics during separate QA; frozen timed builders omit rendering.
The SheetJet agent used openpyxl to create the new workbook, since SheetJet does not
create charts or workbook structures. A create-task timing is therefore a fallback
workflow result, not a SheetJet engine advantage.

This fixture set does not replace the [million-row reader/cache measurements](performance-v0.3.md).
Native Calamine/Polars readers remain stronger choices for some cold analytics.
The current skill trials focus on complete preservation-aware tasks, not just parsing.

## Release checks

- 72 regression tests passed; Ruff checks and formatting passed.
- The skill passed its frontmatter/reference scaffold validator.
- The v0.4 wheel was built and installed into an isolated directory; both runnable
  demos passed from that installation.
- All [10 preservation stress fixtures](benchmarks/stress-v4.json) passed from the
  installed wheel (100-row setting, including opaque-feature canaries).
- Relocated Python submissions passed a registry smoke check; the OpenAI runtime
  binding passed JavaScript syntax checking. The original OpenAI builder passed
  every measured execution in its installed runtime.

These are local checks. Remote GitHub CI and native Excel execution are not asserted.
