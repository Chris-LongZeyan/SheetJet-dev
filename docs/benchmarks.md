# Reproducible benchmarks

These are **archived v0.1 measurements**. The engine now copies compressed ZIP
members and reuses per-sheet Parquet projections. See the
[v0.2 comparison](peer-benchmarks.md) for current multi-sheet results against
openpyxl, pandas, Calamine and Polars. Old timings describe the implementation
measured then; rerunning the commands with v0.2 does not reproduce the old code path.

The useful claim is bounded context and preservation with measured performance.
These measurements do not establish superiority over every spreadsheet tool or
every workload. The naïve-context baseline is deliberately different from ordinary
deterministic spreadsheet analysis; the openpyxl streaming baseline shows that
distinction explicitly.

## 100,000 transaction rows

Windows 11, Python 3.13.7, DuckDB 1.5.5, openpyxl 3.1.5, lxml 6.1.1. Three isolated
processes per method; values below are medians. Five columns, plus an assumptions
sheet. [Raw results, environment and fixture hash](benchmarks/100k.json).

| Task | SheetJet | openpyxl read-only | Observation |
|---|---:|---:|---|
| Initial metadata | 0.090 s | 0.318 s | Both avoid full cell-object loading; returned metadata differs |
| Last-row value | 1.289 s | 5.873 s | Same value verified; roughly 4.6× faster in this run |
| Regional totals, cold | 4.380 s | 6.038 s | Same four totals verified; includes SheetJet staging |
| Repeated regional totals | 0.083 s | Not measured | SheetJet table already staged; not a cold comparison |

Cold aggregation peak RSS was **114.2 MiB for SheetJet and 53.2 MiB for openpyxl**.
DuckDB has a fixed overhead and SheetJet did not win memory use on this case. Last-row
lookup used 30.6 MiB versus 53.3 MiB. Peak RSS was sampled every 10 ms and may miss
short spikes; it is total process RSS, not the incremental cost of the task.

The naïve workflow took 6.758 s and peaked at 82.1 MiB while materializing all rows
and their JSON. Its raw data payload was **3,770,196 characters**; the four-row answer
payload was **78 characters**. Both SheetJet and ordinary deterministic openpyxl
aggregation produce that small answer. With a deliberately rough characters/4
estimate, those payloads are 942,549 versus 20 tokens. They are not measured model
token counts, API bills, or complete agent conversations. Query/inspection envelopes,
instructions and schemas are excluded from those payload figures.

Changing one assumption in a separate small sheet of this workbook took 1.074 s,
including ZIP copying and validation, at 49.8 MiB peak RSS. This does not measure
editing a cell inside the large transaction sheet. Uncompressed bytes of ten
untouched package parts were verified by ZIP CRC and size; regression and stress
tests additionally compare untouched payloads byte for byte.

## 1,000,000 transaction rows

The same environment and fixture schema, with **5,000,007 populated cells** across
the transaction and assumptions sheets. This is **one trial per method**, a scale
check rather than a statistically stable performance estimate.
[Raw results and fixture hash](benchmarks/million.json).

| Task | SheetJet | openpyxl read-only |
|---|---:|---:|
| Initial metadata | 0.096 s | 0.317 s |
| Last-row value | 15.156 s | 66.681 s |
| Regional totals, cold | 42.265 s | 78.134 s |
| Regional totals, already staged | 0.105 s | Not measured |

Aggregation peak RSS was 144.2 MiB for SheetJet versus 129.0 MiB for openpyxl streaming.
The naïve materialized-context workflow peaked at 422.6 MiB. One assumption edit
in the separate small sheet took 7.311 s, including copying the large workbook and
validating the output, with 50.1 MiB peak RSS. This illustrates the unavoidable ZIP
copy cost even when only one small sheet changes.

Every aggregation and lookup answer matched the independently computed fixture
answer. Native Excel calculation, UI rendering and macro execution were not tested.
Run three or more repetitions on your own hardware before making comparative claims.

The four-row aggregation answer payload was 82 characters versus 38,701,897
characters for the materialized transaction rows. That excludes envelopes and
instructions; the same small answer is available through deterministic openpyxl
aggregation. SheetJet's additional contract is to enforce a bounded return value
instead of relying only on prompting discipline.

## Methodology

```console
python -m pip install -e ".[test]"
python -m benchmarks.run --rows 100000 --repeats 3 --output benchmark-output/100k/results.json
python -m benchmarks.run --rows 1000000 --repeats 3 --output benchmark-output/million/results.json
python -m benchmarks.stress --rows 10000
```

Each timed measurement starts a fresh Python process. The OS file cache is not
flushed. Timings include relevant imports, workbook open, staging, execution and
cleanup, except the explicitly labeled warm query. Warm-query memory still includes
staging. Benchmark fixtures are synthetic, predominantly numeric with repeated
inline strings. Shared-string-heavy, compressed, network-hosted, formula-heavy or
styling-heavy production files can behave differently.

The openpyxl baseline uses `read_only=True, data_only=True`. Aggregate baselines use
the same row-level inputs and group arithmetic, then compare with an independent
calculation from the fixture definition. Naïve context materialization is simulated
locally; no LLM is called. An Excel formula engine is never part of these tests.

The full fixture suite covers transaction data, a financial model, repeated formulas,
dense formatting, hidden sheets/names, merged regions, charts/pivot definitions,
multiple tables per sheet, 50 connected sheets, and opaque VBA bytes. The macro
fixture is a canary and cannot prove executable VBA compatibility. Generated pivot
definitions load in openpyxl, but native Excel refresh/rendering has not been tested.

The [recorded stress run](benchmarks/stress.json) passed all ten fixture families
with a 1,000-row generator parameter. The stress runner exercises discovery, formula patterns, formatting changes,
table reconciliation, dependency inspection, cell edits and byte preservation across
all ten families. Structural operations such as adding columns or creating tables
and charts are unsupported in the editing API; fixture generation is not evidence
that those operations are supported.
