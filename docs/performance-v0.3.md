# v0.3: fewer copies, reusable query sessions

Measured on 5 October 2026, Windows 11, Python 3.13.7. All tables below use three
fresh subprocess trials per method. Inputs are the **same fixture bytes** used for
v0.2: four selected transaction sheets plus an assumptions sheet and excluded hidden
notes. The 100k fixture uses inline strings; the million-row fixture uses shared strings.

The strongest new result is faster cache reopening. Cold reading improves more
modestly against a fresh v0.2 rerun, and native readers remain faster. SheetJet's
preservation-focused edits remain its largest measured advantage over openpyxl.

## Fresh v0.2 versus v0.3, one million rows

The v0.2 baseline is the wheel built from commit `5e221ef`. Both versions ran the
current benchmark runner against the same workbook, with the same dependencies.
Each version ran as a separate sequential batch; versions were not interleaved and
OS caches were not flushed. This is a local engineering comparison, not a claim of
hardware-independent superiority.

| Workload | v0.2 median, s | v0.3 median, s | v0.2 trial range, s | v0.3 trial range, s |
|---|---:|---:|---:|---:|
| Cold aggregation | 14.309 | 12.563 | 12.396–14.515 | 12.372–12.778 |
| Reopen cache + aggregate | 0.311 | 0.124 | 0.231–0.326 | 0.112–0.131 |

The cold median is 12.2% lower, but trial ranges overlap. The cache-reopen
median is 60.2% lower, with no overlap in these trials. The fresh v0.2 cold
median is substantially lower than its older single-trial result of 21.638 s; use
the fresh rerun for before/after claims. Benchmark noise and run conditions matter.

[Fresh v0.2 raw trials](benchmarks/v0.2-repeat-million.json) ·
[v0.3 million-row trials](benchmarks/v0.3-million.json).

## Current peer results

| Workload / engine | 100k median, s | Million median, s | Million median peak RSS, MiB |
|---|---:|---:|---:|
| Cold SheetJet | 1.758 | 12.563 | 121.6 |
| Cold Polars/Calamine | 0.612 | 1.902 | 177.8 |
| Cold python-calamine | 0.268 | 2.614 | 163.4 |
| Reopen SheetJet cache | 0.091 | 0.124 | 119.8 |
| Polars-Parquet query | 0.007 | 0.010 | 170.0 |

Reopen/cache rows exclude preparation, and both workflows disclose its cost:

- SheetJet: 1.672 s at 100k; 12.572 s at one million rows.
- Polars: 0.691 s at 100k; 2.123 s at one million rows.

RSS includes setup and validation, so it is not the incremental memory of a reopened
session. The million-row rerun focuses on native-reader and cache competitors;
openpyxl and pandas are included in the complete 100k matrix below. Their earlier
million-row results remain explicitly archived in the [v0.2 report](peer-benchmarks.md).

| Other 100k workload | Median seconds | Median peak RSS, MiB |
|---|---:|---:|
| openpyxl read-only | 2.863 | 47.9 |
| pandas/openpyxl | 3.807 | 122.2 |
| pandas/Calamine | 0.824 | 128.5 |
| SheetJet assumption edit | 0.105 | 38.5 |
| openpyxl load/save assumption edit | 7.247 | 258.8 |

The edit targets the small `Assumptions!B1` sheet; it does not measure editing a cell
inside a large transaction sheet. SheetJet includes its audit and validation work.
This comparison does not establish native Excel rendering or functioning VBA.

## Repeated queries and the cost of materialization

Each subprocess performs one warm-up and then 20 aggregations with all answers
independently checked. The table reports the median of 60 individual query latencies
per method (three trials × 20), **not cold latency**. Raw `seconds` records contain
the whole 20-query phase. No answer cache is used.

| Query storage / engine | 100k latency, ms | Million latency, ms | Million preparation, s | Million peak RSS, MiB |
|---|---:|---:|---:|---:|
| SheetJet, lazy Parquet views | 7.583 | 10.900 | 12.086 | 121.5 |
| Polars, lazy Parquet scans | 2.544 | 4.629 | 1.950 | 196.5 |
| SheetJet, session table | 2.081 | 3.292 | 13.009 | 145.4 |
| Polars, materialized DataFrame | 1.442 | 3.149 | 2.142 | 177.8 |

`book.query_engine.materialize("Sales")` pays for a temporary table when repeated
scans warrant it. Preparation includes Excel ingestion, projection creation, optional
materialization and one warm-up query. The table costs more temporary storage and
memory; it is not automatically applied to every load. Polars also receives a
materialized baseline so the comparison does not credit SheetJet merely for copying
data into a more convenient representation.

## What changed and what is checked

- Expat projects selected scalar cells without allocating worksheet trees. Numeric
  text remains exact until typed SQL conversion; formula/error policies are retained.
- CSV conversion writes typed Parquet directly. Cache reopening publishes lazy views
  instead of importing every row into another database.
- A read-only SQL connection remains open across related queries. Only loaded
  projection files are allowlisted; configuration is locked before user SQL executes.
- Optional materialization works for single-sheet projections and multi-sheet unions.
  Failed reloads retain published data, and edits to unrelated sheets preserve reuse.
- Search-index initialization is deferred until a search/index operation needs it.

66 regression tests pass locally, including randomized sparse UTF-8/UTF-16 decoder
comparisons, exact decimals, namespaces, phonetic text, duplicate/mismatched coordinates,
formula/error policy, ZIP64, cache invalidation, reload atomicity and SQL restrictions.
All [10 synthetic stress fixture families](benchmarks/stress-v3.json) pass. These
checks are not native Excel fidelity certification or coverage of every real workbook.

## Reproduce

```console
python -m pip install -e ".[test,benchmark]"
python -m pytest -q
python -m examples.multisheet
python -m benchmarks.peers --rows 25000 --sheets 4 --repeats 3 --output benchmark-output/v03-100k/results.json
python -m benchmarks.peers --rows 250000 --sheets 4 --repeats 3 --shared-strings --methods sheetjet polars_calamine calamine sheetjet_cached polars_parquet sheetjet_session polars_session sheetjet_materialized polars_materialized --output benchmark-output/v03-million/results.json
```

Use `--workbook` to reuse identical generated inputs instead of regenerating them.
To repeat the baseline, install commit `5e221ef` into a separate directory using
`pip install --no-deps --target`, put that directory first on `PYTHONPATH`, and run
the current runner with `--methods sheetjet sheetjet_cached`. The recorded baseline
used a locally built v0.2 wheel. Raw reports retain file, package-source and runner
SHA-256 fingerprints, versions, all trials and verification results.

[42-trial 100k matrix](benchmarks/v0.3-100k.json) ·
[27-trial million-row matrix](benchmarks/v0.3-million.json) ·
[6-trial v0.2 baseline rerun](benchmarks/v0.2-repeat-million.json).

The methodology uses synthetic integer transactions and fixed randomized method
order. RSS is sampled every 10 ms and can miss brief peaks. Cold timings include
relevant imports; setup, fresh cache reopening and warm sessions are explicitly
different workloads. Formula-heavy models, unique-string-heavy data, wide sheets,
native rendering, refresh behavior and other machines need separate evidence.

These results support a faster cache workflow and a modest cold-read improvement.
They do not support superiority over every peer in every aspect. Native readers and
Polars remain better on several measured read workloads; structural edits and native
Excel recalculation remain outside SheetJet's contract.
