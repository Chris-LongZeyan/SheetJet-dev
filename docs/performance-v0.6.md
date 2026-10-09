# v0.6 reconciliation measurements

This benchmark measures the new cross-workbook reconciliation workflow. It is a
cold-versus-cached comparison of SheetJet, **not a peer ranking**. The earlier
[v0.5 frozen skill comparison](performance-v0.5.md) remains separate evidence.

## Workload and results

Two synthetic workbooks contain **500,000 before records and 499,846 after records**,
distributed across four selected sheets each. Headers reorder, records reverse
order and move between sheets, and a hidden unrelated formula sheet is excluded.
Each record has an integer ID, region, and DECIMAL amount.

The independently expected result is 5,000 additions, 5,154 removals, 4,899 changed
records, and 489,947 unchanged records. Each run exports all **15,053 differences**
and returns a sample of three. The recorded summary is about 1.6k JSON characters.

| Phase | Median elapsed | Observed range | Median sampled peak RSS |
|---|---:|---:|---:|
| Cold projection caches, 3 fresh processes | 20.792 s | 19.808–26.646 s | 114.0 MiB |
| Reused projection caches, 3 fresh processes | 1.308 s | 1.280–1.346 s | 108.3 MiB |
| One-time cache preparation | 19.611 s | One observation | 114.0 MiB |

The cached median is 15.9× faster than the cold median **after paying preparation
cost**. This reflects reusable typed projections, not faster cold XLSX parsing.
OS filesystem caches were not flushed. Cold-run variation is visible in the table;
three repetitions are descriptive measurements, not a statistical guarantee.

## Method

- Windows 11 build 26200, Python 3.13.7, SheetJet 0.6.0, DuckDB 1.5.5,
  openpyxl 3.1.5, psutil 7.2.2.
- Before/after XLSX sizes: 8,986,543 and 9,077,665 bytes. ZIP compression and data
  shape matter; record count alone does not predict performance.
- The parent times each fresh worker process, including imports, typing/staging
  where required, key validation, comparison, sorted complete JSONL export, and exit.
- Workbook generation and independent export verification are outside timed work.
- RSS is sampled every 10 ms inside the worker, beginning before SheetJet import.
  Short-lived peaks may be missed; this is process RSS, not a DuckDB allocation cap.
- Cold runs use distinct empty projection-cache directories. Cached runs reuse
  one prepared cache; all eight selected projections hit on every cached run.
  Pair order alternates. No other benchmark or test suite ran concurrently.
- Verification reads every exported record, checks ascending unique IDs, classifies
  changes by independent integer arithmetic, checks exact amounts, and checks total
  counts. Input SHA-256 values must remain unchanged. All seven executions passed.
- Raw evidence records versions, input hashes, source hashes, each elapsed/RSS
  observation, counts, samples, and cache hits. Workbook timestamps make generated
  ZIP hashes vary across reproductions; arithmetic expectations stay fixed.

[Raw measurements](benchmarks/reconciliation-v0.6.json) ·
[Benchmark runner](../benchmarks/reconciliation.py) ·
[Fixture generator](../examples/reconcile.py)

## Reproduce

From the repository root, choose a fresh output directory:

```console
python -m pip install -e ".[test]"
python -m benchmarks.reconciliation --rows 500000 --repeats 3 --output benchmark-output/reconciliation-v0.6
```

The runner saves `report.json`, generated inputs/spec, per-run complete change
files, and cache directories. Large runs require temporary disk space. Use a
smaller `--rows` value for a smoke test; it is a different workload.

## Scope and remaining work

These fixtures cover a narrow, homogeneous three-column dataset, exact numeric
changes, row/header reordering, sheet movement, and projection reuse. They do not
establish speed on wide text-heavy files, formula recalculation, native Excel
fidelity, or adversarial workbook packages. This release has no reconciliation
comparison against another agent skill; such a claim requires equivalent tasks,
formula semantics, matching rules, complete outputs, and independent validation.

Regression tests separately cover composite keys, null transitions, duplicate keys
within/across sheets, high-precision decimals, ignored formula columns, formula
cache opt-in, response budgets, empty inputs, quoted names, export races, and
interrupted writes. See the [feature contract](reconciliation.md).
