# Contributing to SheetJet

Install Python 3.11+ and run `python -m pip install -e ".[test]"`, then
`python -m pytest -q`. No API keys or Excel installation are needed for unit tests.
CI runs Windows and Linux. Keep synthetic fixtures generated in temporary directories.

Run `python -m coverage run -m pytest -q` followed by `python -m coverage report`
for the same 85% combined statement/branch coverage floor enforced by CI. XML reports
are saved as CI artifacts. SonarQube's automatic analysis is separate and cannot
import coverage; moving to CI-based Sonar analysis requires a deliberate migration.
Keep its quality gate and rules intact. The only explicit analysis exclusions are
the four byte-frozen historical submissions; source, tests, examples, and maintained
benchmark harnesses remain analyzed. See [SECURITY.md](SECURITY.md) for their scope.

Good contributions include a small failing workbook fixture, preservation checks
for an Excel feature, query-planning improvements, and reproducible performance work.
Use synthetic data or files you can legally redistribute. Do not upload confidential
workbooks, cached text indexes, or audit logs from private files.

Every performance claim should include the command, fixture, file hash, dependency
versions, OS, elapsed time, peak-memory method, answer check and context-size method.
Report regressions and tasks where the baseline wins. Do not describe character/4
estimates as measured model tokens or compare warmed caches to cold baselines without
labeling them. Benchmarks are opt-in; unit tests must stay small.

For edits, prove what changed and what survived. Unknown OOXML parts are valuable
data. Prefer an explicit unsupported-operation error to a lossy round trip. Structural
features need tests for affected formulas, names, tables, charts and links before
being advertised as supported. Native Excel behavior requires native Excel testing.

For performance changes, run `python -m benchmarks.peers` with the `benchmark` extra
installed. Include fast native peers, not only openpyxl; include a cached peer workflow
when presenting cache speedups. Keep compressed-copy tests covering ZIP64, descriptors,
Unicode names and comments. New cache formats must version their key and test changed
sheet invalidation, formula-policy isolation, null/empty preservation and corruption
recovery. Benchmark code should not change while a recorded matrix is running.

Projection optimizations must retain formula/error checks and numeric lexemes.
Use the randomized sparse-sheet decoder comparisons, namespace/encoding tests and
restricted-query tests when modifying that path. Measure cold reads, cache reopening,
steady queries and materialized queries separately, with preparation time disclosed
for both SheetJet and peers. Raw reports record source hashes to identify the code
that produced the measurements.
