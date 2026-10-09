"""Fresh-process cold/cached reconciliation with independent JSONL verification."""

import argparse
import hashlib
import json
import os
import platform
import statistics
import subprocess
import sys
import threading
import time
from decimal import Decimal
from pathlib import Path


def worker(directory, cache, output):
    import psutil

    process = psutil.Process()
    peak = [process.memory_info().rss]
    stop = threading.Event()

    def sample():
        while not stop.wait(0.01):
            peak[0] = max(peak[0], process.memory_info().rss)

    thread = threading.Thread(target=sample, daemon=True)
    thread.start()
    try:
        from sheetjet import reconcile

        spec = json.loads((directory / "spec.json").read_text())
        summary = reconcile(
            directory / "before.xlsx",
            directory / "after.xlsx",
            **spec,
            cache_dir=cache,
            output=output,
            sample_limit=3,
        )
        print(json.dumps({"summary": summary, "peak_rss_mib": peak[0] / 1024**2}))
    finally:
        stop.set()
        thread.join()


def _verify_record(record, key, kind, rows, expected):
    if key > rows:
        assert kind == "added" and key <= rows + expected["added"]
        assert record["before"] is None
    elif key % 97 == 0:
        assert kind == "removed" and record["after"] is None
    else:
        assert key % 101 == 0 and kind == "changed"
        assert record["changed_columns"] == ["Amount"]
    for side in ("before", "after"):
        if record[side] is not None:
            cents = key * 37 + (5 if side == "after" and key <= rows and key % 101 == 0 else 0)
            assert Decimal(record[side]["values"]["Amount"]) == Decimal(cents) / 100


def verify(path, rows, expected):
    counts = {"added": 0, "removed": 0, "changed": 0}
    previous = 0
    with path.open(encoding="utf-8") as stream:
        for line in stream:
            record = json.loads(line)
            key, kind = record["key"]["ID"], record["kind"]
            assert key > previous  # Unique and deterministically ordered across sheets.
            previous = key
            counts[kind] += 1
            _verify_record(record, key, kind, rows, expected)
    assert counts == {k: expected[k] for k in counts}
    return sum(counts.values())


def _trial(directory, rows, expected, name, cache):
    output = directory / (name + ".jsonl")
    start = time.perf_counter()
    completed = subprocess.run(
        [
            sys.executable,
            "-m",
            "benchmarks.reconciliation",
            "--worker",
            str(directory / "inputs"),
            "--cache",
            str(cache),
            "--output",
            str(output),
        ],
        capture_output=True,
        text=True,
        timeout=600,
        env=os.environ.copy(),
    )
    seconds = time.perf_counter() - start
    if completed.returncode:
        raise RuntimeError(completed.stderr[-4000:])
    result = json.loads(completed.stdout)
    assert result["summary"]["counts"] == expected
    result.update(seconds=seconds, phase=name, verified_changes=verify(output, rows, expected))
    print(f"{name}: {seconds:.3f}s, {result['peak_rss_mib']:.1f} MiB, verified", flush=True)
    return result


def run(directory, rows, repeats):
    from importlib.metadata import version

    from examples.reconcile import fixtures

    directory = directory.resolve()
    paths, _, expected = fixtures(directory / "inputs", rows)
    hashes = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in paths}
    trials = []

    preparation = _trial(directory, rows, expected, "cache-preparation", directory / "warm-cache")
    for i in range(repeats):
        # Reverse each pair on alternate iterations to reduce fixed order bias.
        phases = ["cold", "cached"] if i % 2 == 0 else ["cached", "cold"]
        for phase in phases:
            cache = directory / (f"cold-cache-{i}" if phase == "cold" else "warm-cache")
            result = _trial(directory, rows, expected, f"{phase}-{i}", cache)
            assert result["summary"]["cache_hits"] == {
                "before": 4 if phase == "cached" else 0,
                "after": 4 if phase == "cached" else 0,
            }
            trials.append(result)
    assert hashes == {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in paths}
    report = {
        "platform": platform.platform(),
        "python": platform.python_version(),
        "versions": {p: version(p) for p in ["sheetjet", "duckdb", "openpyxl", "psutil"]},
        "input_hashes": hashes,
        "source_hashes": {
            str(p): hashlib.sha256(p.read_bytes()).hexdigest()
            for p in [Path("src/sheetjet/reconcile.py"), Path(__file__).relative_to(Path.cwd())]
        },
        "rows_before": rows,
        "expected": expected,
        "cache_preparation": preparation,
        "trials": trials,
        "medians": {
            phase: {
                key: statistics.median(t[key] for t in trials if t["phase"].startswith(phase + "-"))
                for key in ["seconds", "peak_rss_mib"]
            }
            for phase in ["cold", "cached"]
        },
        "scope": "Fresh processes; fixture generation and independent export verification excluded. Timings include imports, reconciliation, and complete JSONL export. RSS sampled every 10ms inside worker before SheetJet import. OS caches not flushed. Warm preparation reported separately. No peer ranking.",
    }
    (directory / "report.json").write_text(json.dumps(report, indent=2), encoding="utf-8")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--worker", type=Path)
    parser.add_argument("--cache", type=Path)
    parser.add_argument(
        "--output", type=Path, default=Path("benchmark-output/reconciliation-benchmark")
    )
    parser.add_argument("--rows", type=int, default=100000)
    parser.add_argument("--repeats", type=int, default=3)
    args = parser.parse_args()
    if args.worker:
        worker(args.worker, args.cache, args.output)
    else:
        if args.rows < 1 or args.repeats < 1:
            parser.error("rows and repeats must be positive")
        run(args.output, args.rows, args.repeats)
