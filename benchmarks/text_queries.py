"""Check repeated text-filter latency after first-query startup on the wide fixture."""

import argparse
import hashlib
import json
import statistics
import time
from pathlib import Path

import sheetjet
from sheetjet import Workbook
from sheetjet.paths import workspace_path


def run(source, output, blocks=5, queries=100, workspace=None):
    workspace = workspace if workspace is not None else Path.cwd()
    source = workspace_path(source, workspace)
    output = workspace_path(output, workspace)
    if source == output or (output.exists() and output.samefile(source)):
        raise ValueError("Benchmark output must not replace its input")
    if output.exists():
        raise FileExistsError("Choose a new benchmark output path")
    if min(blocks, queries) < 1:
        raise ValueError("blocks and queries must be positive")
    output.parent.mkdir(parents=True, exist_ok=True)
    start = time.perf_counter()
    with Workbook(source, output.parent / "cache") as book:
        engine = book.query_engine
        engine.load(
            "Wide",
            sheet="Wide",
            ref="A1:CRD41",
            columns=["Region", "Amount"],
            schema={"Region": "VARCHAR", "Amount": "BIGINT"},
            persistent_cache=False,
        )
        sql = 'SELECT SUM("Amount") FROM Wide WHERE "Region" = ?'
        assert engine.query(sql, ["APAC"])["rows"] == [[1260]]
        first = time.perf_counter() - start
        samples = []
        for _ in range(blocks):
            start = time.perf_counter()
            for _ in range(queries):
                assert engine.query(sql, ["APAC"])["rows"] == [[1260]]
            samples.append((time.perf_counter() - start) * 1000 / queries)
    engine_path = Path(sheetjet.__file__).with_name("query.py")
    result = {
        "input_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        "query_engine_sha256": hashlib.sha256(engine_path.read_bytes()).hexdigest(),
        "open_load_first_query_seconds": first,
        "queries_per_block": queries,
        "milliseconds_per_query_by_block": samples,
        "median_milliseconds_per_query": statistics.median(samples),
        "scope": "One process, no persistent data cache; repeated parameterized queries after startup, no output editing. Blocks are not independent processes.",
    }
    output.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(json.dumps(result))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--blocks", type=int, default=5)
    parser.add_argument("--queries", type=int, default=100)
    parser.add_argument("--workspace", type=Path, default=Path.cwd())
    args = parser.parse_args()
    run(args.source, args.output, args.blocks, args.queries, args.workspace)
