from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

from .core import Budget, compact
from .workbook import Workbook


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="sheetjet", description="Query Excel locally. Return only bounded evidence."
    )
    parser.add_argument("--cache-dir")
    parser.add_argument("--max-chars", type=int, default=12000)
    parser.add_argument("--max-cells", type=int, default=2000)
    parser.add_argument("--max-rows", type=int, default=100)
    parser.add_argument(
        "--metrics", action="store_true", help="Write performance counters to stderr"
    )
    sub = parser.add_subparsers(dest="command", required=True)
    p = sub.add_parser("inspect", help="Fast package metadata or deep sheet inspection")
    p.add_argument("file")
    p.add_argument("--sheet")
    p.add_argument("--deep", action="store_true")
    p.add_argument("--section", default="summary")
    p.add_argument("--offset", type=int, default=0)
    p.add_argument("--limit", type=int, default=20)
    p = sub.add_parser("find", help="Lexical search over a reusable local index")
    p.add_argument("file")
    p.add_argument("query")
    p.add_argument("--sheet", action="append")
    p.add_argument("--full", action="store_true")
    p.add_argument("--limit", type=int, default=20)
    for command in ("read", "formulas", "dependencies", "styles", "validate"):
        p = sub.add_parser(command)
        p.add_argument("file")
        p.add_argument("sheet")
        p.add_argument("range")
    p = sub.add_parser("patterns")
    p.add_argument("file")
    p.add_argument("sheet")
    p.add_argument("--limit", type=int, default=20)
    p.add_argument("--offset", type=int, default=0)
    p = sub.add_parser("query", help="Stage named tables with explicit schemas, execute SELECT")
    p.add_argument("file")
    p.add_argument("sql")
    p.add_argument(
        "--schema", required=True, help="JSON file mapping table names to {column: SQL type}"
    )
    p.add_argument("--cached-formulas", action="store_true")
    p.add_argument("--no-persistent-cache", action="store_true")
    p = sub.add_parser(
        "query-sheets", help="Union selected sheet ranges with a shared typed projection"
    )
    p.add_argument("file")
    p.add_argument("sql")
    p.add_argument(
        "--ranges", required=True, help="JSON file mapping sheet names to A1 ranges with headers"
    )
    p.add_argument("--schema", required=True, help="JSON file mapping column names to SQL types")
    p.add_argument("--name", default="Sales", help="SQL name for the combined table")
    p.add_argument("--source-column", default="_sheet")
    p.add_argument("--cached-formulas", action="store_true")
    p.add_argument("--no-persistent-cache", action="store_true")
    p = sub.add_parser(
        "patch", help="Apply a JSON patch to a distinct output with an audit sidecar"
    )
    p.add_argument("file")
    p.add_argument("patch_file")
    p.add_argument("output")
    p.add_argument("--overwrite", action="store_true")
    p = sub.add_parser(
        "reconcile", help="Compare records across two workbooks by unique business keys"
    )
    p.add_argument("before")
    p.add_argument("after")
    p.add_argument(
        "--spec",
        required=True,
        help="JSON with before_ranges, after_ranges, schema, keys, and optional compare_columns",
    )
    p.add_argument("--sample-limit", type=int, default=20)
    p.add_argument("--output", help="Optional complete JSONL change file")
    p.add_argument("--overwrite", action="store_true")
    p.add_argument("--cached-formulas", action="store_true")
    p.add_argument("--no-persistent-cache", action="store_true")
    args = parser.parse_args(argv)
    try:
        budget = Budget(args.max_cells, args.max_rows, args.max_chars)
        if args.command == "reconcile":
            from .reconcile import reconcile

            spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
            start = time.perf_counter()
            result = reconcile(
                args.before,
                args.after,
                **spec,
                sample_limit=args.sample_limit,
                output=args.output,
                overwrite=args.overwrite,
                cache_dir=args.cache_dir,
                budget=budget,
                allow_cached_formulas=args.cached_formulas,
                persistent_cache=not args.no_persistent_cache,
            )
            print(compact(result, budget))
            if args.metrics:
                print(
                    json.dumps(
                        {
                            "seconds": {"reconciliation": time.perf_counter() - start},
                            "cache_hits": result["cache_hits"],
                        }
                    ),
                    file=sys.stderr,
                )
            return 0
        with Workbook(args.file, args.cache_dir, budget) as book:
            cmd = args.command
            if cmd == "inspect":
                result = (
                    book.inspect_sheet(args.sheet, args.deep, args.section, args.offset, args.limit)
                    if args.sheet
                    else book.inspect_workbook(args.offset, args.limit)
                )
            elif cmd == "find":
                result = book.find_text(args.query, args.sheet, args.limit, args.full)
            elif cmd in {"read", "formulas", "dependencies", "styles", "validate"}:
                method = {
                    "read": "read_range",
                    "formulas": "get_formulas",
                    "dependencies": "get_dependencies",
                    "styles": "get_styles",
                    "validate": "validate_range",
                }[cmd]
                result = getattr(book, method)(args.sheet, args.range)
            elif cmd == "patterns":
                result = book.formula_patterns(args.sheet, args.limit, args.offset)
            elif cmd == "query":
                schemas = json.loads(Path(args.schema).read_text(encoding="utf-8"))
                for table, schema in schemas.items():
                    book.query_engine.load(
                        table,
                        schema=schema,
                        columns=list(schema),
                        allow_cached_formulas=args.cached_formulas,
                        persistent_cache=not args.no_persistent_cache,
                    )
                result = book.query_engine.query(args.sql)
            elif cmd == "query-sheets":
                ranges = json.loads(Path(args.ranges).read_text(encoding="utf-8"))
                schema = json.loads(Path(args.schema).read_text(encoding="utf-8"))
                book.query_engine.load_sheets(
                    args.name,
                    ranges,
                    schema,
                    source_column=args.source_column,
                    allow_cached_formulas=args.cached_formulas,
                    persistent_cache=not args.no_persistent_cache,
                )
                result = book.query_engine.query(args.sql)
            else:
                patches = json.loads(Path(args.patch_file).read_text(encoding="utf-8"))
                result = book.patch_cells(patches, args.output, args.overwrite)
            print(compact(result, budget))
            if args.metrics:
                print(compact(book.metrics.snapshot(), budget), file=sys.stderr)
        return 0
    except Exception as error:
        print(
            json.dumps(
                {"error": type(error).__name__, "message": str(error)[:1000]}, ensure_ascii=False
            ),
            file=sys.stderr,
        )
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
