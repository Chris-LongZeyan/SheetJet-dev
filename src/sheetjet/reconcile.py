"""Key-based reconciliation over typed, local workbook projections."""

from __future__ import annotations

import json
import math
import os
import tempfile
from dataclasses import dataclass
from pathlib import Path

from .core import Budget, compact
from .errors import BudgetExceeded, SheetJetError
from .query import _literal, identifier
from .workbook import Workbook


def _value(value):
    if isinstance(value, float) and not math.isfinite(value):
        raise SheetJetError("Reconciliation cannot serialize non-finite numeric values")
    return value if isinstance(value, (type(None), str, int, float, bool)) else str(value)


def _record(row, keys, columns):
    kind, before_sheet, after_sheet, *values = row
    key = dict(zip(keys, map(_value, values[: len(keys)]), strict=True))
    values = values[len(keys) :]
    before = dict(zip(columns, map(_value, values[: len(columns)]), strict=True))
    after = dict(zip(columns, map(_value, values[len(columns) :]), strict=True))
    return {
        "kind": kind,
        "key": key,
        "before": {"sheet": before_sheet, "values": before} if before_sheet is not None else None,
        "after": {"sheet": after_sheet, "values": after} if after_sheet is not None else None,
        "changed_columns": [c for c in columns if before[c] != after[c]]
        if kind == "changed"
        else [],
    }


def _key_count(db, relation, keys, side):
    names = ",".join(map(identifier, keys))
    nulls = " OR ".join(f"{identifier(key)} IS NULL" for key in keys)
    total, missing, distinct = db.execute(
        f"SELECT count(*), count(*) FILTER (WHERE {nulls}), "
        f"count(DISTINCT ({names})) FROM {relation}"
    ).fetchone()
    if missing:
        raise SheetJetError(f"{side}: {missing} rows have null business keys")
    if distinct != total:
        raise SheetJetError(f"{side}: duplicate business keys; choose a unique composite key")
    return total


@dataclass(frozen=True)
class ReconcileOptions:
    """Execution and disclosure settings shared across reconciliation runs."""

    sample_limit: int = 20
    output: str | Path | None = None
    overwrite: bool = False
    cache_dir: str | Path | None = None
    budget: Budget | None = None
    allow_cached_formulas: bool = False
    persistent_cache: bool = True


def _validate_schema(schema):
    if (
        not isinstance(schema, dict)
        or not schema
        or any(not isinstance(c, str) or not c for c in schema)
        or len({c.casefold() for c in schema}) != len(schema)
    ):
        raise SheetJetError("Provide a schema with unique column names")


def _comparison_columns(schema, keys, compare_columns):
    if (
        not isinstance(keys, (list, tuple))
        or not keys
        or any(not isinstance(k, str) or k not in schema for k in keys)
        or len(set(keys)) != len(keys)
    ):
        raise SheetJetError("Keys must be unique schema column names")
    if compare_columns is not None and (
        not isinstance(compare_columns, (list, tuple))
        or any(not isinstance(c, str) for c in compare_columns)
    ):
        raise SheetJetError("Compare columns must be a list of schema column names")
    columns = (
        list(compare_columns)
        if compare_columns is not None
        else [c for c in schema if c not in keys]
    )
    if len(set(columns)) != len(columns) or any(c not in schema or c in keys for c in columns):
        raise SheetJetError("Compare columns must be unique non-key schema columns")
    return columns


def _destination(output, paths, overwrite):
    destination = Path(output).resolve() if output is not None else None
    if destination is not None:
        if any(
            destination == p or (destination.exists() and destination.samefile(p)) for p in paths
        ):
            raise SheetJetError("Reconciliation output must not replace either input")
        if destination.exists() and not overwrite:
            raise SheetJetError("Output exists; use overwrite=True to replace it")
    return destination


def _join_projections(left, right, after_ranges, provenance):
    engine, other = left.query_engine, right.query_engine
    # Reuse the first session's disk-backed database, including its memory limit.
    # Publish the second workbook's selected projections, never arbitrary SQL.
    sources = []
    for i, (sheet, alias) in enumerate(zip(after_ranges, other.unions["records"], strict=True)):
        path, _ = other.sources[alias.casefold()]
        name = f"_reconcile_after_{i}"
        engine._publish(name, path)
        sources.append(
            f"SELECT *, {_literal(sheet)} AS {identifier(provenance)} FROM {identifier(name)}"
        )
    engine._replace_view("after_records", " UNION ALL ".join(sources))
    return engine


def _create_changes(db, keys, columns, provenance):
    first = identifier(keys[0])
    joins = " AND ".join(f"b.{identifier(k)} = a.{identifier(k)}" for k in keys)
    differences = (
        " OR ".join(f"b.{identifier(c)} IS DISTINCT FROM a.{identifier(c)}" for c in columns)
        or "false"
    )
    fields = [
        f"CASE WHEN b.{first} IS NULL THEN 'added' WHEN a.{first} IS NULL THEN 'removed' ELSE 'changed' END AS kind",
        f"b.{identifier(provenance)} AS before_sheet",
        f"a.{identifier(provenance)} AS after_sheet",
        *[f"coalesce(b.{identifier(k)}, a.{identifier(k)}) AS k{i}" for i, k in enumerate(keys)],
        *[f"b.{identifier(c)} AS b{i}" for i, c in enumerate(columns)],
        *[f"a.{identifier(c)} AS a{i}" for i, c in enumerate(columns)],
    ]
    db.execute(
        "CREATE TABLE changes AS SELECT "
        + ",".join(fields)
        + f" FROM records b FULL OUTER JOIN after_records a ON {joins} "
        + f"WHERE b.{first} IS NULL OR a.{first} IS NULL OR ({differences})"
    )


def _change_counts(db, totals, columns):
    counts = {"added": 0, "removed": 0, "changed": 0}
    counts.update(dict(db.execute("SELECT kind,count(*) FROM changes GROUP BY kind").fetchall()))
    counts["unchanged"] = totals["before"] - counts["removed"] - counts["changed"]
    changed_columns = {}
    if columns:
        expressions = [
            f"count(*) FILTER (WHERE kind='changed' AND b{i} IS DISTINCT FROM a{i})"
            for i in range(len(columns))
        ]
        changed_columns = dict(
            zip(
                columns,
                db.execute("SELECT " + ",".join(expressions) + " FROM changes").fetchone(),
                strict=True,
            )
        )
    return counts, changed_columns


def _check_sources(left, right):
    left.package.check()
    right.package.check()
    left.query_engine._check_sources()
    right.query_engine._check_sources()


def _write_records(db, selection, stream, keys, columns):
    cursor = db.execute(selection)
    while rows := cursor.fetchmany(512):
        for row in rows:
            stream.write(
                json.dumps(_record(row, keys, columns), ensure_ascii=False, allow_nan=False) + "\n"
            )
    stream.flush()
    os.fsync(stream.fileno())


def _publish(temporary, destination, overwrite):
    if overwrite:
        os.replace(temporary, destination)
    elif os.name == "nt":
        # Windows rename refuses an existing destination atomically.
        os.rename(temporary, destination)
    else:
        os.link(temporary, destination)


def _export(left, right, destination, selection, keys, columns, overwrite):
    destination.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=".sheetjet-reconcile-", dir=destination.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as stream:
            _write_records(left.query_engine.db, selection, stream, keys, columns)
        _check_sources(left, right)
        _publish(temporary, destination, overwrite)
    finally:
        Path(temporary).unlink(missing_ok=True)


def reconcile(
    before,
    after,
    *,
    before_ranges,
    after_ranges,
    schema,
    keys,
    compare_columns=None,
    options=None,
    **legacy_options,
):
    """Compare typed records by unique keys; return a bounded summary and optional JSONL.

    Prefer ReconcileOptions for execution settings. v0.6 keyword settings remain
    supported, but cannot be combined with options. Selected formula values require
    explicit cached-value opt-in. Inputs are never written.
    """
    if options is not None and legacy_options:
        raise SheetJetError("Use options or individual execution settings, not both")
    options = options if options is not None else ReconcileOptions(**legacy_options)
    if not isinstance(options, ReconcileOptions):
        raise SheetJetError("options must be ReconcileOptions")
    budget = options.budget or Budget()
    sample_limit = options.sample_limit
    if not isinstance(sample_limit, int) or not 0 <= sample_limit <= budget.max_rows:
        raise SheetJetError("sample_limit must be between zero and the response row budget")
    _validate_schema(schema)
    columns = _comparison_columns(schema, keys, compare_columns)
    projection = {c: schema[c] for c in [*keys, *columns]}
    paths = [Path(before).resolve(), Path(after).resolve()]
    destination = _destination(options.output, paths, options.overwrite)
    provenance = "_sheetjet_source"
    while provenance.casefold() in {c.casefold() for c in projection}:
        provenance += "_"
    with (
        Workbook(paths[0], options.cache_dir, budget) as left,
        Workbook(paths[1], options.cache_dir, budget) as right,
    ):
        for book, ranges in [(left, before_ranges), (right, after_ranges)]:
            book.query_engine.load_sheets(
                "records",
                ranges,
                projection,
                source_column=provenance,
                allow_cached_formulas=options.allow_cached_formulas,
                persistent_cache=options.persistent_cache,
            )
        engine = _join_projections(left, right, after_ranges, provenance)
        db = engine.db
        totals = {
            "before": _key_count(db, "records", keys, "before"),
            "after": _key_count(db, "after_records", keys, "after"),
        }
        _create_changes(db, keys, columns, provenance)
        counts, changed_columns = _change_counts(db, totals, columns)
        order = ",".join(f"k{i}" for i in range(len(keys)))
        selection = f"SELECT * FROM changes ORDER BY {order}"
        sample = [
            _record(row, keys, columns)
            for row in db.execute(f"{selection} LIMIT {sample_limit}").fetchall()
        ]
        if len(sample) * (len(keys) + 2 * len(columns) + 2) > budget.max_cells:
            raise BudgetExceeded("Reconciliation sample exceeds cell budget; reduce sample_limit")
        result = {
            "before": str(paths[0]),
            "after": str(paths[1]),
            "before_ranges": before_ranges,
            "after_ranges": after_ranges,
            "schema": projection,
            "keys": list(keys),
            "compare_columns": columns,
            "rows": totals,
            "counts": counts,
            "changed_columns": changed_columns,
            "sample": sample,
            "sample_truncated": sum(counts[k] for k in ("added", "removed", "changed"))
            > len(sample),
            "output": str(destination) if destination else None,
            "formula_policy": "stored caches accepted; not recalculated"
            if options.allow_cached_formulas
            else "formula inputs rejected",
            "cache_hits": {
                "before": left.metrics.counts.get("persistent_table_hits", 0),
                "after": right.metrics.counts.get("persistent_table_hits", 0),
            },
        }
        compact(result, budget)  # Validate disclosure before publishing any output.

        _check_sources(left, right)
        if destination:
            _export(left, right, destination, selection, keys, columns, options.overwrite)
        return result
