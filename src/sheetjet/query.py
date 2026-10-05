"""A disk-backed DuckDB staging engine with explicit typing and bounded results."""

from __future__ import annotations

import csv
import hashlib
import json
import os
import re
import tempfile
from pathlib import Path

from .core import Budget, address, bounds
from .errors import BudgetExceeded, SheetJetError


def identifier(value):
    return '"' + value.replace('"', '""') + '"'


class QueryEngine:
    def __init__(self, workbook, memory_limit="512MB"):
        import duckdb

        self.w = workbook
        self.temp = tempfile.TemporaryDirectory(prefix="sheetjet-query-")
        self.path = str(Path(self.temp.name) / "query.duckdb")
        self.memory_limit = memory_limit
        self.db = duckdb.connect(self.path)
        self.db.execute("SET memory_limit = ?", [memory_limit])
        self.db.execute("SET threads = 4")
        self.db.execute("SET preserve_insertion_order = false")
        self.loaded = {}
        self.unions = {}

    def load(
        self,
        name,
        sheet=None,
        ref=None,
        schema=None,
        columns=None,
        allow_cached_formulas=False,
        persistent_cache=True,
    ):
        """Stage one rectangular table once. Explicit schema avoids sample-based type loss."""
        self.w.package.check()
        table = self.w.package.tables.get(name)
        if table:
            sheet, ref = table["sheet"], table["range"]
            if table["header_rows"] != 1:
                raise SheetJetError("Headerless tables require an explicit range and a header row")
        if not sheet or not ref:
            raise SheetJetError("Provide an exact table name or sheet and range")
        r1, c1, r2, c2 = bounds(ref)
        if table:
            r2 -= table["total_rows"]
        headers = self.w.read_range(sheet, f"{self.w.address(r1, c1)}:{self.w.address(r1, c2)}")[
            "rows"
        ][0]
        if any(not isinstance(x, str) or not x.strip() for x in headers) or len(
            set(headers)
        ) != len(headers):
            raise SheetJetError(
                "Headers must be nonblank, unique strings; choose the correct header row"
            )
        selected = list(headers) if columns is None else columns
        if (
            len(set(selected)) != len(selected)
            or not selected
            or any(c not in headers for c in selected)
        ):
            raise SheetJetError("Projected columns must be unique existing header names")
        schema = schema or {c: "VARCHAR" for c in selected}
        allowed = re.compile(
            r"^(VARCHAR|DOUBLE|BIGINT|BOOLEAN|DATE|TIMESTAMP|DECIMAL\(\d{1,2},\d{1,2}\))$",
            re.IGNORECASE,
        )
        if set(schema) != set(selected) or any(not allowed.fullmatch(t) for t in schema.values()):
            raise SheetJetError(
                "Schema must cover selected columns using supported SQL scalar types"
            )
        schema = {c: schema[c] for c in selected}
        key = json.dumps([sheet, ref, schema, selected, allow_cached_formulas], sort_keys=True)
        if self.loaded.get(name.casefold()) == key:
            self.w.metrics.add("table_cache_hits")
            return {"table": name, "cached": True}
        indexes = [headers.index(c) + c1 for c in selected]
        p = self.w.package
        parts = [p.sheets[sheet]["part"]]
        if p.shared_strings_part:
            parts.append(p.shared_strings_part)
        signature = hashlib.sha256(
            json.dumps(["projection-v2", p.signature(*parts), key, r2]).encode()
        ).hexdigest()
        cached = p.cache_dir / ("projection-" + signature + ".parquet")
        if persistent_cache and cached.is_file():
            import duckdb

            try:
                with self.w.metrics.time("persistent_table_read"):
                    self.db.execute(
                        f"CREATE OR REPLACE TABLE {identifier(name)} AS SELECT * FROM read_parquet(?)",
                        [str(cached)],
                    )
                p.check()
                self.loaded[name.casefold()] = key
                self.w.metrics.add("persistent_table_hits")
                return {
                    "table": name,
                    "rows": self.db.execute(f"SELECT count(*) FROM {identifier(name)}").fetchone()[
                        0
                    ],
                    "columns": selected,
                    "persistent_cache": True,
                    "cached_formulas": allow_cached_formulas,
                }
            except duckdb.Error:
                # Interrupted/corrupt local caches are expendable; source files are never touched.
                cached.unlink(missing_ok=True)
        csvpath = Path(self.temp.name) / "stage.csv"
        rows = 0
        with self.w.metrics.time("table_staging"):
            with csvpath.open("w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file, quoting=csv.QUOTE_ALL)
                # A distinct null marker preserves empty strings without conflating them with NULL.
                import uuid

                null = "__sheetjet_null_" + uuid.uuid4().hex
                next_row = r1 + 1
                data_ref = f"{address(r1 + 1, c1)}:{address(r2, c2)}" if r2 > r1 else None
                source = (
                    p.projected_rows(sheet, data_ref, indexes, allow_cached_formulas)
                    if data_ref
                    else []
                )
                for rn, values in source:
                    while next_row < rn:
                        writer.writerow([null] * len(selected))
                        rows += 1
                        next_row += 1
                    if null in values:
                        raise SheetJetError("Null sentinel collision")
                    writer.writerow([null if v is None else v for v in values])
                    rows += 1
                    next_row = rn + 1
                while next_row <= r2:
                    writer.writerow([null] * len(selected))
                    rows += 1
                    next_row += 1
            # Parsing is deterministic and strict; no inference, silent row skipping, or type coercion to null.
            self.db.execute("BEGIN")
            try:
                self.db.execute(f"DROP TABLE IF EXISTS {identifier(name)}")
                defs = ",".join(identifier(c) + " " + schema[c] for c in selected)
                self.db.execute(f"CREATE TABLE {identifier(name)} ({defs})")
                if rows:
                    self.db.read_csv(
                        str(csvpath),
                        header=False,
                        columns=schema,
                        na_values=null,
                        delimiter=",",
                        quotechar='"',
                        escapechar='"',
                    ).create_view("_sheetjet_stage", replace=True)
                    self.db.execute(f"INSERT INTO {identifier(name)} SELECT * FROM _sheetjet_stage")
                    self.db.execute("DROP VIEW _sheetjet_stage")
                self.db.execute("COMMIT")
            except Exception:
                self.db.execute("ROLLBACK")
                raise
            finally:
                csvpath.unlink(missing_ok=True)
        p.check()
        if persistent_cache:
            import uuid

            pending = cached.with_suffix("." + uuid.uuid4().hex + ".tmp")
            try:
                with self.w.metrics.time("persistent_table_write"):
                    self.db.execute(
                        f"COPY {identifier(name)} TO ? (FORMAT PARQUET, COMPRESSION ZSTD)",
                        [str(pending)],
                    )
                    os.replace(pending, cached)
                self.w.metrics.add("persistent_table_writes")
            finally:
                pending.unlink(missing_ok=True)
        self.loaded[name.casefold()] = key
        return {
            "table": name,
            "rows": rows,
            "columns": selected,
            "cached_formulas": allow_cached_formulas,
            "persistent_cache": False,
        }

    def load_sheets(
        self,
        name,
        ranges,
        schema,
        source_column="_sheet",
        allow_cached_formulas=False,
        persistent_cache=True,
    ):
        """Union exact ranges across selected sheets; reuse each sheet's projection independently.

        The first row of each range is a header. Names/order may differ between
        sheets; projection is by name. Caller-selected ranges prune unrelated sheets.
        """
        if not ranges or not schema or source_column.casefold() in {c.casefold() for c in schema}:
            raise SheetJetError("Provide sheet ranges and a schema with no source-column collision")
        # Immutable staging aliases ensure an already published union survives a failed reload.
        import uuid

        prefix = "_sj_" + uuid.uuid4().hex
        sources = []
        created = []
        try:
            for i, (sheet, ref) in enumerate(ranges.items()):
                alias = f"{prefix}_{i}"
                self.load(
                    alias,
                    sheet=sheet,
                    ref=ref,
                    schema=schema,
                    columns=list(schema),
                    allow_cached_formulas=allow_cached_formulas,
                    persistent_cache=persistent_cache,
                )
                created.append(alias)
                quoted = "'" + sheet.replace("'", "''") + "'"
                sources.append(
                    f"SELECT *, {quoted} AS {identifier(source_column)} FROM {identifier(alias)}"
                )
            self.db.execute(
                f"CREATE OR REPLACE VIEW {identifier(name)} AS " + " UNION ALL ".join(sources)
            )
        except Exception:
            for alias in created:
                self.db.execute(f"DROP TABLE IF EXISTS {identifier(alias)}")
                self.loaded.pop(alias.casefold(), None)
            raise
        for old in self.unions.get(name.casefold(), []):
            self.db.execute(f"DROP TABLE {identifier(old)}")
            self.loaded.pop(old.casefold(), None)
        self.unions[name.casefold()] = created
        return {"table": name, "sheets": list(ranges), "columns": [*schema, source_column]}

    def query(self, sql, params=None, budget: Budget | None = None):
        self.w.package.check()
        budget = budget or self.w.budget
        # Execute queries in a separate read-only connection with external access disabled.
        import duckdb

        statements = self.db.extract_statements(sql)
        if len(statements) != 1 or statements[0].type != duckdb.StatementType.SELECT:
            raise SheetJetError("Only one SELECT query is allowed")
        # Disabling external access also prevents SQL from reading arbitrary local files or installing extensions.
        self.db.close()
        reader = None
        try:
            reader = duckdb.connect(
                self.path,
                read_only=True,
                config={
                    "enable_external_access": False,
                    "memory_limit": self.memory_limit,
                    "threads": 4,
                },
            )
            with self.w.metrics.time("query"):
                result = reader.execute(
                    f"SELECT * FROM ({sql.rstrip().rstrip(';')}) AS answer LIMIT {budget.max_rows + 1}",
                    params or [],
                )
                columns = [d[0] for d in result.description]
                rows = result.fetchall()
        finally:
            if reader:
                reader.close()
            self.db = duckdb.connect(self.path)
            self.db.execute("SET memory_limit = ?", [self.memory_limit])
            self.db.execute("SET threads = 4")
        if len(rows) > budget.max_rows or len(rows) * len(columns) > budget.max_cells:
            raise BudgetExceeded(
                "Query output exceeds response budget; aggregate or use a smaller LIMIT"
            )
        self.w.metrics.add("result_cells", len(rows) * len(columns))
        # JSON-safe values keep decimals exact instead of converting them to float.
        data = {
            "columns": columns,
            "rows": [
                [v if isinstance(v, (type(None), str, int, float, bool)) else str(v) for v in row]
                for row in rows
            ],
        }
        self.w.serialize(data, budget)
        return data

    def close(self):
        self.db.close()
        self.temp.cleanup()
