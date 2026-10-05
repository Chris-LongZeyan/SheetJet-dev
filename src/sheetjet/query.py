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
        self.sources = {}
        self.readonly = False
        self.materialized = set()

    def _write_connection(self):
        if self.readonly:
            import duckdb

            self.db.close()
            self.db = duckdb.connect(self.path)
            self.db.execute("SET memory_limit = ?", [self.memory_limit])
            self.db.execute("SET threads = 4")
            self.db.execute("SET preserve_insertion_order = false")
            self.readonly = False

    @staticmethod
    def _stamp(path):
        st = path.stat()
        return st.st_size, st.st_mtime_ns, st.st_ctime_ns

    def _publish(self, name, path):
        """Publish a projection without copying its data into the session database."""
        path = path.resolve()
        literal = "'" + str(path).replace("'", "''") + "'"
        stamp = self._stamp(path)
        self._replace_view(name, f"SELECT * FROM read_parquet({literal})")
        self.sources[name.casefold()] = (path, stamp)

    def _replace_view(self, name, sql):
        self.db.execute("BEGIN")
        try:
            if name.casefold() in self.materialized:
                self.db.execute(f"DROP TABLE {identifier(name)}")
            self.db.execute(f"CREATE OR REPLACE VIEW {identifier(name)} AS {sql}")
            self.db.execute("COMMIT")
        except Exception:
            self.db.execute("ROLLBACK")
            raise
        self.materialized.discard(name.casefold())

    def _drop(self, name):
        kind = "TABLE" if name.casefold() in self.materialized else "VIEW"
        self.db.execute(f"DROP {kind} IF EXISTS {identifier(name)}")
        self.materialized.discard(name.casefold())
        self.sources.pop(name.casefold(), None)
        self.loaded.pop(name.casefold(), None)

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
        self._write_connection()
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
            {x.casefold() for x in headers}
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
            json.dumps(["projection-v3", p.signature(*parts), key, r2]).encode()
        ).hexdigest()
        cached = p.cache_dir / ("projection-" + signature + ".parquet")
        if persistent_cache and cached.is_file():
            import duckdb

            try:
                with self.w.metrics.time("persistent_table_read"):
                    self._publish(name, cached)
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
            # Convert directly to a typed projection; avoid a second copy in the session DB.
            import uuid

            if not persistent_cache:
                cached = Path(self.temp.name) / (uuid.uuid4().hex + ".parquet")
            pending = cached.with_suffix("." + uuid.uuid4().hex + ".tmp")
            stage_view = "_sheetjet_stage_" + uuid.uuid4().hex
            try:
                if rows:
                    self.db.read_csv(
                        str(csvpath),
                        header=False,
                        columns=schema,
                        na_values=null,
                        delimiter=",",
                        quotechar='"',
                        escapechar='"',
                    ).create_view(stage_view, replace=True)
                else:
                    defs = ",".join(
                        f"CAST(NULL AS {schema[c]}) AS {identifier(c)}" for c in selected
                    )
                    self.db.execute(
                        f"CREATE VIEW {identifier(stage_view)} AS SELECT {defs} WHERE false"
                    )
                p.check()
                with self.w.metrics.time("projection_write"):
                    self.db.execute(
                        f"COPY {identifier(stage_view)} TO ? (FORMAT PARQUET, COMPRESSION ZSTD)",
                        [str(pending)],
                    )
                    p.check()
                    os.replace(pending, cached)
                    self._publish(name, cached)
                if persistent_cache:
                    self.w.metrics.add("persistent_table_writes")
            finally:
                self.db.execute(f"DROP VIEW IF EXISTS {identifier(stage_view)}")
                pending.unlink(missing_ok=True)
                csvpath.unlink(missing_ok=True)
        p.check()
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
            self._replace_view(name, " UNION ALL ".join(sources))
        except Exception:
            for alias in created:
                self._drop(alias)
            raise
        for old in self.unions.get(name.casefold(), []):
            self._drop(old)
        self.unions[name.casefold()] = created
        return {"table": name, "sheets": list(ranges), "columns": [*schema, source_column]}

    def materialize(self, name):
        """Opt into one session-local copy for repeated scans, paying its cost explicitly."""
        import uuid

        self.w.package.check()
        if name.casefold() not in self.loaded and name.casefold() not in self.unions:
            raise SheetJetError("Materialize an already loaded relation")
        self._check_sources()
        if name.casefold() in self.materialized:
            return {"table": name, "materialized": True, "cached": True}
        self._write_connection()
        stage = "_sheetjet_materialized_" + uuid.uuid4().hex
        self.db.execute("BEGIN")
        try:
            with self.w.metrics.time("materialization"):
                self.db.execute(
                    f"CREATE TABLE {identifier(stage)} AS SELECT * FROM {identifier(name)}"
                )
                self.db.execute(f"DROP VIEW {identifier(name)}")
                self.db.execute(f"ALTER TABLE {identifier(stage)} RENAME TO {identifier(name)}")
                self.w.package.check()
                self._check_sources()
                self.db.execute("COMMIT")
        except Exception:
            self.db.execute("ROLLBACK")
            raise
        self.materialized.add(name.casefold())
        return {"table": name, "materialized": True, "cached": False}

    def _check_sources(self):
        for path, stamp in self.sources.values():
            if not path.is_file() or self._stamp(path) != stamp:
                raise SheetJetError(
                    "Projection cache changed during this session; reopen before continuing"
                )

    def query(self, sql, params=None, budget: Budget | None = None):
        self.w.package.check()
        budget = budget or self.w.budget
        # Execute queries in a separate read-only connection with external access disabled.
        import duckdb

        statements = self.db.extract_statements(sql)
        if len(statements) != 1 or statements[0].type != duckdb.StatementType.SELECT:
            raise SheetJetError("Only one SELECT query is allowed")
        self._check_sources()
        # Disabling external access also prevents SQL from reading arbitrary local files or installing extensions.
        if not self.readonly:
            self.db.close()
            try:
                self.db = duckdb.connect(
                    self.path,
                    read_only=True,
                    config={"memory_limit": self.memory_limit, "threads": 4},
                )
                self.db.execute(
                    "SET allowed_paths = ?", [[str(path) for path, _ in self.sources.values()]]
                )
                self.db.execute("SET enable_external_access = false")
                self.db.execute("SET lock_configuration = true")
                self.readonly = True
            except Exception:
                # Never leave a partially configured reader available to a later query.
                self.db.close()
                self.readonly = True
                self._write_connection()
                raise
        with self.w.metrics.time("query"):
            bounded_sql = (
                f"SELECT * FROM ({sql.rstrip().rstrip(';')}) AS answer LIMIT {budget.max_rows + 1}"
            )
            result = self.db.execute(bounded_sql, params or [])
            columns = [d[0] for d in result.description]
            rows = result.fetchall()
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
