"""Incremental, disk-backed lexical index. Formula runs are stored once per pattern."""

from __future__ import annotations

import json
import re
import sqlite3
from collections import Counter

from lxml import etree as ET

from .core import Q, address, position
from .errors import SheetJetError
from .package import release


def normalize(formula, cell):
    from openpyxl.formula.tokenizer import TokenizerError
    from openpyxl.formula.translate import Translator, TranslatorError

    # Translate to a common distant anchor; unsupported translations retain the original formula.
    try:
        return Translator(formula, origin=cell).translate_formula("XFD1048576")
    except (TranslatorError, TokenizerError, ValueError):
        return formula


def dependencies(formula):
    from openpyxl.formula.tokenizer import Tokenizer, TokenizerError

    try:
        refs = [
            t.value
            for t in Tokenizer(formula).items
            if t.type == "OPERAND" and t.subtype == "RANGE"
        ]
    except TokenizerError:
        return {"references": [], "complete": False}
    dynamic = bool(re.search(r"\b(INDIRECT|OFFSET)\s*\(", formula, re.IGNORECASE))
    return {
        "references": list(dict.fromkeys(refs)),
        "complete": not dynamic,
        "scope": "lexical precedents; names, structured references and external links are not resolved",
    }


class Index:
    def __init__(self, package):
        self.p = package
        import hashlib

        key = hashlib.sha256(str(package.path).encode()).hexdigest()[:24]
        self.db = sqlite3.connect(package.cache_dir / ("index-" + key + ".sqlite"))
        self.db.executescript("""
            CREATE TABLE IF NOT EXISTS sheets (name TEXT PRIMARY KEY, signature TEXT, mode TEXT, metadata TEXT);
            CREATE VIRTUAL TABLE IF NOT EXISTS search USING fts5(sheet UNINDEXED, cell UNINDEXED, kind UNINDEXED, body, tokenize='unicode61');
            CREATE TABLE IF NOT EXISTS patterns (sheet TEXT, first TEXT, last TEXT, formula TEXT, pattern TEXT, count INTEGER);
            CREATE INDEX IF NOT EXISTS pattern_sheet ON patterns(sheet);
        """)

    def signature(self, sheet):
        parts = [self.p.sheets[sheet]["part"]]
        if self.p.shared_strings_part:
            parts.append(self.p.shared_strings_part)
        parts.extend(self.p.tables[t]["part"] for t in self.p.sheets[sheet]["tables"])
        return "v2:" + self.p.signature(*parts) + ":" + self.p.sheets[sheet]["state"]

    def build(self, sheet, mode="labels"):
        from openpyxl.formula.translate import Translator

        if mode not in {"labels", "full"}:
            raise SheetJetError("Index mode must be labels or full")
        self.p.check()
        sig = self.signature(sheet)
        prior = self.db.execute(
            "SELECT signature,mode,metadata FROM sheets WHERE name=?", (sheet,)
        ).fetchone()
        if prior and prior[0] == sig and (prior[1] == mode or prior[1] == "full"):
            self.p.metrics.add("index_cache_hits")
            return json.loads(prior[2])
        meta = {
            **self.p.sheets[sheet],
            "nonempty_cells": 0,
            "formula_count": 0,
            "merged_ranges": [],
            "hidden_rows": [],
            "hidden_columns": [],
            "headers": [],
            "formula_groups": [],
            "style_counts": {},
            "index_mode": mode,
        }
        text_batch, run_batch, runs = [], [], {}
        minr, minc, maxr, maxc = 1048576, 16384, 0, 0
        styles = Counter()
        masters = {}

        def flush():
            self.db.executemany("INSERT INTO search VALUES (?,?,?,?)", text_batch)
            self.db.executemany("INSERT INTO patterns VALUES (?,?,?,?,?,?)", run_batch)
            text_batch.clear()
            run_batch.clear()

        def finish(col):
            if col in runs:
                x = runs.pop(col)
                run_batch.append((sheet, x[0], x[1], x[2], x[3], x[4]))

        with self.p.metrics.time("indexing"), self.db:
            self.db.execute("DELETE FROM search WHERE sheet=?", (sheet,))
            self.db.execute("DELETE FROM patterns WHERE sheet=?", (sheet,))
            with self.p.zip.open(self.p.sheets[sheet]["part"]) as stream:
                for _, el in ET.iterparse(
                    stream,
                    events=("end",),
                    tag=(Q + "row", Q + "mergeCell", Q + "col"),
                    resolve_entities=False,
                    no_network=True,
                ):
                    if el.tag == Q + "row":
                        rn = int(el.get("r"))
                        if el.get("hidden") in {"1", "true"}:
                            if meta["hidden_rows"] and meta["hidden_rows"][-1][1] == rn - 1:
                                meta["hidden_rows"][-1][1] = rn
                            else:
                                meta["hidden_rows"].append([rn, rn])
                        labels = []
                        for node in el.iterchildren(Q + "c"):
                            self.p.metrics.add("cells_scanned")
                            cell = self.p.cell(node)
                            self.p.metrics.add("cells_decoded")
                            r, c = position(cell["cell"])
                            styles[cell["style"]] += 1
                            if cell["value"] is not None or "formula" in cell:
                                meta["nonempty_cells"] += 1
                                minr, minc, maxr, maxc = (
                                    min(minr, r),
                                    min(minc, c),
                                    max(maxr, r),
                                    max(maxc, c),
                                )
                            if "formula" in cell:
                                meta["formula_count"] += 1
                                attrs = cell["formula_attributes"]
                                formula = cell["formula"]
                                if attrs.get("t") == "shared":
                                    si = attrs.get("si")
                                    if formula != "=":
                                        masters[si] = (cell["cell"], formula)
                                    elif si in masters:
                                        formula = Translator(
                                            masters[si][1], origin=masters[si][0]
                                        ).translate_formula(cell["cell"])
                                if attrs.get("ref"):
                                    meta["formula_groups"].append(
                                        {
                                            "range": attrs["ref"],
                                            "type": attrs.get("t"),
                                            "anchor": cell["cell"],
                                        }
                                    )
                                pattern = normalize(formula, cell["cell"])
                                if (
                                    c in runs
                                    and runs[c][3] == pattern
                                    and position(runs[c][1])[0] == r - 1
                                ):
                                    runs[c][1] = cell["cell"]
                                    runs[c][4] += 1
                                else:
                                    finish(c)
                                    runs[c] = [cell["cell"], cell["cell"], formula, pattern, 1]
                            else:
                                finish(c)
                            val = cell["value"]
                            if (
                                isinstance(val, str)
                                and "formula" not in cell
                                and cell["type"] != "e"
                            ):
                                if r <= 20:
                                    labels.append([cell["cell"], val[:160]])
                                if mode == "full" or r <= 20 or c <= 3:
                                    text_batch.append(
                                        (
                                            sheet,
                                            cell["cell"],
                                            "header" if r <= 20 else "label" if c <= 3 else "body",
                                            val[:2048],
                                        )
                                    )
                            if len(text_batch) + len(run_batch) >= 4096:
                                flush()
                        if labels and len(meta["headers"]) < 5:
                            meta["headers"].append({"row": rn, "labels": labels[:20]})
                    elif el.tag == Q + "mergeCell":
                        meta["merged_ranges"].append(el.get("ref"))
                    elif el.get("hidden") in {"1", "true"}:
                        meta["hidden_columns"].append([int(el.get("min")), int(el.get("max"))])
                    release(el)
            for col in list(runs):
                finish(col)
            flush()
            meta["style_counts"] = dict(styles)
            meta["actual_range"] = f"{address(minr, minc)}:{address(maxr, maxc)}" if maxr else None
            meta["likely_data_regions"] = [self.p.tables[t]["range"] for t in meta["tables"]] or (
                [meta["actual_range"]] if maxr else []
            )
            self.db.execute(
                "INSERT OR REPLACE INTO sheets VALUES (?,?,?,?)",
                (sheet, sig, mode, json.dumps(meta)),
            )
        return meta

    def search(self, query, sheets=None, limit=20, mode="labels"):
        if not 1 <= limit <= 1000:
            raise SheetJetError("Search limit must be between 1 and 1000")
        selected = list(self.p.sheets) if sheets is None else sheets
        for sheet in selected:
            self.build(sheet, mode)
        terms = re.findall(r"\w+", query, re.UNICODE)
        if not terms or not selected:
            return []
        # Quote tokens: users cannot inject FTS operators. OR surfaces partially matching regions.
        expr = " OR ".join('"' + term.replace('"', '""') + '"' for term in terms)
        placeholders = ",".join("?" for _ in selected)
        scope = " AND kind IN ('header','label')" if mode == "labels" else ""
        rows = self.db.execute(
            f"SELECT sheet,cell,kind,body,bm25(search) FROM search WHERE search MATCH ? AND sheet IN ({placeholders}){scope} ORDER BY bm25(search) LIMIT ?",
            [expr, *selected, limit],
        ).fetchall()
        return [
            dict(zip(("sheet", "cell", "kind", "text", "score"), row, strict=True)) for row in rows
        ]

    def patterns(self, sheet, limit=20, offset=0):
        if offset < 0 or not 1 <= limit <= 1000:
            raise SheetJetError("Use nonnegative offset and a limit between 1 and 1000")
        self.build(sheet)
        rows = self.db.execute(
            "SELECT first,last,formula,count FROM patterns WHERE sheet=? ORDER BY count DESC,first LIMIT ? OFFSET ?",
            (sheet, limit, offset),
        ).fetchall()
        return [{"range": f"{r[0]}:{r[1]}", "example": r[2], "count": r[3]} for r in rows]

    def close(self):
        self.db.close()
