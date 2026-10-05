"""Bounded-memory OOXML access; never round-trip a workbook through an object model."""

from __future__ import annotations

import hashlib
import json
import posixpath
import sqlite3
from collections import OrderedDict
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from lxml import etree as ET

from .core import NS, REL, Metrics, Q, bounds, position
from .errors import SheetJetError, UnsupportedOperation


def xml(data):
    return ET.fromstring(data, ET.XMLParser(resolve_entities=False, no_network=True))


def release(element):
    element.clear()
    while element.getprevious() is not None:
        del element.getparent()[0]


class Package:
    def __init__(self, path, cache_dir, metrics: Metrics):
        self.path = Path(path).resolve()
        self.metrics = metrics
        if self.path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise UnsupportedOperation(
                "Only XLSX/XLSM OOXML is supported; convert XLS/XLSB explicitly first"
            )
        self.stamp = self._stamp()
        try:
            self.zip = ZipFile(self.path)
        except BadZipFile as e:
            raise SheetJetError("Not an unencrypted OOXML ZIP workbook") from e
        self.parts = {i.filename: i for i in self.zip.infolist()}
        if len(self.parts) != len(self.zip.infolist()):
            self.zip.close()
            raise SheetJetError("Duplicate ZIP members are ambiguous")
        if any(i.flag_bits & 1 for i in self.parts.values()):
            self.zip.close()
            raise UnsupportedOperation("Encrypted ZIP members are unsupported")
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.strings = None
        self.string_lru = OrderedDict()
        self._metadata()

    def _stamp(self):
        st = self.path.stat()
        return st.st_size, st.st_mtime_ns, st.st_ctime_ns

    def check(self):
        if self._stamp() != self.stamp:
            raise SheetJetError("Workbook changed during this session; reopen before continuing")

    def signature(self, *parts):
        records = [
            (p, self.parts[p].CRC, self.parts[p].file_size) for p in parts if p in self.parts
        ]
        return hashlib.sha256(json.dumps(records).encode()).hexdigest()

    def relationships(self, part):
        relpath = posixpath.join(
            posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels"
        )
        if relpath not in self.parts:
            return {}
        out = {}
        for node in xml(self.zip.read(relpath)):
            if node.get("TargetMode") == "External":
                out[node.get("Id")] = {
                    "external": True,
                    "target": node.get("Target"),
                    "type": node.get("Type"),
                }
                continue
            target = node.get("Target", "")
            target = (
                posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
                if not target.startswith("/")
                else target[1:]
            )
            if target.startswith("../") or target not in self.parts:
                raise SheetJetError(f"Invalid package relationship: {target}")
            out[node.get("Id")] = {"target": target, "type": node.get("Type"), "external": False}
        return out

    def _metadata(self):
        if "xl/workbook.xml" not in self.parts:
            raise UnsupportedOperation(
                "Workbook must use the standard xl/workbook.xml package layout"
            )
        root = xml(self.zip.read("xl/workbook.xml"))
        if root.tag != Q + "workbook":
            raise UnsupportedOperation("Strict OOXML is not supported; save as transitional XLSX")
        self.rels = self.relationships("xl/workbook.xml")
        self.shared_strings_part = next(
            (r["target"] for r in self.rels.values() if r["type"].endswith("/sharedStrings")), None
        )
        self.sheets = {}
        self.tables = {}
        for item in root.findall(Q + "sheets/" + Q + "sheet"):
            rel = self.rels[item.get("{" + REL + "}id")]
            if not rel["type"].endswith("/worksheet"):
                continue  # Chartsheets remain opaque and are preserved by patches.
            name, part = item.get("name"), rel["target"]
            meta = {
                "name": name,
                "part": part,
                "state": item.get("state", "visible"),
                "range": None,
                "range_source": "declared (may be stale)",
                "tables": [],
            }
            with self.zip.open(part) as stream:
                for _, element in ET.iterparse(
                    stream, events=("start",), resolve_entities=False, no_network=True
                ):
                    if element.tag == Q + "dimension":
                        meta["range"] = element.get("ref")
                    elif element.tag == Q + "pane":
                        meta["freeze_panes"] = dict(element.attrib)
                    elif element.tag == Q + "sheetData":
                        break
            for link in self.relationships(part).values():
                if link["type"].endswith("/table") and not link["external"]:
                    table = xml(self.zip.read(link["target"]))
                    t = {
                        "name": table.get("displayName"),
                        "sheet": name,
                        "range": table.get("ref"),
                        "part": link["target"],
                        "header_rows": int(table.get("headerRowCount", "1")),
                        "total_rows": int(table.get("totalsRowCount", "0")),
                        "columns": [
                            c.get("name")
                            for c in table.findall(Q + "tableColumns/" + Q + "tableColumn")
                        ],
                    }
                    if t["name"] in self.tables:
                        raise SheetJetError("Duplicate table names")
                    self.tables[t["name"]] = t
                    meta["tables"].append(t["name"])
            self.sheets[name] = meta
        self.names = [
            dict(n.attrib, formula=n.text or "")
            for n in root.findall(Q + "definedNames/" + Q + "definedName")
        ]
        calc = root.find(Q + "calcPr")
        self.calculation = dict(calc.attrib) if calc is not None else {}
        self.features = {
            "external_links": sum(
                p.startswith("xl/externalLinks/externalLink") and p.endswith(".xml")
                for p in self.parts
            ),
            "charts": sum(
                p.startswith("xl/charts/chart") and p.endswith(".xml") for p in self.parts
            ),
            "pivots": sum(
                p.startswith("xl/pivotTables/") and p.endswith(".xml") for p in self.parts
            ),
            "vba": "xl/vbaProject.bin" in self.parts,
            "signed": any(p.startswith("_xmlsignatures/") for p in self.parts),
        }

    def _load_strings(self):
        part = self.shared_strings_part
        if part is None:
            raise SheetJetError("Cell references a missing shared string table")
        dbpath = self.cache_dir / ("strings-" + self.signature(part) + ".sqlite")
        self.strings = sqlite3.connect(dbpath)
        self.strings.execute(
            "CREATE TABLE IF NOT EXISTS strings (id INTEGER PRIMARY KEY, value TEXT)"
        )
        self.strings.execute("CREATE TABLE IF NOT EXISTS ready (ok INTEGER)")
        if not self.strings.execute("SELECT 1 FROM ready").fetchone():
            with self.strings:
                self.strings.execute("DELETE FROM strings")
                batch = []
                with self.zip.open(part) as stream:
                    for _, el in ET.iterparse(
                        stream,
                        events=("end",),
                        tag=Q + "si",
                        resolve_entities=False,
                        no_network=True,
                    ):
                        # Phonetic rPh runs are annotations, not the displayed value.
                        text = "".join(
                            el.xpath("./m:t/text() | ./m:r/m:t/text()", namespaces={"m": NS})
                        )
                        batch.append((self.metrics.counts.get("shared_strings_loaded", 0), text))
                        self.metrics.add("shared_strings_loaded")
                        if len(batch) >= 4096:
                            self.strings.executemany("INSERT INTO strings VALUES (?,?)", batch)
                            batch.clear()
                        release(el)
                self.strings.executemany("INSERT INTO strings VALUES (?,?)", batch)
                self.strings.execute("INSERT INTO ready VALUES (1)")

    def string(self, key):
        if key in self.string_lru:
            self.string_lru.move_to_end(key)
            return self.string_lru[key]
        if self.strings is None:
            self._load_strings()
        row = self.strings.execute("SELECT value FROM strings WHERE id=?", (int(key),)).fetchone()
        if row is None:
            raise SheetJetError(f"Invalid shared string index {key}")
        self.string_lru[key] = row[0]
        if len(self.string_lru) > 4096:
            self.string_lru.popitem(last=False)
        return row[0]

    def cell(self, node, *, values_only=False, allow_cached_formulas=False, numeric_text=False):
        kind = node.get("t", "n")
        raw, formula, inline = None, None, None
        for child in node:
            if child.tag == Q + "v":
                raw = child.text
            elif child.tag == Q + "f":
                formula = child
            elif child.tag == Q + "is":
                inline = child
        if raw == "" and kind in {"n", "b", "s"}:
            raw = None
        value = raw
        if kind == "s" and raw is not None:
            value = self.string(raw)
        elif kind == "inlineStr":
            pieces = []
            if inline is not None:
                for child in inline:
                    if child.tag == Q + "t":
                        pieces.append(child.text or "")
                    elif child.tag == Q + "r":
                        pieces.extend(n.text or "" for n in child if n.tag == Q + "t")
            value = "".join(pieces)
        elif kind == "b" and raw is not None:
            value = raw == "1"
        elif kind == "n" and raw is not None:
            value = (
                raw
                if numeric_text
                else float(raw)
                if any(c in raw.lower() for c in ".e")
                else int(raw)
            )
        if values_only:
            if kind == "e":
                raise SheetJetError(f"Excel error at {node.get('r')}: {value}")
            if formula is not None and (not allow_cached_formulas or value is None):
                raise SheetJetError(
                    "Formula inputs require explicit cached-value opt-in and nonempty caches; SheetJet does not calculate Excel formulas"
                )
            return value
        return {
            "cell": node.get("r"),
            "value": value,
            "type": kind,
            "style": int(node.get("s", "0")),
            **(
                {"formula": "=" + (formula.text or ""), "formula_attributes": dict(formula.attrib)}
                if formula is not None
                else {}
            ),
        }

    def projected_rows(self, sheet, ref, columns, allow_cached_formulas=False):
        """Decode scalar projections directly, without allocating cell metadata dictionaries."""
        from .projection import projected_rows

        yield from projected_rows(self, sheet, ref, columns, allow_cached_formulas)

    def rows(self, sheet, ref=None, columns=None):
        self.check()
        if sheet not in self.sheets:
            raise SheetJetError(f"Unknown sheet {sheet!r}")
        r1, c1, r2, c2 = bounds(ref) if ref else (1, 1, 1048576, 16384)
        with self.zip.open(self.sheets[sheet]["part"]) as stream:
            for _, row in ET.iterparse(
                stream, events=("end",), tag=Q + "row", resolve_entities=False, no_network=True
            ):
                rownum = int(row.get("r"))
                if rownum > r2:
                    break
                result = []
                if rownum >= r1:
                    scanned = 0
                    for cell in row.iterchildren(Q + "c"):
                        scanned += 1
                        _, col = position(cell.get("r"))
                        if c1 <= col <= c2 and (columns is None or col in columns):
                            result.append(self.cell(cell))
                    self.metrics.add("cells_scanned", scanned)
                    self.metrics.add("cells_decoded", len(result))
                yield rownum, result
                release(row)

    def close(self):
        if self.strings:
            self.strings.close()
        self.zip.close()
