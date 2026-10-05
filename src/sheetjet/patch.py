"""Surgical OOXML patching: unchanged XML bytes and opaque package members survive.

Expat locates byte spans while a decompressed sheet is spooled to disk. Only requested
cell spans, necessary new rows, and dimension metadata are rewritten. Memory is O(patches).
"""

from __future__ import annotations

import copy
import hashlib
import json
import math
import os
import re
import shutil
import tempfile
from collections import defaultdict, deque
from pathlib import Path
from xml.parsers import expat
from xml.sax.saxutils import quoteattr

from lxml import etree as ET

from .core import NS, Q, address, bounds, contains, position
from .errors import SheetJetError, UnsupportedOperation
from .package import xml

SUPPORTED = {"SET_VALUE", "SET_FORMULA", "COPY_FORMULA", "SET_STYLE", "COPY_STYLE"}


def digest_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def tag_end(file, start):
    file.seek(start)
    data = bytearray()
    quote = None
    while True:
        char = file.read(1)
        if not char:
            raise SheetJetError("Unterminated XML tag")
        data.extend(char)
        if char in (b"'", b'"'):
            quote = None if quote == char else char if quote is None else quote
        elif char == b">" and quote is None:
            return start + len(data), bytes(data)


def _rewrite_sheet(package, sheet, operations, directory):
    part = package.sheets[sheet]["part"]
    source = Path(directory) / (hashlib.sha256(part.encode()).hexdigest()[:12] + ".xml")
    with package.zip.open(part) as incoming, source.open("wb") as out:
        shutil.copyfileobj(incoming, out, 1024 * 1024)
    edits = defaultdict(list)
    for op in operations:
        edits[op["cell"]].append(op)
    targets = set(edits)
    spans, inserts, rows_seen, group_ranges, merges = {}, {}, set(), [], []
    row_for = defaultdict(list)
    for cell in targets:
        row_for[position(cell)[0]].append(cell)
    pending_rows = deque(sorted(row_for))
    dimension = None
    root_attrs = {}
    prefix = ""
    row_state = None
    active_cell = None
    sheet_data_end = None
    empty_sheet_data = None
    parser = expat.ParserCreate()
    parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)

    def reject_doctype(*_):
        raise UnsupportedOperation("DTD-bearing workbook parts cannot be patched")

    parser.StartDoctypeDeclHandler = reject_doctype

    with source.open("rb") as scan, source.open("rb") as seek:

        def start(tag, attrs):
            nonlocal row_state, active_cell, dimension, prefix, root_attrs, empty_sheet_data
            local = tag.split(":")[-1]
            at = parser.CurrentByteIndex
            if local == "worksheet":
                root_attrs = {
                    k: v for k, v in attrs.items() if k == "xmlns" or k.startswith("xmlns:")
                }
                prefix = tag[: -len("worksheet")]
            elif local == "dimension":
                end, raw = tag_end(seek, at)
                dimension = (at, end, raw, attrs.get("ref"))
            elif local == "sheetData":
                end, raw = tag_end(seek, at)
                if raw.rstrip().endswith(b"/>"):
                    empty_sheet_data = (at, end, tag)
            elif local == "row":
                rn = int(attrs["r"])
                # Missing row insertion happens before the next existing row.
                while pending_rows and pending_rows[0] < rn:
                    inserts[pending_rows.popleft()] = at
                if pending_rows and pending_rows[0] == rn:
                    pending_rows.popleft()
                rows_seen.add(rn) if rn in row_for else None
                if rn in row_for:
                    end, raw = tag_end(seek, at)
                    row_state = {
                        "r": rn,
                        "start": at,
                        "open_end": end,
                        "raw": raw,
                        "self_close": raw.rstrip().endswith(b"/>"),
                        "positions": {},
                    }
            elif local == "c":
                ref = attrs.get("r")
                if row_state is not None:
                    col = position(ref)[1]
                    for target in row_for[row_state["r"]]:
                        if (
                            target not in spans
                            and target not in row_state["positions"]
                            and position(target)[1] < col
                        ):
                            row_state["positions"][target] = at
                if ref in targets:
                    end, raw = tag_end(seek, at)
                    active_cell = (ref, at, end, raw.rstrip().endswith(b"/>"))
            elif local == "f" and attrs.get("ref"):
                group_ranges.append(attrs["ref"])
            elif local == "mergeCell":
                merges.append(attrs["ref"])

        def end(tag):
            nonlocal active_cell, row_state, sheet_data_end
            local = tag.split(":")[-1]
            at = parser.CurrentByteIndex
            if local == "c" and active_cell:
                ref, begin, open_end, empty = active_cell
                finish = open_end if empty else tag_end(seek, at)[0]
                seek.seek(begin)
                spans[ref] = (begin, finish, seek.read(finish - begin))
                active_cell = None
            elif local == "row" and row_state:
                row_state["end"] = at
                for target in row_for[row_state["r"]]:
                    if target not in spans:
                        row_state["positions"].setdefault(target, at)
                inserts[row_state["r"]] = row_state
                row_state = None
            elif local == "sheetData":
                sheet_data_end = at

        parser.StartElementHandler = start
        parser.EndElementHandler = end
        for chunk in iter(lambda: scan.read(1024 * 1024), b""):
            parser.Parse(chunk, False)
        parser.Parse(b"", True)
    if sheet_data_end is None:
        raise SheetJetError("Worksheet has no sheetData element")
    for cell in targets:
        if any(contains(r, cell) for r in group_ranges):
            raise UnsupportedOperation(
                f"{sheet}!{cell} overlaps a shared/array formula group; edit through Excel"
            )
        if any(contains(r, cell) and position(cell) != position(r.split(":")[0]) for r in merges):
            raise UnsupportedOperation(f"{sheet}!{cell} is not the anchor of its merged range")
        for table in package.tables.values():
            if table["sheet"] == sheet and contains(table["range"], cell):
                r1, _, r2, _ = bounds(table["range"])
                if (
                    position(cell)[0] < r1 + table["header_rows"]
                    or position(cell)[0] > r2 - table["total_rows"]
                ):
                    raise UnsupportedOperation(
                        "Table header/totals edits require a table-aware structural operation"
                    )

    # Wrap fragments with the original namespace bindings, including extension prefixes.
    ns_attrs = " ".join(f"{k}={quoteattr(v)}" for k, v in root_attrs.items())

    def parse_cell(raw):
        return xml(f"<wrapper {ns_attrs}>".encode() + raw + b"</wrapper>")[0]

    replacements, logs, newcells = [], [], {}
    for cell in sorted(targets, key=position):
        old = spans.get(cell)
        node = parse_cell(old[2]) if old else ET.Element(Q + "c", r=cell, nsmap={None: NS})
        before = package.cell(node)
        if before.get("formula_attributes", {}).get("t") in {"array", "shared", "dataTable"}:
            raise UnsupportedOperation("Editing special formula cells requires native Excel")
        for op in edits[cell]:
            if "expected" in op and before["value"] != op["expected"]:
                raise SheetJetError(f"Precondition failed at {sheet}!{cell}")
            if "expected_formula" in op and before.get("formula") != op["expected_formula"]:
                raise SheetJetError(f"Formula precondition failed at {sheet}!{cell}")
            kind = op["operation"]
            if kind in {"SET_STYLE", "COPY_STYLE"}:
                node.set("s", str(op["style_id"]))
            else:
                if node.get("cm") is not None or node.get("vm") is not None:
                    raise UnsupportedOperation(
                        "Rich data/dynamic-array metadata cells require native Excel editing"
                    )
                for child in list(node):
                    if child.tag in {Q + "v", Q + "f", Q + "is"}:
                        node.remove(child)
                node.attrib.pop("t", None)
                if kind in {"SET_FORMULA", "COPY_FORMULA"}:
                    ET.SubElement(node, Q + "f").text = op["formula"][1:]
                else:
                    value = op["value"]
                    if value is None:
                        pass
                    elif isinstance(value, bool):
                        node.set("t", "b")
                        ET.SubElement(node, Q + "v").text = "1" if value else "0"
                    elif isinstance(value, (int, float)):
                        ET.SubElement(node, Q + "v").text = str(value)
                    else:
                        node.set("t", "inlineStr")
                        text = ET.SubElement(ET.SubElement(node, Q + "is"), Q + "t")
                        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                        text.text = value
        after = package.cell(node)
        raw = ET.tostring(node, encoding="utf-8", with_tail=False)
        if old:
            replacements.append((old[0], old[1], raw))
        else:
            newcells[cell] = raw
        logs.append({"sheet": sheet, "cell": cell, "before": before, "after": after})

    additions = defaultdict(list)
    newrows = defaultdict(list)
    for rn, cells in row_for.items():
        missing = [c for c in sorted(cells, key=position) if c in newcells]
        if not missing:
            continue
        state = inserts.get(rn)
        if isinstance(state, dict):
            if state["self_close"]:
                opening = re.sub(rb"/\s*>$", b">", state["raw"])
                replacements.append(
                    (
                        state["start"],
                        state["open_end"],
                        opening
                        + b"".join(newcells[c] for c in missing)
                        + f"</{prefix}row>".encode(),
                    )
                )
            else:
                for c in missing:
                    additions[state["positions"][c]].append(newcells[c])
        else:
            raw = (
                f'<{prefix}row r="{rn}">'.encode()
                + b"".join(newcells[c] for c in missing)
                + f"</{prefix}row>".encode()
            )
            newrows[state if state is not None else sheet_data_end].append((rn, raw))
    for at, cells in additions.items():
        replacements.append((at, at, b"".join(cells)))
    for at, rows in newrows.items():
        raw = b"".join(raw for _, raw in sorted(rows))
        if empty_sheet_data:
            begin, finish, tag = empty_sheet_data
            replacements.append((begin, finish, f"<{tag}>".encode() + raw + f"</{tag}>".encode()))
        else:
            replacements.append((at, at, raw))
    if dimension:
        begin, end, raw, ref = dimension
        r1, c1, r2, c2 = bounds(ref)
        for cell in targets:
            r, c = position(cell)
            r1, c1, r2, c2 = min(r1, r), min(c1, c), max(r2, r), max(c2, c)
        newref = f"{address(r1, c1)}:{address(r2, c2)}"
        if bounds(ref) != bounds(newref):
            raw = re.sub(rb"\bref\s*=\s*(['\"])[^'\"]*\1", f'ref="{newref}"'.encode(), raw)
            replacements.append((begin, end, raw))
    target = source.with_suffix(".patched.xml")
    # This digest is an edit certificate: all bytes outside authorized spans are copied verbatim.
    untouched = hashlib.sha256()
    with source.open("rb") as src, target.open("wb") as dst:
        cursor = 0
        for start, end, raw in sorted(replacements, key=lambda x: (x[0], x[1])):
            if start < cursor:
                raise SheetJetError("Overlapping XML patch spans")
            remaining = start - cursor
            while remaining:
                chunk = src.read(min(remaining, 1024 * 1024))
                if not chunk:
                    raise SheetJetError("Unexpected end of worksheet")
                dst.write(chunk)
                untouched.update(chunk)
                remaining -= len(chunk)
            dst.write(raw)
            src.seek(end)
            cursor = end
        for chunk in iter(lambda: src.read(1024 * 1024), b""):
            dst.write(chunk)
            untouched.update(chunk)
    return (
        target,
        logs,
        {
            "part": part,
            "rewritten_spans": len(replacements),
            "untouched_xml_sha256": untouched.hexdigest(),
        },
    )


def apply_patch(workbook, operations, output, overwrite=False):
    p = workbook.package
    p.check()
    output = Path(output).resolve()
    if output == p.path:
        raise SheetJetError("Use a distinct output path; source workbooks are never overwritten")
    if output.suffix.lower() != p.path.suffix.lower():
        raise SheetJetError("Output must retain the source extension")
    if output.exists() and not overwrite:
        raise SheetJetError("Output exists; choose another path or explicitly enable overwrite")
    if p.features["signed"]:
        raise UnsupportedOperation("Editing would invalidate package signatures")
    if not operations or len(operations) > 10000:
        raise SheetJetError("Patch must contain between 1 and 10000 operations")
    ops = copy.deepcopy(operations)
    groups = defaultdict(list)
    touched = set()
    source_requests = defaultdict(set)
    for op in ops:
        if op.get("operation") in {"COPY_FORMULA", "COPY_STYLE"}:
            source_sheet = op.get("source_sheet", op.get("sheet"))
            if source_sheet not in p.sheets:
                raise SheetJetError(f"Unknown source sheet {source_sheet!r}")
            source_cell = address(*position(op.get("source_cell", "")))
            op["source_cell"] = source_cell
            source_requests[source_sheet].add(source_cell)
    sources = {
        sheet: {
            cell["cell"]: cell for cell in workbook.read_cells(sheet, sorted(cells, key=position))
        }
        for sheet, cells in source_requests.items()
    }
    styles = (
        xml(p.zip.read("xl/styles.xml")).find(Q + "cellXfs") if "xl/styles.xml" in p.parts else None
    )
    nstyles = len(styles) if styles is not None else 1
    for op in ops:
        kind = op.get("operation")
        if kind not in SUPPORTED:
            raise UnsupportedOperation(
                f"{kind} is not preservation-safe. Supported operations: {', '.join(sorted(SUPPORTED))}"
            )
        sheet = op.get("sheet")
        if sheet not in p.sheets:
            raise SheetJetError(f"Unknown sheet {sheet!r}")
        r, c = position(op.get("cell", ""))
        op["cell"] = address(r, c)
        identity = (sheet, op["cell"], "style" if "STYLE" in kind else "content")
        if identity in touched:
            raise SheetJetError("Multiple patches to the same cell property are ambiguous")
        touched.add(identity)
        if kind.startswith("COPY_"):
            source = sources[op.get("source_sheet", sheet)][op["source_cell"]]
            if kind == "COPY_FORMULA":
                from openpyxl.formula.translate import Translator

                if "formula" not in source or source.get("formula_attributes", {}).get("t"):
                    raise UnsupportedOperation("COPY_FORMULA requires an ordinary source formula")
                op["formula"] = Translator(
                    source["formula"], origin=source["cell"]
                ).translate_formula(op["cell"])
            else:
                op["style_id"] = source["style"]
        if kind in {"SET_FORMULA", "COPY_FORMULA"}:
            f = op.get("formula")
            if not isinstance(f, str) or not f.startswith("=") or len(f) < 2 or len(f) > 8192:
                raise SheetJetError(
                    "Formula must begin with '=' and fit Excel's formula length limit"
                )
        elif kind in {"SET_STYLE", "COPY_STYLE"}:
            style = op.get("style_id")
            if not isinstance(style, int) or isinstance(style, bool) or not 0 <= style < nstyles:
                raise SheetJetError("Style ID must refer to an existing cell style")
        else:
            if "value" not in op:
                raise SheetJetError("SET_VALUE requires a value (null explicitly clears a cell)")
            value = op["value"]
            if not isinstance(value, (str, int, float, bool, type(None))):
                raise SheetJetError(
                    "Use a JSON scalar value; dates need an explicit serial and style"
                )
            if isinstance(value, float) and not math.isfinite(value):
                raise SheetJetError("NaN and infinity are not Excel numbers")
            if isinstance(value, str) and len(value.encode("utf-16-le")) // 2 > 32767:
                raise SheetJetError("Text exceeds Excel's 32767 UTF-16-unit limit")
        groups[sheet].append(op)
    output.parent.mkdir(parents=True, exist_ok=True)
    recalc = any(op["operation"] in {"SET_VALUE", "SET_FORMULA", "COPY_FORMULA"} for op in ops)
    with (
        workbook.metrics.time("patching"),
        tempfile.TemporaryDirectory(prefix="sheetjet-patch-", dir=output.parent) as directory,
    ):
        changes, logs, certificates = {}, [], []
        for sheet, edits in groups.items():
            file, log, certificate = _rewrite_sheet(p, sheet, edits, directory)
            changes[p.sheets[sheet]["part"]] = file
            logs.extend(log)
            certificates.append(certificate)
        removed = set()
        if recalc:
            root = xml(p.zip.read("xl/workbook.xml"))
            calc = root.find(Q + "calcPr")
            if calc is None:
                calc = ET.Element(Q + "calcPr")
                following = {
                    Q + tag
                    for tag in (
                        "oleSize",
                        "customWorkbookViews",
                        "pivotCaches",
                        "smartTagPr",
                        "smartTagTypes",
                        "webPublishing",
                        "fileRecoveryPr",
                        "webPublishObjects",
                        "extLst",
                    )
                }
                insertion = next(
                    (i for i, node in enumerate(root) if node.tag in following), len(root)
                )
                root.insert(insertion, calc)
            calc.set("fullCalcOnLoad", "1")
            calc.set("forceFullCalc", "1")
            changes["xl/workbook.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            # A stale calculation chain may point at removed/replaced formulas.
            relpath = "xl/_rels/workbook.xml.rels"
            relroot = xml(p.zip.read(relpath))
            for node in list(relroot):
                if node.get("Type", "").endswith("/calcChain"):
                    removed.add(p.rels[node.get("Id")]["target"])
                    relroot.remove(node)
            if removed:
                changes[relpath] = ET.tostring(relroot, encoding="utf-8", xml_declaration=True)
                types = xml(p.zip.read("[Content_Types].xml"))
                for node in list(types):
                    if node.get("PartName", "").lstrip("/") in removed:
                        types.remove(node)
                changes["[Content_Types].xml"] = ET.tostring(
                    types, encoding="utf-8", xml_declaration=True
                )
        pending = Path(directory) / output.name
        from .archive import copy_archive

        compressed_records = copy_archive(p, pending, changes, removed, directory)
        with workbook.metrics.time("validation"):
            from .validation import validate_patch

            verification = validate_patch(workbook, pending, logs, set(changes), removed)
        p.check()
        report = {
            "source": str(p.path),
            "output": str(output),
            "operations": len(ops),
            "changed_cells": len(logs),
            "changed_parts": sorted(changes),
            "removed_parts": sorted(removed),
            "certificates": certificates,
            "compressed_records": compressed_records,
            "validation": verification,
            "requires_recalculation": recalc,
            "formula_cache_status": "potentially stale until recalculated"
            if recalc
            else "unchanged",
            "changes": logs,
        }
        # Persist the audit before publishing the workbook. No truncated log reaches the model.
        audit = output.with_suffix(output.suffix + ".sheetjet.json")
        if audit.exists() and not overwrite:
            raise SheetJetError("Audit path exists; choose another output or enable overwrite")
        pending_audit = Path(directory) / "audit.json"
        pending_audit.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        if output.exists() and not overwrite:
            raise SheetJetError("Output appeared during editing; refusing to overwrite")
        os.replace(pending, output)
        os.replace(pending_audit, audit)
        workbook.metrics.add("write_operations", len(ops))
        return {
            k: v
            for k, v in report.items()
            if k not in {"changes", "certificates", "compressed_records"}
        } | {"audit": str(audit)}
