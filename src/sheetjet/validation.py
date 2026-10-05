from __future__ import annotations

from collections import defaultdict

from lxml import etree as ET

from .core import Q, address, contains, position
from .errors import SheetJetError
from .package import Package, release


def validate_patch(workbook, output, logs, changed_parts, removed_parts):
    original = workbook.package
    candidate = Package(output, original.cache_dir, workbook.metrics)
    try:
        if set(candidate.parts) != set(original.parts) - removed_parts:
            raise SheetJetError("Unexpected package member addition/removal")
        if candidate.sheets != original.sheets:
            # Dimension may expand, but sheet identity, ordering, state and tables must not change.
            def structure(p):
                return [
                    (s, {k: v for k, v in m.items() if k not in {"range"}})
                    for s, m in p.sheets.items()
                ]

            if structure(candidate) != structure(original):
                raise SheetJetError("Unexpected worksheet structure change")
        if candidate.names != original.names or candidate.tables != original.tables:
            raise SheetJetError("Names or tables changed unexpectedly")
        verified = 0
        for name, info in original.parts.items():
            if name in changed_parts or name in removed_parts:
                continue
            other = candidate.parts[name]
            if (info.CRC, info.file_size) != (other.CRC, other.file_size):
                raise SheetJetError(f"Untouched package part changed: {name}")
            verified += 1
        wanted = defaultdict(dict)
        for log in logs:
            wanted[log["sheet"]][log["cell"]] = log["after"]
        checked = 0
        for sheet, cells in wanted.items():
            with candidate.zip.open(candidate.sheets[sheet]["part"]) as stream:
                for _, row in ET.iterparse(
                    stream, events=("end",), tag=Q + "row", resolve_entities=False, no_network=True
                ):
                    for node in row.iterchildren(Q + "c"):
                        ref = node.get("r")
                        if ref in cells:
                            if candidate.cell(node) != cells.pop(ref):
                                raise SheetJetError(
                                    f"Saved cell did not match patch: {sheet}!{ref}"
                                )
                            checked += 1
                    release(row)
            if cells:
                raise SheetJetError(f"Patched cells missing: {list(cells)}")
        return {
            "xml_parse": "passed",
            "changed_cells_checked": checked,
            "untouched_parts_crc_verified": verified,
            "native_excel_recalculation": "not performed",
        }
    finally:
        candidate.close()


def validate_range(workbook, sheet, ref, expected=None):
    cells = workbook.get_cells(sheet, ref)
    errors = [{"cell": c["cell"], "error": c["value"]} for c in cells if c["type"] == "e"]
    uncached = [c["cell"] for c in cells if "formula" in c and c["value"] is None]
    mismatches = []
    lookup = {c["cell"]: c for c in cells}
    for cell, value in (expected or {}).items():
        cell = address(*position(cell))
        if not contains(ref, cell):
            raise SheetJetError("Expected-value checks must be inside the validated range")
        if lookup.get(cell, {}).get("value") != value:
            mismatches.append(cell)
    result = {
        "sheet": sheet,
        "range": ref,
        "stored_errors": errors,
        "uncached_formulas": uncached,
        "mismatches": mismatches,
        "valid": not (errors or mismatches),
        "calculated": False,
    }
    workbook.serialize(result)
    return result
