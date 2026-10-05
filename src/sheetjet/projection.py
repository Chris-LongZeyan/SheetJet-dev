"""Streaming scalar OOXML projection without allocating XML trees or cell objects."""

from collections import deque
from xml.parsers import expat

from .core import NS, address, bounds
from .errors import SheetJetError, UnsupportedOperation


class _RangeComplete(Exception):
    pass


def projected_rows(package, sheet, ref, columns, allow_cached_formulas=False):
    package.check()
    r1, _, r2, _ = bounds(ref)
    selected = {address(1, c)[:-1]: i for i, c in enumerate(columns)}
    q = NS + "}"
    worksheet, sheet_data, row_tag, cell_tag = (
        q + t for t in ("worksheet", "sheetData", "row", "c")
    )
    value_tag, formula_tag, inline_tag, text_tag, run_tag = (
        q + t for t in ("v", "f", "is", "t", "r")
    )
    parser = expat.ParserCreate(namespace_separator="}")
    parser.buffer_text = True
    parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)
    path, pending = [], deque()
    rn = previous = 0
    values = None
    slot = None
    kind, cell_ref = "n", ""
    raw, inline = [], []
    formula, capture = False, None
    scanned = decoded = 0
    seen = 0

    def reject_doctype(*_):
        raise UnsupportedOperation("DTD-bearing workbook parts cannot be projected")

    def start(tag, attrs):
        nonlocal rn, previous, values, slot, kind, cell_ref, raw, inline, formula, capture
        nonlocal scanned, decoded, seen
        path.append(tag)
        depth = len(path)
        if depth == 3 and tag == row_tag and path[:2] == [worksheet, sheet_data]:
            rn = int(attrs["r"])
            if rn <= previous or rn > 1048576:
                raise SheetJetError("Worksheet rows must have increasing valid row numbers")
            previous = rn
            if rn > r2:
                raise _RangeComplete
            values = [None] * len(columns) if rn >= r1 else None
            seen = 0
        elif depth == 4 and tag == cell_tag and values is not None and path[-2] == row_tag:
            scanned += 1
            cell_ref = attrs.get("r", "")
            letters = cell_ref.rstrip("0123456789")
            slot = selected.get(letters)
            if slot is not None:
                if seen & (1 << slot):
                    raise SheetJetError(f"Duplicate projected cell {cell_ref}")
                seen |= 1 << slot
                if cell_ref[len(letters) :] != str(rn):
                    raise SheetJetError(f"Cell {cell_ref!r} does not match its worksheet row")
                decoded += 1
                kind = attrs.get("t", "n")
                raw, inline, formula = [], [], False
        elif slot is not None:
            if depth == 5 and tag == formula_tag:
                formula = True
            elif depth == 5 and tag == value_tag:
                capture = raw
            elif tag == text_tag and (
                (depth == 6 and path[-2] == inline_tag)
                or (depth == 7 and path[-3:-1] == [inline_tag, run_tag])
            ):
                capture = inline

    def data(text):
        if capture is not None:
            capture.append(text)

    def end(tag):
        nonlocal values, slot, capture
        depth = len(path)
        if capture is not None and tag in (value_tag, text_tag):
            capture = None
        if depth == 4 and tag == cell_tag and slot is not None:
            value = "".join(raw) or None
            if kind == "inlineStr":
                value = "".join(inline)
            elif kind == "s" and value is not None:
                value = package.string(value)
            elif kind == "b" and value is not None:
                value = value == "1"
            elif kind == "e":
                raise SheetJetError(f"Excel error at {cell_ref}: {value}")
            if formula and (not allow_cached_formulas or value is None):
                raise SheetJetError(
                    "Formula inputs require explicit cached-value opt-in and nonempty caches; "
                    "SheetJet does not calculate Excel formulas"
                )
            values[slot] = value
            slot = None
        elif depth == 3 and tag == row_tag and values is not None:
            pending.append((rn, values))
            values = None
            if rn == r2:
                raise _RangeComplete
        path.pop()

    parser.StartDoctypeDeclHandler = reject_doctype
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CharacterDataHandler = data
    try:
        with package.zip.open(package.sheets[sheet]["part"]) as stream:
            try:
                while chunk := stream.read(128 * 1024):
                    parser.Parse(chunk, False)
                    while pending:
                        yield pending.popleft()
                parser.Parse(b"", True)
            except _RangeComplete:
                pass
            while pending:
                yield pending.popleft()
    finally:
        package.metrics.add("cells_scanned", scanned)
        package.metrics.add("cells_decoded", decoded)
