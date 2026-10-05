from __future__ import annotations

import json
import math
import re
import time
from contextlib import contextmanager
from dataclasses import dataclass, field
from typing import Any

from .errors import BudgetExceeded, SheetJetError

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
Q = "{" + NS + "}"
CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$")


def position(ref: str) -> tuple[int, int]:
    match = CELL.fullmatch(ref)
    if not match:
        raise SheetJetError(f"Expected an A1 cell reference: {ref!r}")
    col = 0
    for char in match[1].upper():
        col = col * 26 + ord(char) - 64
    row = int(match[2])
    if row > 1048576 or col > 16384:
        raise SheetJetError("Reference exceeds Excel worksheet limits")
    return row, col


def address(row: int, col: int) -> str:
    letters = ""
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row}"


def bounds(ref: str) -> tuple[int, int, int, int]:
    parts = ref.split(":")
    if len(parts) > 2:
        raise SheetJetError("Expected a single rectangular A1 range")
    r1, c1 = position(parts[0])
    r2, c2 = position(parts[-1])
    if r2 < r1 or c2 < c1:
        raise SheetJetError("Range endpoints are reversed")
    return r1, c1, r2, c2


def contains(ref: str, cell: str) -> bool:
    r1, c1, r2, c2 = bounds(ref)
    r, c = position(cell)
    return r1 <= r <= r2 and c1 <= c <= c2


@dataclass
class Metrics:
    seconds: dict[str, float] = field(default_factory=dict)
    counts: dict[str, int] = field(default_factory=dict)

    def add(self, name: str, count: int = 1):
        self.counts[name] = self.counts.get(name, 0) + count

    @contextmanager
    def time(self, name: str):
        start = time.perf_counter()
        try:
            yield
        finally:
            self.seconds[name] = self.seconds.get(name, 0) + time.perf_counter() - start

    def snapshot(self):
        return {
            "seconds": {k: round(v, 6) for k, v in self.seconds.items()},
            "counts": dict(self.counts),
        }


@dataclass(frozen=True)
class Budget:
    max_cells: int = 2000
    max_rows: int = 100
    max_chars: int = 12000

    def __post_init__(self):
        if min(self.max_cells, self.max_rows, self.max_chars) < 1:
            raise ValueError("Budgets must be positive")

    def check_range(self, ref):
        r1, c1, r2, c2 = bounds(ref)
        if (r2 - r1 + 1) * (c2 - c1 + 1) > self.max_cells:
            raise BudgetExceeded("Range exceeds cell budget; narrow it or use a local query")


def compact(data: Any, budget: Budget | None = None, metrics: Metrics | None = None) -> str:
    """Never silently truncate an answer or a scalar. Caller must refine the query."""
    budget = budget or Budget()
    value = json.dumps(data, ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    if len(value) > budget.max_chars:
        raise BudgetExceeded(
            f"Response is {len(value)} characters; limit is {budget.max_chars}. Narrow the request."
        )
    if metrics:
        metrics.add("characters_exposed", len(value))
        metrics.add("estimated_tokens_exposed", math.ceil(len(value) / 4))
    return value
