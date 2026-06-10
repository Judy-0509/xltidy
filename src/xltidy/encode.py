from __future__ import annotations

from numbers import Number

from .coords import rc_to_a1
from .models import CellGrid


def _render(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "#bool"
    if isinstance(value, Number):
        return "#num"  # I2: 값 전사 금지
    text = str(value).strip()
    return f'"{text}"' if text else ""


def encode(grid: CellGrid) -> str:
    lines = [f"SHEET: {grid.sheet} (rows 1..{grid.n_rows}, cols 1..{grid.n_cols})"]
    if grid.merged:
        lines.append("MERGED: " + ", ".join(
            f"{rc_to_a1(m.r1, m.c1)}:{rc_to_a1(m.r2, m.c2)}" for m in grid.merged))
    for cell in sorted(grid.cells, key=lambda c: (c.row, c.col)):
        r = _render(cell.value)
        if r:
            lines.append(f"{rc_to_a1(cell.row, cell.col)}\t{r}")
    return "\n".join(lines)
