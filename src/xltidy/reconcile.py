from __future__ import annotations

from dataclasses import dataclass, field
from numbers import Number

from .coords import col_to_idx, parse_range
from .models import CellGrid
from .spec import TableSpec


@dataclass
class ReconcileReport:
    ok: bool = True
    issues: list[str] = field(default_factory=list)


def _close(a: float, b: float) -> bool:
    return abs(a - b) <= max(1e-6, 1e-4 * abs(b))


def reconcile_table(grid: CellGrid, table: TableSpec) -> list[str]:
    if not (table.totals and table.region and table.header and table.value_block):
        return []
    issues: list[str] = []
    r1, c1, r2, c2 = parse_range(f"{table.region.start}:{table.region.end}")
    header_row = max(table.header.rows) if table.header.rows else r1
    data_start = header_row + 1
    idx_col = col_to_idx(table.index_columns[0].col)
    vb_start = col_to_idx(table.value_block.cols[0])
    vb_end = col_to_idx(table.value_block.cols[-1])
    labels = {t.label for t in table.totals}

    subtotal_rows = [row for row in range(data_start, r2 + 1)
                     if str(grid.value_filled(row, idx_col)) in labels]
    component_rows = [row for row in range(data_start, r2 + 1)
                      if str(grid.value_filled(row, idx_col)) not in labels]
    for srow in subtotal_rows:
        label = str(grid.value_filled(srow, idx_col))
        for vc in range(vb_start, vb_end + 1):
            reported = grid.at(srow, vc)
            if not isinstance(reported, Number):
                continue
            total = sum(grid.at(row, vc) for row in component_rows if isinstance(grid.at(row, vc), Number))
            if not _close(float(reported), float(total)):
                issues.append(f"table {table.name} col {vc} subtotal '{label}': reported {reported} != sum {total}")
    return issues


def reconcile_pivot(data_cells_sum: float, grand_total: float | None, table_name: str) -> list[str]:
    if grand_total is None:
        return []
    if not _close(float(data_cells_sum), float(grand_total)):
        return [f"pivot {table_name}: data sum {data_cells_sum} != grand total {grand_total}"]
    return []
