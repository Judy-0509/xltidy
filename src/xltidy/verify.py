from __future__ import annotations

import random
from numbers import Number

import pandas as pd

from .apply import _metric_name
from .coords import col_to_idx, parse_range
from .models import CellGrid
from .spec import TableSpec

_MISSING = object()


def _eq(a, b) -> bool:
    if isinstance(a, Number) and isinstance(b, Number):
        return abs(float(a) - float(b)) <= max(1e-6, 1e-4 * abs(float(b)))
    return a == b


def verify_table(grid: CellGrid, table: TableSpec, frame: pd.DataFrame, *,
                 period: str | None, sample: int | None = 50, seed: int = 0) -> list[str]:
    """Independent output verification for a kind=table extraction.

    Two checks, both useful even when the sheet has no subtotals (reconcile
    can't help there):
      1. count check  -- len(frame) == (#data rows excl. subtotals) x (#value cols)
      2. sample round-trip -- pick `sample` source value-cells at random and
         assert each landed in the output under the right (dims, metric, period).
         sample=None verifies every cell.

    Returns a list of human-readable issues ([] == all good).
    """
    if table.kind != "table" or not (table.region and table.header and table.value_block):
        return []
    issues: list[str] = []
    r1, c1, r2, c2 = parse_range(f"{table.region.start}:{table.region.end}")
    header_row = max(table.header.rows) if table.header.rows else r1
    data_start = header_row + 1
    idx_cols = [(col_to_idx(ic.col), ic.name) for ic in table.index_columns]
    vb = list(range(col_to_idx(table.value_block.cols[0]),
                    col_to_idx(table.value_block.cols[-1]) + 1))
    sub = {t.label for t in table.totals}
    idx0 = col_to_idx(table.index_columns[0].col) if table.index_columns else None

    data_rows = [r for r in range(data_start, r2 + 1)
                 if idx0 is None or str(grid.value_filled(r, idx0)) not in sub]
    expected = len(data_rows) * len(vb)
    if len(frame) != expected:
        issues.append(f"{table.name}: row count {len(frame)} != expected {expected} "
                      f"(data_rows={len(data_rows)} x value_cols={len(vb)})")

    # Build a one-shot lookup of the output so each sampled cell is O(1).
    var, valn, pern = table.unpivot.var_name, table.unpivot.value_name, table.period.name
    lut: dict = {}
    for rec in frame.to_dict("records"):
        lut[(tuple(rec.get(name) for _, name in idx_cols), rec.get(var), rec.get(pern))] = rec.get(valn)

    pairs = [(r, vc) for r in data_rows for vc in vb]
    if sample is not None and sample < len(pairs):
        pairs = random.Random(seed).sample(pairs, sample)

    mism = 0
    for r, vc in pairs:
        key = (tuple(grid.value_filled(r, ci) for ci, _ in idx_cols),
               _metric_name(table, header_row, vc), period)
        got = lut.get(key, _MISSING)
        src = grid.at(r, vc)
        if got is _MISSING or not _eq(got, src):
            mism += 1
            if mism <= 5:
                shown = "MISSING" if got is _MISSING else repr(got)
                issues.append(f"{table.name}: source cell ({r},{vc})={src!r} "
                              f"missing/mismatched in output (got {shown})")
    if mism > 5:
        issues.append(f"{table.name}: ...and {mism - 5} more sampled cell mismatches")
    return issues
