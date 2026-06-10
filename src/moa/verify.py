from __future__ import annotations

import math
import random
from collections import Counter
from numbers import Number

import pandas as pd

from .apply import _metric_name
from .coords import col_to_idx, parse_range
from .models import CellGrid
from .spec import TableSpec


def _norm(v):
    """Normalize a value for use as a Counter key (np.float64 -> float, NaN -> None)."""
    if isinstance(v, bool):
        return v
    if isinstance(v, Number):
        f = float(v)
        return None if math.isnan(f) else f
    return v


def verify_table(grid: CellGrid, table: TableSpec, frame: pd.DataFrame, *,
                 period: str | None, sample: int | None = 50, seed: int = 0) -> list[str]:
    """Independent output verification for a kind=table extraction.

    Two checks, both useful even when the sheet has no subtotals (reconcile
    can't help there):
      1. count check       -- len(frame) == (#data rows excl. subtotals) x (#value cols)
      2. sample round-trip -- pick `sample` source value-cells at random and assert
         each landed in the output as a row with the right (dims, metric, period,
         value). Uses a multiset (Counter) so **duplicate index labels** (e.g. a
         vertically merged category spanning rows) are handled correctly rather
         than collapsing. sample=None (or <=0) verifies every cell.

    Note: merged *numeric body* cells are not supported -- value cells are read
    per-cell; a merged numeric region yields None for interior cells.

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

    # Multiset of full output rows; membership handles duplicate dimension keys.
    var, valn, pern = table.unpivot.var_name, table.unpivot.value_name, table.period.name
    actual: Counter = Counter(
        (tuple(_norm(rec.get(name)) for _, name in idx_cols),
         _norm(rec.get(var)), _norm(rec.get(pern)), _norm(rec.get(valn)))
        for rec in frame.to_dict("records")
    )

    pairs = [(r, vc) for r in data_rows for vc in vb]
    if sample is not None and sample > 0 and sample < len(pairs):
        pairs = random.Random(seed).sample(pairs, sample)

    mism = 0
    for r, vc in pairs:
        exp = (tuple(_norm(grid.value_filled(r, ci)) for ci, _ in idx_cols),
               _norm(_metric_name(table, header_row, vc)), _norm(period),
               _norm(grid.at(r, vc)))
        if actual.get(exp, 0) <= 0:
            mism += 1
            if mism <= 5:
                issues.append(f"{table.name}: source cell ({r},{vc})={grid.at(r, vc)!r} "
                              f"not found in output under {exp[:3]}")
    if mism > 5:
        issues.append(f"{table.name}: ...and {mism - 5} more sampled cell mismatches")
    return issues
