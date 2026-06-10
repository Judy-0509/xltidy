from __future__ import annotations

import re
from dataclasses import dataclass, field

import pandas as pd

from .coords import a1_to_rc, col_to_idx, parse_range, rc_to_a1
from .models import CellGrid
from .reconcile import ReconcileReport, reconcile_pivot, reconcile_table
from .spec import PeriodSpec, TableSpec, TemplateSpec


@dataclass
class ApplyResult:
    tables: dict[str, pd.DataFrame] = field(default_factory=dict)
    reconcile: ReconcileReport = field(default_factory=ReconcileReport)
    drift: list[str] = field(default_factory=list)
    verify: list[str] = field(default_factory=list)


def resolve_period(period: PeriodSpec, filename: str, grid: CellGrid | None = None) -> str | None:
    src = period.source
    if src.from_ == "filename" and src.pattern:
        m = re.search(src.pattern, filename or "")
        if m:
            return "-".join(g for g in m.groups() if g is not None)
    if src.cell and grid is not None:
        r, c = a1_to_rc(src.cell)
        v = grid.value_filled(r, c)
        if v is not None:
            return str(v)
    return None


def _metric_name(table: TableSpec, header_row: int, col: int) -> str:
    a1 = rc_to_a1(header_row, col)
    for cs in table.column_semantics:
        if cs.source == a1:
            return cs.name
    return a1


def apply_table(grid: CellGrid, table: TableSpec, *, period: str | None) -> pd.DataFrame:
    assert table.region and table.header and table.value_block
    r1, c1, r2, c2 = parse_range(f"{table.region.start}:{table.region.end}")
    header_row = max(table.header.rows) if table.header.rows else r1
    data_start = header_row + 1
    idx_cols = [(col_to_idx(ic.col), ic.name) for ic in table.index_columns]
    vb_start, vb_end = col_to_idx(table.value_block.cols[0]), col_to_idx(table.value_block.cols[-1])
    subtotal_labels = {t.label for t in table.totals}
    idx_name0 = table.index_columns[0].name if table.index_columns else None

    records: list[dict] = []
    for row in range(data_start, r2 + 1):
        dims = {name: grid.value_filled(row, ci) for ci, name in idx_cols}  # 병합 라벨 해소 (I5)
        if idx_name0 and str(dims.get(idx_name0)) in subtotal_labels:
            continue  # 소계행 제외
        for vc in range(vb_start, vb_end + 1):
            rec = dict(dims)
            rec[table.unpivot.var_name] = _metric_name(table, header_row, vc)
            rec[table.unpivot.value_name] = grid.at(row, vc)
            rec[table.period.name] = period
            records.append(rec)
    return pd.DataFrame.from_records(records)


def finalize_pivot(raw: pd.DataFrame, table: TableSpec, *, period: str | None) -> pd.DataFrame:
    out = raw.copy()
    out[table.period.name] = period  # 피벗은 COM이 이미 long; period만 부착
    return out


def apply_workbook(path: str, spec: TemplateSpec, *, sheet_extractor=None,
                   pivot_extractor=None, list_sheets_fn=None,
                   period: str | None = None, filename: str | None = None,
                   verify: bool = False, verify_sample: int | None = 50) -> ApplyResult:
    if sheet_extractor is None:
        from .extract import extract as sheet_extractor
    if pivot_extractor is None:
        from .pivot import extract_pivot as pivot_extractor
    fname = filename if filename is not None else path

    tables: dict[str, pd.DataFrame] = {}
    issues: list[str] = []
    verify_issues: list[str] = []
    for sheet in spec.sheets:
        skey = sheet.sheet_match.value
        for table in sheet.tables:
            if table.kind == "table":
                grid = sheet_extractor(path, skey)
                pval = period if period is not None else resolve_period(table.period, fname, grid)
                frame = apply_table(grid, table, period=pval)
                tables[table.name] = frame
                issues += reconcile_table(grid, table)
                if verify:
                    from .verify import verify_table
                    verify_issues += verify_table(grid, table, frame, period=pval, sample=verify_sample)
            else:  # pivot
                raw, gt = pivot_extractor(path, skey, table.pivot_name)
                pval = period if period is not None else resolve_period(table.period, fname, None)
                frame = finalize_pivot(raw, table, period=pval)
                tables[table.name] = frame
                issues += reconcile_pivot(float(frame["value"].sum()), gt, table.name)
    return ApplyResult(tables=tables, reconcile=ReconcileReport(ok=not issues, issues=issues),
                       verify=verify_issues)
