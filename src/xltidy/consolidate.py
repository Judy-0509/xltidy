from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from .apply import apply_workbook
from .coords import a1_to_rc, parse_range
from .models import CellGrid
from .spec import TemplateSpec


@dataclass
class ConsolidateResult:
    tables: dict[str, pd.DataFrame] = field(default_factory=dict)
    drift_by_file: dict[str, list[str]] = field(default_factory=dict)


def detect_drift(spec: TemplateSpec, *, available_sheets: list[str],
                 grid_by_sheet: dict[str, CellGrid]) -> list[str]:
    issues: list[str] = []
    for sheet in spec.sheets:
        skey = str(sheet.sheet_match.value)
        if sheet.sheet_match.by == "name" and skey not in available_sheets:
            issues.append(f"selected sheet '{skey}' missing/renamed (available: {available_sheets})")
            continue
        grid = grid_by_sheet.get(skey)
        for t in sheet.tables:
            if t.kind != "table" or grid is None:
                continue
            if t.region:
                r1, c1, r2, c2 = parse_range(f"{t.region.start}:{t.region.end}")
                if r2 > grid.n_rows or c2 > grid.n_cols:
                    issues.append(f"{t.name}: region out of bounds")
            for cs in t.column_semantics:
                if cs.source_text is None:
                    continue
                r, c = a1_to_rc(cs.source)
                if str(grid.value_filled(r, c)).strip() != cs.source_text.strip():
                    issues.append(f"{t.name}: header {cs.source} expected '{cs.source_text}' got "
                                  f"'{grid.value_filled(r, c)}' (renamed/moved)")
    return issues


def consolidate(files: list[str], spec: TemplateSpec, *, list_sheets_fn,
                sheet_extractor, pivot_extractor, on_drift: str = "stop") -> ConsolidateResult:
    acc: dict[str, list[pd.DataFrame]] = {}
    drift_by_file: dict[str, list[str]] = {}
    for path in files:
        available = [s.name for s in list_sheets_fn(path)]
        # 표 시트만 드리프트용으로 추출
        grid_by_sheet: dict[str, CellGrid] = {}
        for sheet in spec.sheets:
            skey = str(sheet.sheet_match.value)
            if any(t.kind == "table" for t in sheet.tables) and skey in available:
                grid_by_sheet[skey] = sheet_extractor(path, sheet.sheet_match.value)
        drift = detect_drift(spec, available_sheets=available, grid_by_sheet=grid_by_sheet)
        if drift:
            drift_by_file[path] = drift
            if on_drift == "stop":
                continue
        res = apply_workbook(path, spec, sheet_extractor=sheet_extractor,
                             pivot_extractor=pivot_extractor, list_sheets_fn=list_sheets_fn,
                             filename=path)
        for name, frame in res.tables.items():
            acc.setdefault(name, []).append(frame)
    tables = {name: pd.concat(frames, ignore_index=True) for name, frames in acc.items()}
    return ConsolidateResult(tables=tables, drift_by_file=drift_by_file)
