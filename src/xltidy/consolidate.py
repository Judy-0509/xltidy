from __future__ import annotations

import contextlib
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
    verify_by_file: dict[str, list[str]] = field(default_factory=dict)
    period_issues: list[str] = field(default_factory=list)


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


def consolidate(files: list[str], spec: TemplateSpec, *, session_factory=None,
                list_sheets_fn=None, sheet_extractor=None, pivot_extractor=None,
                on_drift: str = "stop", verify: bool = False,
                verify_sample: int | None = 50) -> ConsolidateResult:
    """모든 버전을 같은 스펙으로 적용해 period 축으로 누적한 하나의 DB.

    **파일당 워크북을 1회만 연다**: session_factory(path)가 컨텍스트매니저로 세션을
    주면 그 한 세션을 드리프트 검사와 적용에 함께 재사용한다. session_factory 가
    없으면(테스트/그리드) 주입된 path 기반 추출기를 FnSession 으로 감싼다.
    """
    if on_drift not in ("stop", "continue"):
        raise ValueError(f"on_drift must be 'stop' or 'continue', got {on_drift!r}")
    from .session import FnSession

    pnames = {t.period.name for s in spec.sheets for t in s.tables}
    acc: dict[str, list[pd.DataFrame]] = {}
    drift_by_file: dict[str, list[str]] = {}
    verify_by_file: dict[str, list[str]] = {}
    period_by_file: dict[str, set] = {}

    for path in files:
        if session_factory is not None:
            cm = session_factory(path)
        else:
            cm = contextlib.nullcontext(
                FnSession(path, list_sheets_fn, sheet_extractor, pivot_extractor))
        with cm as sess:
            available = [s.name for s in sess.sheet_infos()]
            # 표 시트만 드리프트용으로 추출(세션 캐시 → 적용 단계가 재사용)
            grid_by_sheet: dict[str, CellGrid] = {}
            for sheet in spec.sheets:
                skey = str(sheet.sheet_match.value)
                if any(t.kind == "table" for t in sheet.tables) and skey in available:
                    grid_by_sheet[skey] = sess.grid(sheet.sheet_match.value)
            drift = detect_drift(spec, available_sheets=available, grid_by_sheet=grid_by_sheet)
            if drift:
                drift_by_file[path] = drift
                if on_drift == "stop":
                    continue
            res = apply_workbook(path, spec, session=sess, filename=path,
                                 verify=verify, verify_sample=verify_sample)
            if res.verify:
                verify_by_file[path] = res.verify
            pers: set = set()
            for frame in res.tables.values():
                for pn in pnames:
                    if pn in frame.columns:
                        pers |= set(pd.unique(frame[pn].dropna()))
                        if frame[pn].isna().any():
                            pers.add(None)
            period_by_file[path] = pers
            for name, frame in res.tables.items():
                acc.setdefault(name, []).append(frame)

    tables = {name: pd.concat(frames, ignore_index=True) for name, frames in acc.items()}
    return ConsolidateResult(tables=tables, drift_by_file=drift_by_file,
                             verify_by_file=verify_by_file,
                             period_issues=_period_collisions(period_by_file))


def _period_collisions(period_by_file: dict[str, set]) -> list[str]:
    """버전 구분 실패를 잡아낸다: 기간 미해결(None) 또는 서로 다른 파일이 같은 기간.

    드리프트와 동일하게 '검토 필요' 신호로 다뤄야 한다(CLI는 exit 2). 그렇지 않으면
    월/분기 버전이 조용히 같은 period 로 합쳐져 합본이 애매해진다.
    """
    issues: list[str] = []
    for path, pers in period_by_file.items():
        if (None in pers) or (not pers):
            issues.append(f"{path}: period could not be resolved (None) -> versions will be "
                          f"indistinguishable; fix the period pattern/cell")
    seen: dict = {}
    for path, pers in period_by_file.items():
        for p in pers:
            if p is None:
                continue
            if p in seen and seen[p] != path:
                issues.append(f"period {p!r} produced by multiple files ({seen[p]} and {path}) "
                              f"-> indistinguishable in the combined DB; make each version's "
                              f"filename/period yield a unique value")
            else:
                seen[p] = path
    return issues
