from __future__ import annotations

from ._xl import open_book
from .models import Cell, CellGrid, MergedRange, SheetInfo

_VIS = {-1: "visible", 0: "hidden", 2: "very_hidden"}


def _as_2d(value):
    if value is None:
        return [[None]]
    if not isinstance(value, (list, tuple)):
        return [[value]]
    value = [list(row) if isinstance(row, (list, tuple)) else row for row in value]
    if value and not isinstance(value[0], list):
        return [value]
    return value


def sheet_infos(wb) -> list[SheetInfo]:
    """이미 열린 워크북에서 모든 시트 정보(숨김·피벗수 포함)를 읽는다."""
    infos: list[SheetInfo] = []
    for i, sht in enumerate(wb.sheets, start=1):
        used = sht.used_range
        try:
            ur, uc = used.last_cell.row, used.last_cell.column
            # 빈 시트 판별: used_range가 A1 단일 셀일 때만 값 1개를 확인한다.
            # (예전처럼 used.value 를 무조건 읽으면 시트 전체 2D 배열이 COM으로
            # 전송돼 대형 시트에서 sheets 명령이 수십 초씩 걸린다.)
            if ur == 1 and uc == 1 and used.value is None:
                ur = uc = 0
        except Exception:
            ur = uc = 0
        infos.append(SheetInfo(
            name=sht.name, index=i,
            visibility=_VIS.get(int(sht.api.Visible), "visible"),
            used_rows=ur, used_cols=uc,
            n_pivots=int(sht.api.PivotTables().Count)))
    return infos


def grid_from_sheet(wb, sheet: str | int | None = None) -> CellGrid:
    """이미 열린 워크북의 한 시트를 CellGrid로 읽는다(값 없는 셀 제외)."""
    sht = wb.sheets.active if sheet is None else wb.sheets[sheet]
    used = sht.used_range
    n_rows, n_cols = used.last_cell.row, used.last_cell.column
    r0, c0 = used.row, used.column
    vals = _as_2d(used.value)
    fmls = _as_2d(used.formula)
    cells: list[Cell] = []
    for i, row_vals in enumerate(vals):
        for j, val in enumerate(row_vals):
            if val is None:
                continue
            f = fmls[i][j] if fmls else None
            cells.append(Cell(row=r0 + i, col=c0 + j, value=val,
                              formula=f if isinstance(f, str) and f.startswith("=") else None))
    return CellGrid(sheet=sht.name, n_rows=n_rows, n_cols=n_cols,
                    cells=cells, merged=_read_merged(sht, cells),
                    hints=_sheet_hints(sht), col_formats=_col_formats(sht, cells))


# col_formats 샘플링 상한: 열당 COM 1회라 보통 싸지만, 수백 열짜리 덤프에서
# 추출이 늘어지는 것은 막는다.
_MAX_FORMAT_COLS = 120


def _col_formats(sht, cells: list[Cell]) -> dict[int, str]:
    """열마다 **마지막 값 셀** 1개의 number format을 샘플링한다(General 제외).

    헤더가 아니라 본문 쪽 셀을 고르는 이유: 값 열의 의미(%, 날짜, 통화)를
    구분하는 게 목적이고, 마지막 행이 본문일 확률이 가장 높다.
    """
    last_by_col: dict[int, int] = {}
    for c in cells:
        if c.row > last_by_col.get(c.col, 0):
            last_by_col[c.col] = c.row
    formats: dict[int, str] = {}
    for col in sorted(last_by_col)[:_MAX_FORMAT_COLS]:
        try:
            fmt = sht.range((last_by_col[col], col)).number_format
        except Exception:
            continue
        if fmt and fmt != "General":
            formats[col] = str(fmt)
    return formats


def _sheet_hints(sht) -> list[str]:
    """Excel이 이미 알고 있는 구조를 한 줄씩 모은다(전부 best-effort).

    - ListObject(정식 Excel 표): 범위·헤더가 확정값 → 추론 불필요
    - 피벗 테이블 출력 범위: LLM이 그 위에 kind:table 을 그리는 것을 방지
    - 이 시트를 가리키는 명명 범위: 표 영역 후보
    """
    hints: list[str] = []
    try:
        los = sht.api.ListObjects
        for i in range(1, int(los.Count) + 1):
            lo = los.Item(i)
            hints.append(f"LISTOBJECT {lo.Name}: {str(lo.Range.Address).replace('$', '')}")
    except Exception:
        pass
    try:
        pts = sht.api.PivotTables()
        for i in range(1, int(pts.Count) + 1):
            pt = pts.Item(i)
            hints.append(f"PIVOT {pt.Name}: {str(pt.TableRange2.Address).replace('$', '')}")
    except Exception:
        pass
    try:
        for nm in sht.book.names:
            try:
                rng = nm.refers_to_range
                if rng is not None and rng.sheet.name == sht.name:
                    hints.append(f"NAME {nm.name}: {rng.address.replace('$', '')}")
            except Exception:
                continue  # 깨진 이름(#REF!) 등은 건너뜀
    except Exception:
        pass
    return hints


def list_sheets(path: str) -> list[SheetInfo]:
    """단발성: 파일을 1회 열어 시트 목록만 읽고 닫는다."""
    with open_book(path) as wb:
        return sheet_infos(wb)


def extract(path: str, sheet: str | int | None = None) -> CellGrid:
    """단발성: 파일을 1회 열어 한 시트를 읽고 닫는다."""
    with open_book(path) as wb:
        return grid_from_sheet(wb, sheet)


def _read_merged(sht, cells: list[Cell]) -> list[MergedRange]:
    """Enumerate merged ranges without scanning every cell over COM.

    Two speedups vs. a full per-cell scan (which is O(used cells) COM round
    trips and times out on large sheets):
      1. Fast path: if the whole used range reports MergeCells == False, there
         are no merges anywhere -> return [] immediately (the common case for
         big flat data dumps).
      2. Otherwise scan every non-empty cell EXCEPT the dense numeric body.
         Merge anchors in survey data are labels and headers -- strings, dates,
         booleans -- so we skip only plain int/float cells (the numeric body,
         which is never merged in practice). This catches date-valued year
         headers and numeric-coded labels while staying fast on big tables.
         Limitation: a merge anchored by a bare number is not auto-detected.
    """
    used = sht.used_range
    try:
        if used.api.MergeCells is False:
            return []
    except Exception:
        pass  # mixed/unknown -> fall through to the targeted scan

    merged: list[MergedRange] = []
    seen: set[tuple[int, int, int, int]] = set()
    for c in cells:
        if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
            continue  # skip dense numeric body for speed (not a merge anchor)
        api = sht.range((c.row, c.col)).api
        if api.MergeCells:
            a = api.MergeArea
            key = (a.Row, a.Column, a.Row + a.Rows.Count - 1, a.Column + a.Columns.Count - 1)
            if key not in seen and (key[0], key[1]) != (key[2], key[3]):
                seen.add(key)
                merged.append(MergedRange(r1=key[0], c1=key[1], r2=key[2], c2=key[3]))
    return merged
