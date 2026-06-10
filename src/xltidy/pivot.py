from __future__ import annotations

import pandas as pd

from ._xl import open_book

# XlPivotCellType constants (read empirically from the COM type library on the
# dev machine): xlPivotCellValue=0, xlPivotCellSubtotal=2, xlPivotCellGrandTotal=3,
# xlPivotCellDataPivotField=8. The contract is to keep ONLY data cells and skip
# subtotal/grand-total cells.
#
# DEVIATION (documented): on this Excel build, ``cell.PivotCell.PivotCellType``
# returns 0 for EVERY cell (data, subtotal, and grand total) even after
# save+reopen and RefreshTable — the property is unusable through the win32com
# binding here. We therefore discriminate structurally, which is exactly
# equivalent to the PivotCellType semantics:
#   - data cell:  RowItems.Count == #row fields AND ColumnItems.Count == #col fields
#   - subtotal:   fewer row/col items than the field counts  -> skipped
#   - grand total: RowItems.Count == 0 AND ColumnItems.Count == 0 -> captured, skipped
# Label extraction still uses PivotCell.RowItems/.ColumnItems per the contract.


def _safe_name(obj) -> str:
    try:
        return str(obj.Name)
    except Exception:
        return "?"


def _coll(accessor):
    """xlwings 의 COMRetry 래퍼에서 일부 필드 접근자(PivotFields/PageFields)는 컬렉션이
    아니라 **호출해야 하는 메서드 래퍼**를 돌려준다. 컬렉션이면 그대로, 아니면 호출."""
    try:
        accessor.Count
        return accessor
    except Exception:
        return accessor()


def _clear_all_filters(pt, wb) -> list[str]:
    """피벗의 모든 필터를 해제해 **전체 데이터**를 읽도록 만든다.

    해제 대상: 보고서(페이지) 필터, 행/열 필드의 라벨·값 필터 및 **숨김 항목**,
    연결된 슬라이서·타임라인. 메모리상에서만 바꾸고 저장하지 않으므로 디스크 원본은
    그대로다(open_book 가 SaveChanges=False).

    best-effort 지만 **조용히 필터된 부분집합을 내보내지 않도록**, 해제에 실패한
    항목을 모아 반환한다(호출측이 크게 경고). 빈 필터에 ClearAllFilters 는 보통
    no-op 이지만 OLAP/계산필드/보호된 시트 등에서 예외가 날 수 있다.
    """
    failures: list[str] = []
    # 행/열/페이지 필드만: 라벨/값 필터 + 숨김 항목 해제
    # (데이터/숨김 필드는 필터 불가 → 건너뛰어 헛경보 방지)
    try:
        pfs = _coll(pt.PivotFields)
        for i in range(1, int(pfs.Count) + 1):
            pf = pfs.Item(i)
            try:
                if int(pf.Orientation) not in (1, 2, 3):  # xlRow/Column/Page 만
                    continue
            except Exception:
                pass
            try:
                pf.ClearAllFilters()
            except Exception as e:
                failures.append(f"field '{_safe_name(pf)}': {e}")
    except Exception as e:
        failures.append(f"PivotFields: {e}")
    # 페이지(보고서) 필터: 모든 항목 표시로 복귀(로캘/OLAP 대비 best-effort)
    try:
        pgs = _coll(pt.PageFields)
        for i in range(1, int(pgs.Count) + 1):
            try:
                pgs.Item(i).CurrentPage = "(All)"
            except Exception:
                pass
    except Exception:
        pass
    # 슬라이서/타임라인 캐시
    try:
        scs = _coll(wb.api.SlicerCaches)
        for i in range(1, int(scs.Count) + 1):
            try:
                scs.Item(i).ClearAllFilters()
            except Exception as e:
                failures.append(f"slicer[{i}]: {e}")
    except Exception as e:
        failures.append(f"SlicerCaches: {e}")
    try:
        pt.RefreshTable()
    except Exception as e:
        failures.append(f"RefreshTable: {e}")
    return failures


def pivot_from_sheet(wb, sheet: str | int, pivot_name: str | None = None,
                     *, clear_filters: bool = True):
    """이미 열린 워크북의 피벗을 tidy long으로. 반환 (frame, grand_total).

    clear_filters=True(기본): 읽기 전에 모든 필터를 해제해 **전체 데이터**를
    가져온다. False면 현재 화면에 보이는(필터 적용된) 상태 그대로 읽는다.
    단일 데이터 필드 한정(다중이면 경고 후 그대로 순회 — 알려진 v1 한계).
    """
    sht = wb.sheets[sheet]
    pts = sht.api.PivotTables()
    if pts.Count == 0:
        raise ValueError(f"no pivot table on sheet {sheet!r}")
    pt = pts.Item(pivot_name) if pivot_name else pts.Item(1)

    if clear_filters:
        failures = _clear_all_filters(pt, wb)
        if failures:
            print(f"[xltidy] WARNING: pivot {pivot_name or 1} on sheet {sheet!r}: could not "
                  f"clear {len(failures)} filter(s) -> result may be a FILTERED subset, not "
                  f"all rows: {'; '.join(failures[:5])}")

    # 카운트/본문은 필터 해제 후 상태로 읽는다
    data_fields = [pt.DataFields.Item(i + 1).SourceName
                   for i in range(pt.DataFields.Count)]
    if pt.DataFields.Count > 1:
        # v1: 단일 데이터 필드만. 다중은 명시적으로 알리고 첫 필드만.
        print(f"[xltidy] WARNING: pivot has {pt.DataFields.Count} data fields "
              f"{data_fields}; v1 uses the first only.")

    n_row_fields = int(pt.RowFields.Count)
    n_col_fields = int(pt.ColumnFields.Count)

    records: list[dict] = []
    grand_total: float | None = None
    body = pt.DataBodyRange
    for i in range(1, body.Count + 1):
        cell = body.Cells(i)
        pc = cell.PivotCell
        ric = pc.RowItems.Count
        cic = pc.ColumnItems.Count
        if ric == 0 and cic == 0:
            grand_total = float(cell.Value) if cell.Value is not None else grand_total
            continue  # 총합계 스킵
        if ric != n_row_fields or cic != n_col_fields:
            continue  # 소계 스킵 (행/열 필드 수보다 항목이 적음)
        rec: dict = {}
        for k in range(ric):
            item = pc.RowItems.Item(k + 1)
            pf = item.Parent  # PivotField
            rec[str(pf.Name)] = str(item.Name)
        col_label = None
        for k in range(cic):
            name = str(pc.ColumnItems.Item(k + 1).Name)
            col_label = name if col_label is None else f"{col_label}/{name}"
        if col_label is not None:
            rec["field"] = col_label
        rec["value"] = cell.Value
        records.append(rec)

    frame = pd.DataFrame.from_records(records)
    # 행필드 이름 정규화: 한국어 헤더 그대로 둠. 호출측이 사용.
    if "지역" in frame.columns:
        frame = frame.rename(columns={"지역": "region"})
    return frame, grand_total


def extract_pivot(path: str, sheet: str | int, pivot_name: str | None = None,
                  *, clear_filters: bool = True):
    """단발성: 파일을 1회(쓰기 가능)으로 열어 피벗을 읽고 닫는다.

    필터 해제를 위해 쓰기 가능으로 열되 저장하지 않는다(디스크 원본 보존).
    """
    with open_book(path, read_only=False) as wb:
        return pivot_from_sheet(wb, sheet, pivot_name, clear_filters=clear_filters)
