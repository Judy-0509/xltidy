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


def _flatten_values(raw, count: int) -> list:
    """Range.Value 결과를 행 우선 1차원 리스트로. 단일 셀이면 스칼라가 온다."""
    if count == 1 and not isinstance(raw, (list, tuple)):
        return [raw]
    return [v for row in raw for v in (row if isinstance(row, (list, tuple)) else (row,))]


def _safe_name(obj) -> str:
    try:
        return str(obj.Name)
    except Exception:
        return "?"


def _cell_df_name(pc) -> str | None:
    """데이터 셀이 속한 데이터 필드 이름. PivotCell.DataField 가 정답이고
    (이 빌드에서 실측 확인), 없으면 PivotField 로 폴백(데이터 영역 셀에선 동일)."""
    try:
        return str(pc.DataField.Name)
    except Exception:
        try:
            return str(pc.PivotField.Name)
        except Exception:
            return None


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
            print(f"[moa] WARNING: pivot {pivot_name or 1} on sheet {sheet!r}: could not "
                  f"clear {len(failures)} filter(s) -> result may be a FILTERED subset, not "
                  f"all rows: {'; '.join(failures[:5])}")

    # 카운트/본문은 필터 해제 후 상태로 읽는다
    data_fields = [pt.DataFields.Item(i + 1).SourceName
                   for i in range(pt.DataFields.Count)]
    first_df_names: set[str] | None = None
    if pt.DataFields.Count > 1:
        # v1: 단일 데이터 필드만. 다중이면 첫 필드 소속 셀만 통과시킨다
        # (필터 없이 순회하면 두 필드의 값이 구분 없이 섞여 들어간다).
        print(f"[moa] WARNING: pivot has {pt.DataFields.Count} data fields "
              f"{data_fields}; v1 keeps the FIRST only — other fields' cells are "
              f"skipped. Result may be incomplete; prefer single-data-field pivots.")
        df1 = pt.DataFields.Item(1)
        first_df_names = {str(df1.Name)}
        try:
            first_df_names.add(str(df1.SourceName))
        except Exception:
            pass

    n_row_fields = int(pt.RowFields.Count)
    n_col_fields = int(pt.ColumnFields.Count)
    if pt.DataFields.Count > 1:
        # 다중 데이터 필드면 "Σ값" 의사필드가 행/열 축에 끼어 Count 를 1 올리지만,
        # PivotCell.RowItems/ColumnItems 에는 데이터 필드 항목이 빠진다(실측).
        # 보정하지 않으면 구조 판별이 모든 데이터 셀을 소계로 오인해 스킵한다.
        try:
            ori = int(pt.DataPivotField.Orientation)
            if ori == 1:    # xlRowField
                n_row_fields -= 1
            elif ori == 2:  # xlColumnField
                n_col_fields -= 1
        except Exception:
            pass

    records: list[dict] = []
    grand_total: float | None = None
    body = pt.DataBodyRange  # 값이 하나도 없는 피벗이면 None
    if body is None:
        return pd.DataFrame({"value": pd.Series(dtype=float)}), None
    # 본문 값은 1회 일괄 전송으로 읽는다. Range.Cells(i)와 Range.Value 의 평탄화
    # 순서는 둘 다 행 우선(row-major)이라 인덱스가 일치한다. PivotCell(라벨)은
    # 셀별 COM 접근이 불가피하지만, 값까지 셀별로 읽는 왕복은 이렇게 없앤다.
    body_values = _flatten_values(body.Value, int(body.Count))
    for i in range(1, body.Count + 1):
        cell = body.Cells(i)
        pc = cell.PivotCell
        if first_df_names is not None and _cell_df_name(pc) not in first_df_names:
            continue  # 다른 데이터 필드의 셀
        ric = pc.RowItems.Count
        cic = pc.ColumnItems.Count
        if ric == 0 and cic == 0:
            v = body_values[i - 1]
            grand_total = float(v) if v is not None else grand_total
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
        rec["value"] = body_values[i - 1]
        records.append(rec)

    frame = pd.DataFrame.from_records(records)
    if "value" not in frame.columns:  # 모든 셀이 스킵된 경우(빈/전부 소계)
        frame["value"] = pd.Series(dtype=float)
    return frame, grand_total


def extract_pivot(path: str, sheet: str | int, pivot_name: str | None = None,
                  *, clear_filters: bool = True):
    """단발성: 파일을 1회(쓰기 가능)으로 열어 피벗을 읽고 닫는다.

    필터 해제를 위해 쓰기 가능으로 열되 저장하지 않는다(디스크 원본 보존).
    """
    with open_book(path, read_only=False) as wb:
        return pivot_from_sheet(wb, sheet, pivot_name, clear_filters=clear_filters)
