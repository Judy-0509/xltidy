from __future__ import annotations

import pandas as pd
import xlwings as xw

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


def extract_pivot(path: str, sheet: str | int, pivot_name: str | None = None):
    """피벗을 tidy long으로. 반환 (frame, grand_total). 단일 데이터 필드 한정."""
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(path, read_only=False, update_links=False)
        try:
            sht = wb.sheets[sheet]
            pts = sht.api.PivotTables()
            if pts.Count == 0:
                raise ValueError(f"no pivot table on sheet {sheet!r}")
            pt = pts.Item(pivot_name) if pivot_name else pts.Item(1)

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
        finally:
            wb.close()
    finally:
        app.quit()
