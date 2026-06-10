from __future__ import annotations

import xlwings as xw

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


def list_sheets(path: str) -> list[SheetInfo]:
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(path, read_only=True, update_links=False)
        try:
            infos: list[SheetInfo] = []
            for i, sht in enumerate(wb.sheets, start=1):
                used = sht.used_range
                try:
                    ur, uc = used.last_cell.row, used.last_cell.column
                    if used.value is None:
                        ur = uc = 0
                except Exception:
                    ur = uc = 0
                infos.append(SheetInfo(
                    name=sht.name, index=i,
                    visibility=_VIS.get(int(sht.api.Visible), "visible"),
                    used_rows=ur, used_cols=uc,
                    n_pivots=int(sht.api.PivotTables().Count)))
            return infos
        finally:
            wb.close()
    finally:
        app.quit()


def extract(path: str, sheet: str | int | None = None) -> CellGrid:
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(path, read_only=True, update_links=False)
        try:
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
                            cells=cells, merged=_read_merged(sht))
        finally:
            wb.close()
    finally:
        app.quit()


def _read_merged(sht) -> list[MergedRange]:
    merged: list[MergedRange] = []
    seen: set[tuple[int, int, int, int]] = set()
    for cell in sht.used_range:
        if cell.api.MergeCells:
            a = cell.api.MergeArea
            r1, c1 = a.Row, a.Column
            key = (r1, c1, r1 + a.Rows.Count - 1, c1 + a.Columns.Count - 1)
            if key not in seen and (key[0], key[1]) != (key[2], key[3]):
                seen.add(key)
                merged.append(MergedRange(r1=key[0], c1=key[1], r2=key[2], c2=key[3]))
    return merged
