from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel


class Cell(BaseModel):
    row: int
    col: int
    value: Any = None
    formula: str | None = None
    number_format: str | None = None


class MergedRange(BaseModel):
    r1: int
    c1: int
    r2: int
    c2: int


class CellGrid(BaseModel):
    sheet: str
    n_rows: int
    n_cols: int
    cells: list[Cell] = []
    merged: list[MergedRange] = []

    def _index(self) -> dict[tuple[int, int], Any]:
        return {(c.row, c.col): c.value for c in self.cells}

    def at(self, row: int, col: int) -> Any:
        return self._index().get((row, col))

    def value_filled(self, row: int, col: int) -> Any:
        """병합셀 내부 좌표는 앵커(좌상단) 값으로 해소 (I5).

        인덱스/라벨/헤더 읽기는 반드시 이 메서드를 써야 세로 병합된
        조사기관 표가 올바르게 풀린다.
        """
        direct = self.at(row, col)
        if direct is not None:
            return direct
        for m in self.merged:
            if m.r1 <= row <= m.r2 and m.c1 <= col <= m.c2:
                return self.at(m.r1, m.c1)
        return None


class SheetInfo(BaseModel):
    name: str
    index: int  # 1-based 워크북 내 위치
    visibility: Literal["visible", "hidden", "very_hidden"]
    used_rows: int
    used_cols: int
    n_pivots: int = 0
