from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, PrivateAttr


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
    # Excel이 이미 알고 있는 구조(ListObject/피벗 범위/명명 범위) 한 줄씩.
    # encode가 그대로 LLM 프롬프트에 싣는다 — 추론이 아니라 확정 정보.
    hints: list[str] = []
    # 열 idx -> 대표 number format (General 제외). 값 열의 의미(%, 날짜, 통화)
    # 구분용. 열당 1셀만 샘플링하므로 혼합 서식 열은 마지막 값 셀 기준.
    col_formats: dict[int, str] = {}

    # Lazily-built (row, col) -> value lookup, cached so at()/value_filled are O(1).
    # Built once on first access. Do NOT mutate `cells` after the first lookup
    # (the tool builds a grid once via extract and only reads it); call
    # invalidate() if you must.
    _lookup: dict[tuple[int, int], Any] | None = PrivateAttr(default=None)

    def _index(self) -> dict[tuple[int, int], Any]:
        if self._lookup is None:
            self._lookup = {(c.row, c.col): c.value for c in self.cells}
        return self._lookup

    def invalidate(self) -> None:
        self._lookup = None

    def model_copy(self, **kwargs):  # type: ignore[override]
        copied = super().model_copy(**kwargs)
        copied._lookup = None  # never carry a stale index into a copy
        return copied

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
