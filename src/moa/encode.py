from __future__ import annotations

from numbers import Number

from .coords import rc_to_a1
from .models import CellGrid


def _render(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "#bool"
    if isinstance(value, Number):
        return "#num"  # I2: 값 전사 금지
    text = str(value).strip()
    return f'"{text}"' if text else ""


def encode(grid: CellGrid, *, head_rows: int = 40, tail_rows: int = 10) -> str:
    """시트를 LLM용 텍스트로 인코딩. 값(숫자)은 #num으로 가린다.

    대형 시트는 구조 추론에 전 행이 필요 없으므로 **값이 있는 행** 기준으로
    머리 `head_rows` + 꼬리 `tail_rows`만 싣고 중간은 생략한다("하단"은
    used_range의 유령 행이 아니라 실제 값이 있는 마지막 행으로 정의). 생략 시
    `DATA ROWS: a..b` 와 `... omitted ...` 를 남겨 LLM이 region 끝 행을 보이는
    마지막 행이 아니라 b로 잡도록 한다. head_rows/tail_rows<=0 이면 캡 해제.
    """
    lines = [f"SHEET: {grid.sheet} (rows 1..{grid.n_rows}, cols 1..{grid.n_cols})"]
    if grid.merged:
        lines.append("MERGED: " + ", ".join(
            f"{rc_to_a1(m.r1, m.c1)}:{rc_to_a1(m.r2, m.c2)}" for m in grid.merged))

    by_row: dict[int, list] = {}
    for cell in sorted(grid.cells, key=lambda c: (c.row, c.col)):
        by_row.setdefault(cell.row, []).append(cell)
    data_rows = sorted(by_row)  # 값이 실제로 있는 행만 = 진짜 데이터 범위

    sampling = head_rows > 0 and tail_rows > 0 and len(data_rows) > head_rows + tail_rows
    if sampling:
        shown = data_rows[:head_rows] + data_rows[-tail_rows:]
        gap_after = data_rows[head_rows - 1]                      # 마지막 머리 행
        gap_first, gap_last = data_rows[head_rows], data_rows[-tail_rows - 1]
        n_hidden = len(data_rows) - head_rows - tail_rows
        lines.append(
            f"DATA ROWS: {data_rows[0]}..{data_rows[-1]} "
            f"({len(data_rows)} value rows; first {head_rows} + last {tail_rows} shown, "
            f"{n_hidden} omitted)")
    else:
        shown, gap_after = data_rows, None

    for row in shown:
        for cell in by_row[row]:
            r = _render(cell.value)
            if r:
                lines.append(f"{rc_to_a1(cell.row, cell.col)}\t{r}")
        if row == gap_after:  # 머리/꼬리 경계에 생략 표시 1회
            lines.append(f"... rows {gap_first}..{gap_last} omitted ({n_hidden} value rows) ...")
    return "\n".join(lines)
