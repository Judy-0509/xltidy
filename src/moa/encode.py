from __future__ import annotations

from numbers import Number

from .coords import idx_to_col, rc_to_a1
from .models import CellGrid

# 표 사이 빈 행이 이만큼 넘게 벌어지면 별개 영역(밴드)으로 본다.
# 표 안의 한 줄 띄움(간격 2)은 같은 밴드로 유지한다.
_BAND_GAP = 2
_FORMULA_MAX = 60


def _render(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "#bool"
    if isinstance(value, Number):
        return "#num"  # I2: 값 전사 금지
    text = str(value).strip()
    return f'"{text}"' if text else ""


def _render_cell(cell) -> str:
    r = _render(cell.value)
    if r == "#num" and cell.formula:
        # 수식 머리만 노출 — 합계행/파생열을 추측이 아니라 확정으로 잡게 한다.
        # (값은 여전히 전사하지 않는다: 수식 텍스트는 구조 정보다.)
        f = cell.formula
        r = f"#num{f[:_FORMULA_MAX]}{'...' if len(f) > _FORMULA_MAX else ''}"
    return r


def _bands(data_rows: list[int]) -> list[list[int]]:
    """값 행 목록을 빈 행 구간으로 분리된 연속 밴드들로 나눈다."""
    bands: list[list[int]] = []
    for row in data_rows:
        if bands and row - bands[-1][-1] <= _BAND_GAP:
            bands[-1].append(row)
        else:
            bands.append([row])
    return bands


def encode(grid: CellGrid, *, head_rows: int = 40, tail_rows: int = 10,
           head_cols: int = 40, tail_cols: int = 8) -> str:
    """시트를 LLM용 텍스트로 인코딩. 값(숫자)은 #num으로 가린다.

    - 구조 힌트(LISTOBJECT/PIVOT/NAME)와 열 서식(COL FORMATS)을 머리에 싣는다.
    - 빈 행으로 분리된 데이터 밴드를 찾아, 2개 이상이면 REGIONS 로 알린다
      (한 시트 다중 표). 행 샘플링은 시트 전체가 아니라 **밴드별로** 적용해
      중간에서 시작하는 두 번째 표의 머리가 잘리지 않게 한다.
    - 행: 밴드별로 head `head_rows` + tail `tail_rows`만 싣고 중간은
      `... omitted ...`. head_rows<=0 이면 행 캡 전체 해제, tail_rows<=0 이면
      head만 싣는다(끝은 DATA ROWS 로 알림).
    - 열: 값 있는 열이 head_cols+tail_cols 를 넘으면 같은 방식으로 자르고
      DATA COLS 로 알린다. head_cols<=0 이면 열 캡 해제.
    """
    lines = [f"SHEET: {grid.sheet} (rows 1..{grid.n_rows}, cols 1..{grid.n_cols})"]
    if grid.merged:
        lines.append("MERGED: " + ", ".join(
            f"{rc_to_a1(m.r1, m.c1)}:{rc_to_a1(m.r2, m.c2)}" for m in grid.merged))
    lines.extend(grid.hints)

    by_row: dict[int, list] = {}
    for cell in sorted(grid.cells, key=lambda c: (c.row, c.col)):
        by_row.setdefault(cell.row, []).append(cell)
    data_rows = sorted(by_row)  # 값이 실제로 있는 행만 = 진짜 데이터 범위
    data_cols = sorted({c.col for c in grid.cells})

    # 열 샘플링
    shown_cols: set[int] | None = None
    if head_cols > 0 and len(data_cols) > head_cols + max(tail_cols, 0):
        kept = data_cols[:head_cols] + (data_cols[-tail_cols:] if tail_cols > 0 else [])
        shown_cols = set(kept)
        lines.append(
            f"DATA COLS: {idx_to_col(data_cols[0])}..{idx_to_col(data_cols[-1])} "
            f"({len(data_cols)} value cols; {len(data_cols) - len(kept)} omitted)")

    if grid.col_formats:
        fmts = ", ".join(f"{idx_to_col(c)}:{f}" for c, f in sorted(grid.col_formats.items())
                         if shown_cols is None or c in shown_cols)
        if fmts:
            lines.append(f"COL FORMATS: {fmts}")

    bands = _bands(data_rows)
    if len(bands) > 1:
        regions = []
        for band in bands:
            cols = [c.col for r in band for c in by_row[r]]
            regions.append(f"{rc_to_a1(band[0], min(cols))}:{rc_to_a1(band[-1], max(cols))} "
                           f"({len(band)} value rows)")
        lines.append("REGIONS: " + ", ".join(regions))

    total_hidden = sum(max(len(b) - head_rows - max(tail_rows, 0), 0) for b in bands) \
        if head_rows > 0 else 0
    if total_hidden:
        lines.append(
            f"DATA ROWS: {data_rows[0]}..{data_rows[-1]} "
            f"({len(data_rows)} value rows; first {head_rows} + last {tail_rows} "
            f"per region shown, {total_hidden} omitted)")

    for band in bands:
        sampled = head_rows > 0 and len(band) > head_rows + max(tail_rows, 0)
        if sampled:
            tail = band[-tail_rows:] if tail_rows > 0 else []
            shown = band[:head_rows] + tail
            marker_after = band[head_rows - 1]
            gap_first = band[head_rows]
            gap_last = band[-tail_rows - 1] if tail_rows > 0 else band[-1]
            n_hidden = len(band) - len(shown)
        else:
            shown, marker_after = band, None
        for row in shown:
            for cell in by_row[row]:
                if shown_cols is not None and cell.col not in shown_cols:
                    continue
                r = _render_cell(cell)
                if r:
                    lines.append(f"{rc_to_a1(cell.row, cell.col)}\t{r}")
            if row == marker_after:
                lines.append(f"... rows {gap_first}..{gap_last} omitted ({n_hidden} value rows) ...")
    return "\n".join(lines)
