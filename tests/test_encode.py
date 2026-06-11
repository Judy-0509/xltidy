from tests.fixtures import sample_grid
from moa.encode import encode
from moa.models import Cell, CellGrid


def test_encode():
    out = encode(sample_grid())
    assert "SHEET: 데이터" in out and "rows 1..8" in out
    assert "B5" in out and "산업" in out
    assert "#num" in out  # 숫자값 마스킹 (I2)
    assert encode(sample_grid()) == out  # 결정론


def _flat_grid(n: int) -> CellGrid:
    """n행짜리 평평한 표: A열 라벨(문자), B열 값(#num)."""
    cells: list[Cell] = []
    for r in range(1, n + 1):
        cells.append(Cell(row=r, col=1, value=f"label{r}"))
        cells.append(Cell(row=r, col=2, value=float(r)))
    return CellGrid(sheet="big", n_rows=n, n_cols=2, cells=cells)


def test_no_cap_when_grid_fits():
    out = encode(_flat_grid(20), head_rows=40, tail_rows=10)
    assert "DATA ROWS" not in out and "omitted" not in out
    assert '"label20"' in out  # 전 행 표시


def test_caps_large_grid_head_and_tail():
    out = encode(_flat_grid(200), head_rows=40, tail_rows=10)
    # 진짜 끝(값 있는 마지막 행)을 명시해야 region.end를 잡을 수 있다
    assert "DATA ROWS: 1..200" in out
    assert "150 value rows" in out  # 200 - 40 - 10
    assert '"label40"' in out and '"label41"' not in out       # 머리 경계
    assert '"label191"' in out and '"label190"' not in out     # 꼬리 경계(200-10+1)


def test_cap_disabled_with_zero():
    out = encode(_flat_grid(200), head_rows=0, tail_rows=10)
    assert "omitted" not in out and '"label100"' in out


def test_bottom_is_value_rows_not_phantom_used_range():
    # n_rows(=used_range)가 유령 행으로 부풀려져 있어도, 표시되는 끝은
    # 값이 있는 마지막 행(50)이어야 한다.
    base = _flat_grid(50)
    g = CellGrid(sheet=base.sheet, n_rows=9999, n_cols=2, cells=base.cells)  # phantom bottom
    out = encode(g, head_rows=40, tail_rows=5)
    assert "DATA ROWS: 1..50" in out  # 9999가 아니라 50


def test_formula_shown_for_numeric_cells():
    # 합계행 수식은 구조 신호 → #num 뒤에 수식 머리를 노출(값 전사는 여전히 금지)
    g = CellGrid(sheet="s", n_rows=3, n_cols=1, cells=[
        Cell(row=1, col=1, value=1.0), Cell(row=2, col=1, value=2.0),
        Cell(row=3, col=1, value=3.0, formula="=SUM(A1:A2)")])
    out = encode(g)
    assert "A3\t#num=SUM(A1:A2)" in out
    assert "A1\t#num\n" in out  # 일반 숫자는 그대로 #num


def test_multi_band_regions_and_per_band_sampling():
    # 1..100행 표 + 200..299행 두 번째 표: REGIONS 로 두 영역을 알리고,
    # 행 샘플링은 밴드별이라 두 번째 표의 머리(200행)도 보여야 한다.
    cells = []
    for r in list(range(1, 101)) + list(range(200, 300)):
        cells.append(Cell(row=r, col=1, value=f"label{r}"))
        cells.append(Cell(row=r, col=2, value=float(r)))
    g = CellGrid(sheet="multi", n_rows=300, n_cols=2, cells=cells)
    out = encode(g, head_rows=20, tail_rows=5)
    assert "REGIONS: A1:B100 (100 value rows), A200:B299 (100 value rows)" in out
    assert '"label200"' in out      # 두 번째 밴드의 머리가 살아 있음
    assert '"label21"' not in out   # 첫 밴드 head 경계
    assert "DATA ROWS: 1..299" in out


def test_hints_and_col_formats_lines():
    g = CellGrid(sheet="s", n_rows=2, n_cols=3,
                 cells=[Cell(row=1, col=3, value=1.0)],
                 hints=["LISTOBJECT 표1: A1:C9", "PIVOT PT1: E1:G9"],
                 col_formats={3: "0.0%"})
    out = encode(g)
    assert "LISTOBJECT 표1: A1:C9" in out and "PIVOT PT1: E1:G9" in out
    assert "COL FORMATS: C:0.0%" in out


def test_col_cap_on_wide_grid():
    cells = [Cell(row=1, col=c, value=float(c)) for c in range(1, 101)]
    g = CellGrid(sheet="wide", n_rows=1, n_cols=100, cells=cells)
    out = encode(g, head_cols=10, tail_cols=2)
    assert "DATA COLS: A..CV (100 value cols; 88 omitted)" in out
    assert "J1\t#num" in out and "K1" not in out      # 머리 10열 경계
    assert "CU1\t#num" in out                           # 꼬리 2열


def test_tail_zero_keeps_head_cap():
    # --tail 0 = "head만 보내기". 예전엔 캡 전체가 풀렸다(회귀 방지).
    out = encode(_flat_grid(200), head_rows=40, tail_rows=0)
    assert '"label40"' in out and '"label41"' not in out
    assert "omitted" in out and "DATA ROWS: 1..200" in out
