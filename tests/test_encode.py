from tests.fixtures import sample_grid
from xltidy.encode import encode
from xltidy.models import Cell, CellGrid


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
