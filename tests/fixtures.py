import pandas as pd

from xltidy.models import Cell, CellGrid


def sample_grid() -> CellGrid:
    """B5:D8. B=산업, C5/D5=월 헤더, 8행=합계 소계."""
    cells = [
        Cell(row=5, col=2, value="산업"),
        Cell(row=5, col=3, value="2024년 1월"),
        Cell(row=5, col=4, value="2024년 2월"),
        Cell(row=6, col=2, value="제조업"), Cell(row=6, col=3, value=100.0), Cell(row=6, col=4, value=110.0),
        Cell(row=7, col=2, value="서비스업"), Cell(row=7, col=3, value=200.0), Cell(row=7, col=4, value=210.0),
        Cell(row=8, col=2, value="합계"), Cell(row=8, col=3, value=300.0), Cell(row=8, col=4, value=320.0),
    ]
    return CellGrid(sheet="데이터", n_rows=8, n_cols=4, cells=cells)


def sample_pivot_raw() -> tuple[pd.DataFrame, float]:
    """extract_pivot가 돌려줄 모양: 이미 long (소계/총합계 제외). (frame, grand_total)."""
    frame = pd.DataFrame({
        "region": ["서울", "서울", "부산", "부산"],
        "field": ["인구", "가구", "인구", "가구"],
        "value": [10.0, 4.0, 20.0, 8.0],
    })
    return frame, 42.0  # 10+4+20+8
