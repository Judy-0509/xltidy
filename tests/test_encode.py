from tests.fixtures import sample_grid
from xltidy.encode import encode


def test_encode():
    out = encode(sample_grid())
    assert "SHEET: 데이터" in out and "rows 1..8" in out
    assert "B5" in out and "산업" in out
    assert "#num" in out  # 숫자값 마스킹 (I2)
    assert encode(sample_grid()) == out  # 결정론
