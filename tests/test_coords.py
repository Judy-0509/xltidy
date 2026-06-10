from moa.coords import col_to_idx, idx_to_col, a1_to_rc, rc_to_a1, parse_range


def test_roundtrip():
    assert col_to_idx("A") == 1 and col_to_idx("AA") == 27
    assert idx_to_col(27) == "AA"
    assert a1_to_rc("C5") == (5, 3)
    assert rc_to_a1(5, 3) == "C5"
    assert parse_range("B5:D8") == (5, 2, 8, 4)
    assert parse_range("A1") == (1, 1, 1, 1)
