from xltidy.models import Cell, CellGrid, MergedRange, SheetInfo


def test_at_and_roundtrip():
    g = CellGrid(sheet="데이터", n_rows=8, n_cols=4,
                 cells=[Cell(row=6, col=2, value="제조업")])
    assert g.at(6, 2) == "제조업"
    assert g.at(7, 7) is None
    g2 = CellGrid.model_validate_json(g.model_dump_json())
    assert g2.at(6, 2) == "제조업"


def test_value_filled_resolves_vertical_merge():
    g = CellGrid(sheet="s", n_rows=9, n_cols=2,
                 cells=[Cell(row=6, col=2, value="제조업")],
                 merged=[MergedRange(r1=6, c1=2, r2=9, c2=2)])
    assert g.value_filled(6, 2) == "제조업"  # 앵커
    assert g.value_filled(8, 2) == "제조업"  # 병합 내부
    assert g.value_filled(1, 1) is None


def test_sheet_info_fields():
    si = SheetInfo(name="요약", index=2, visibility="hidden", used_rows=10, used_cols=5, n_pivots=1)
    assert si.visibility == "hidden" and si.n_pivots == 1
