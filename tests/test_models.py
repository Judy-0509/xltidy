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


def test_index_is_cached():
    # perf regression guard: the (row,col)->value index must be built once,
    # not rebuilt on every at() call (that was an O(N^2) timeout on big sheets).
    g = CellGrid(sheet="s", n_rows=2, n_cols=2, cells=[Cell(row=1, col=1, value=1)])
    assert g._index() is g._index()  # same cached dict object
    g.invalidate()
    assert g._index() is not None and g.at(1, 1) == 1


def test_large_grid_lookup_correct():
    # 5000 rows x 4 cols = 20000 cells; with the cache this is instant.
    cells = [Cell(row=r, col=c, value=r * 10 + c)
             for r in range(1, 5001) for c in range(1, 5)]
    g = CellGrid(sheet="big", n_rows=5000, n_cols=4, cells=cells)
    assert g.at(1, 1) == 11
    assert g.at(5000, 4) == 50004
    assert g.at(2500, 3) == 25003
