from tests.fixtures import sample_grid
from moa.apply import apply_table
from moa.spec import TemplateSpec, sample_spec_dict
from moa.verify import verify_table


def _table():
    return TemplateSpec.model_validate(sample_spec_dict()).sheets[0].tables[0]


def test_verify_ok_full_sample():
    t = _table()
    frame = apply_table(sample_grid(), t, version="2024Q1")
    # sample=None -> check every cell round-tripped
    assert verify_table(sample_grid(), t, frame, version="2024Q1", sample=None) == []


def test_verify_flags_count_mismatch():
    t = _table()
    frame = apply_table(sample_grid(), t, version="2024Q1").iloc[:2]  # drop rows
    issues = verify_table(sample_grid(), t, frame, version="2024Q1", sample=None)
    assert any("row count" in i for i in issues)


def test_verify_flags_value_mismatch():
    t = _table()
    frame = apply_table(sample_grid(), t, version="2024Q1").copy()
    frame.loc[0, "value"] = 99999.0  # corrupt one output value
    issues = verify_table(sample_grid(), t, frame, version="2024Q1", sample=None)
    assert issues


def _dup_grid_and_table():
    from moa.models import Cell, CellGrid
    grid = CellGrid(sheet="데이터", n_rows=5, n_cols=3, cells=[
        Cell(row=2, col=2, value="지역"), Cell(row=2, col=3, value="값"),
        Cell(row=3, col=2, value="서울"), Cell(row=3, col=3, value=10.0),
        Cell(row=4, col=2, value="서울"), Cell(row=4, col=3, value=20.0),  # same label, diff value
        Cell(row=5, col=2, value="합계"), Cell(row=5, col=3, value=30.0),
    ])
    table = TemplateSpec.model_validate({
        "template_id": "dup", "version": 1, "sheets": [{
            "sheet_match": {"by": "name", "value": "데이터"},
            "tables": [{"name": "t", "kind": "table",
                "region": {"start": "B2", "end": "C5"},
                "header": {"orientation": "top", "levels": 1, "rows": [2]},
                "index_columns": [{"col": "B", "name": "region", "type": "str"}],
                "value_block": {"cols": ["C", "C"]},
                "unpivot": {"var_name": "metric", "value_name": "value"},
                "column_semantics": [{"source": "C2", "name": "값", "type": "number"}],
                "version": {"source": {"from": "filename", "pattern": r"(\d{4})Q([1-4])"}, "name": "version"},
                "totals": [{"kind": "row_subtotal", "label": "합계", "over": "region"}]}]}]}).sheets[0].tables[0]
    return grid, table


def test_verify_handles_duplicate_index_labels():
    # Regression for the dict-overwrite collision bug: two rows labelled "서울"
    # with different values must both verify (Counter multiset), not collapse.
    grid, table = _dup_grid_and_table()
    frame = apply_table(grid, table, version="2024Q1")
    assert len(frame) == 2
    assert verify_table(grid, table, frame, version="2024Q1", sample=None) == []
    # corrupting one of the duplicate rows must still be caught
    frame.loc[frame.index[0], "value"] = 12345.0
    assert verify_table(grid, table, frame, version="2024Q1", sample=None)
