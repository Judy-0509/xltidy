from tests.fixtures import sample_grid, sample_pivot_raw
from xltidy.models import Cell, CellGrid, MergedRange
from xltidy.spec import TemplateSpec, sample_spec_dict
from xltidy.apply import apply_table, finalize_pivot, resolve_period, apply_workbook


def _table():
    return TemplateSpec.model_validate(sample_spec_dict()).sheets[0].tables[0]


def test_resolve_period_from_filename():
    assert resolve_period(_table().period, "report_2024Q1.xlsx", sample_grid()) == "2024-1"


def test_apply_table_long_excludes_subtotal():
    df = apply_table(sample_grid(), _table(), period="2024Q1")
    assert len(df) == 4
    assert set(df.columns) == {"industry", "month", "value", "period"}
    assert "합계" not in set(df["industry"])
    row = df[(df.industry == "제조업") & (df.month == "2024-01")].iloc[0]
    assert row["value"] == 100.0 and row["period"] == "2024Q1"


def test_apply_table_fills_merged_index():
    grid = CellGrid(sheet="데이터", n_rows=5, n_cols=3, cells=[
        Cell(row=2, col=2, value="지역"), Cell(row=2, col=3, value="인구"),
        Cell(row=3, col=2, value="서울"), Cell(row=3, col=3, value=10.0),
        Cell(row=4, col=3, value=20.0),
        Cell(row=5, col=2, value="합계"), Cell(row=5, col=3, value=30.0),
    ], merged=[MergedRange(r1=3, c1=2, r2=4, c2=2)])
    table = TemplateSpec.model_validate({
        "template_id": "m", "version": 1, "sheets": [{
            "sheet_match": {"by": "name", "value": "데이터"},
            "tables": [{"name": "t", "kind": "table",
                "region": {"start": "B2", "end": "C5"},
                "header": {"orientation": "top", "levels": 1, "rows": [2]},
                "index_columns": [{"col": "B", "name": "region", "type": "str"}],
                "value_block": {"cols": ["C", "C"]},
                "unpivot": {"var_name": "metric", "value_name": "value"},
                "column_semantics": [{"source": "C2", "name": "인구", "type": "number"}],
                "period": {"source": {"from": "filename", "pattern": r"(\d{4})Q([1-4])"}, "name": "period"},
                "totals": [{"kind": "row_subtotal", "label": "합계", "over": "region"}]}]}]}).sheets[0].tables[0]
    df = apply_table(grid, table, period="2024Q1")
    assert len(df) == 2 and set(df["region"]) == {"서울"}


def test_finalize_pivot_attaches_period():
    raw, _ = sample_pivot_raw()
    table = TemplateSpec.model_validate({
        "template_id": "p", "version": 1, "sheets": [{
            "sheet_match": {"by": "name", "value": "피벗"},
            "tables": [{"name": "pv", "kind": "pivot", "pivot_name": None,
                "period": {"source": {"from": "filename", "pattern": r"(\d{4})Q([1-4])"}, "name": "period"}}]}]}).sheets[0].tables[0]
    df = finalize_pivot(raw, table, period="2024Q1")
    assert "period" in df.columns and set(df["period"]) == {"2024Q1"}
    assert df["value"].sum() == 42.0


def test_apply_workbook_table_plus_pivot_injected():
    spec = TemplateSpec.model_validate({
        "template_id": "wb", "version": 1, "sheets": [
            sample_spec_dict()["sheets"][0],
            {"sheet_match": {"by": "name", "value": "피벗"},
             "tables": [{"name": "pv", "kind": "pivot", "pivot_name": None,
                         "period": {"source": {"from": "filename", "pattern": r"(\d{4})Q([1-4])"}, "name": "period"}}]},
        ]})
    grids = {"데이터": sample_grid()}
    res = apply_workbook(
        "f_2024Q1.xlsx", spec,
        sheet_extractor=lambda p, s: grids[str(s)],
        pivot_extractor=lambda p, s, name: sample_pivot_raw(),
        list_sheets_fn=lambda p: [],
    )
    assert set(res.tables) == {"by_industry", "pv"}
    assert len(res.tables["by_industry"]) == 4
    assert res.reconcile.ok is True  # 표 소계 + 피벗 총합계 모두 정합
