import pandas as pd
from tests.fixtures import sample_grid, sample_pivot_raw
from xltidy.models import Cell
from xltidy.spec import TemplateSpec, sample_spec_dict
from xltidy.consolidate import detect_drift, consolidate


def _spec():
    return TemplateSpec.model_validate(sample_spec_dict())


def test_drift_none_when_matching():
    assert detect_drift(_spec(), available_sheets=["데이터"], grid_by_sheet={"데이터": sample_grid()}) == []


def test_drift_flags_missing_sheet():
    issues = detect_drift(_spec(), available_sheets=["다른시트"], grid_by_sheet={})
    assert any("데이터" in i for i in issues)


def test_drift_flags_renamed_header():
    g = sample_grid()
    g.cells = [c for c in g.cells if not (c.row == 5 and c.col == 3)]
    g.cells.append(Cell(row=5, col=3, value="엉뚱"))
    issues = detect_drift(_spec(), available_sheets=["데이터"], grid_by_sheet={"데이터": g})
    assert any("C5" in i for i in issues)


def test_consolidate_stacks_periods():
    grids = {"데이터": sample_grid()}
    res = consolidate(
        ["f_2024Q1.xlsx", "f_2024Q2.xlsx"], _spec(),
        list_sheets_fn=lambda p: [type("S", (), {"name": "데이터"})()],
        sheet_extractor=lambda p, s: grids[str(s)],
        pivot_extractor=lambda p, s, n: sample_pivot_raw(),
    )
    assert set(res.tables["by_industry"]["period"]) == {"2024-1", "2024-2"}
    assert len(res.tables["by_industry"]) == 8
    assert res.drift_by_file == {}


def test_consolidate_stop_excludes_drift_file():
    bad = sample_grid()
    bad.cells = [c for c in bad.cells if not (c.row == 5 and c.col == 3)]
    bad.cells.append(Cell(row=5, col=3, value="엉뚱"))
    grids = {"f_2024Q1.xlsx": sample_grid(), "f_2024Q2.xlsx": bad}
    res = consolidate(
        list(grids), _spec(),
        list_sheets_fn=lambda p: [type("S", (), {"name": "데이터"})()],
        sheet_extractor=lambda p, s: grids[p],
        pivot_extractor=lambda p, s, n: sample_pivot_raw(),
        on_drift="stop",
    )
    assert set(res.tables["by_industry"]["period"]) == {"2024-1"}
    assert "f_2024Q2.xlsx" in res.drift_by_file
