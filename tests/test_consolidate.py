import pandas as pd
from tests.fixtures import sample_grid, sample_pivot_raw
from moa.models import Cell
from moa.spec import TemplateSpec, sample_spec_dict
from moa.consolidate import detect_drift, consolidate


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


def test_consolidate_stacks_versions():
    grids = {"데이터": sample_grid()}
    res = consolidate(
        ["f_2024Q1.xlsx", "f_2024Q2.xlsx"], _spec(),
        list_sheets_fn=lambda p: [type("S", (), {"name": "데이터"})()],
        sheet_extractor=lambda p, s: grids[str(s)],
        pivot_extractor=lambda p, s, n: sample_pivot_raw(),
    )
    assert set(res.tables["by_industry"]["version"]) == {"2024Q1", "2024Q2"}
    assert len(res.tables["by_industry"]) == 8
    assert res.drift_by_file == {}


def test_consolidate_flags_version_collision():
    # 두 파일이 같은 버전(2024Q1)을 내면 합본에서 구분 불가 -> version_issues 로 신고
    grids = {"데이터": sample_grid()}
    res = consolidate(
        ["a_2024Q1.xlsx", "b_2024Q1.xlsx"], _spec(),
        list_sheets_fn=lambda p: [type("S", (), {"name": "데이터"})()],
        sheet_extractor=lambda p, s: grids[str(s)],
        pivot_extractor=lambda p, s, n: sample_pivot_raw(),
    )
    assert res.version_issues
    assert any("2024Q1" in i for i in res.version_issues)


def test_consolidate_flags_unresolved_version():
    # 파일명에 버전이 없으면 version=None -> 버전 구분 불가로 신고
    grids = {"데이터": sample_grid()}
    res = consolidate(
        ["plain.xlsx"], _spec(),
        list_sheets_fn=lambda p: [type("S", (), {"name": "데이터"})()],
        sheet_extractor=lambda p, s: grids[str(s)],
        pivot_extractor=lambda p, s, n: sample_pivot_raw(),
    )
    assert any("None" in i or "resolve" in i for i in res.version_issues)


def test_consolidate_no_version_issue_when_distinct():
    grids = {"데이터": sample_grid()}
    res = consolidate(
        ["f_2024Q1.xlsx", "f_2024Q2.xlsx"], _spec(),
        list_sheets_fn=lambda p: [type("S", (), {"name": "데이터"})()],
        sheet_extractor=lambda p, s: grids[str(s)],
        pivot_extractor=lambda p, s, n: sample_pivot_raw(),
    )
    assert res.version_issues == []


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
    assert set(res.tables["by_industry"]["version"]) == {"2024Q1"}
    assert "f_2024Q2.xlsx" in res.drift_by_file


def test_consolidate_progress_events_ok():
    sp = _spec()
    sheet_name = str(sp.sheets[0].sheet_match.value)
    events = []
    res = consolidate(
        ["f_2024Q1.xlsx", "f_2024Q2.xlsx"], sp,
        list_sheets_fn=lambda p: [type("S", (), {"name": sheet_name})()],
        sheet_extractor=lambda p, s: sample_grid(),
        pivot_extractor=lambda p, s, n: sample_pivot_raw(),
        progress=events.append,
    )
    assert list(res.tables)
    assert [e["event"] for e in events] == ["file_start", "file_done", "file_start", "file_done"]
    assert [(e["index"], e["total"]) for e in events] == [(1, 2), (1, 2), (2, 2), (2, 2)]
    done = [e for e in events if e["event"] == "file_done"]
    assert [e["status"] for e in done] == ["ok", "ok"]
    assert all(e["rows"] > 0 for e in done)
    assert all(e["versions"] for e in done)


def test_consolidate_progress_events_drift_skip():
    sp = _spec()
    sheet_name = str(sp.sheets[0].sheet_match.value)
    bad = sample_grid()
    bad.cells = [c for c in bad.cells if not (c.row == 5 and c.col == 3)]
    bad.cells.append(Cell(row=5, col=3, value="?됰슧"))
    events = []
    res = consolidate(
        ["f_2024Q1.xlsx"], sp,
        list_sheets_fn=lambda p: [type("S", (), {"name": sheet_name})()],
        sheet_extractor=lambda p, s: bad,
        pivot_extractor=lambda p, s, n: sample_pivot_raw(),
        on_drift="stop",
        progress=events.append,
    )
    assert res.tables == {}
    assert [e["event"] for e in events] == ["file_start", "file_done"]
    done = events[-1]
    assert done["status"] == "drift_skip"
    assert done["rows"] == 0
    assert done["issues"]
