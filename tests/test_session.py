import moa.session as S


def test_fnsession_routes_calls_to_injected_fns():
    sess = S.FnSession(
        "f.xlsx",
        list_sheets_fn=lambda p: [("info", p)],
        sheet_extractor=lambda p, s: ("grid", p, s),
        pivot_extractor=lambda p, s, n: ("pivot", p, s, n),
    )
    assert sess.sheet_infos() == [("info", "f.xlsx")]
    assert sess.grid("S1") == ("grid", "f.xlsx", "S1")
    assert sess.pivot("P1", "name") == ("pivot", "f.xlsx", "P1", "name")


def test_fnsession_sheet_infos_empty_without_fn():
    sess = S.FnSession("f.xlsx", sheet_extractor=lambda p, s: None)
    assert sess.sheet_infos() == []


def test_excelsession_caches_grid_per_sheet(monkeypatch):
    calls = {"n": 0}

    def fake_grid(wb, sheet):
        calls["n"] += 1
        return f"grid:{sheet}"

    monkeypatch.setattr(S, "grid_from_sheet", fake_grid)
    sess = S.ExcelSession(wb=object())
    assert sess.grid("데이터") == "grid:데이터"
    assert sess.grid("데이터") == "grid:데이터"  # 같은 시트 → 캐시, 재읽기 없음
    assert calls["n"] == 1
    sess.grid("다른시트")
    assert calls["n"] == 2  # 다른 시트는 새로 읽음
