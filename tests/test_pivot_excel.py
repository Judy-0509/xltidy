import pytest
xw = pytest.importorskip("xlwings")


def _build_pivot(app, path):
    import xlwings as xw
    wb = app.books.add()
    src = wb.sheets[0]; src.name = "src"
    src.range("A1").value = [["지역", "구분", "val"],
                             ["서울", "인구", 10], ["서울", "가구", 4],
                             ["부산", "인구", 20], ["부산", "가구", 8]]
    pv_sht = wb.sheets.add("피벗", after=src)
    cache = wb.api.PivotCaches().Create(SourceType=1, SourceData="src!A1:C5")
    pt = cache.CreatePivotTable(TableDestination=pv_sht.range("A3").api, TableName="PT1")
    pt.PivotFields("지역").Orientation = 1   # xlRowField
    pt.PivotFields("구분").Orientation = 1   # 2번째 행필드 → 소계 발생
    # NOTE: data-field source header must be ASCII ("val"); this Excel build
    # rejects Orientation=xlDataField on a non-ASCII field name (한글 "값" 실패).
    # The data-field name is in no assertion; row-field/value semantics unchanged.
    df = pt.PivotFields("val"); df.Orientation = 4; df.Function = -4157  # xlDataField, xlSum
    wb.save(path); wb.close()


@pytest.mark.excel
def test_extract_pivot_long_and_grand_total(tmp_path):
    from xltidy.pivot import extract_pivot
    path = str(tmp_path / "p.xlsx")
    app = xw.App(visible=False, add_book=False)
    try:
        _build_pivot(app, path)
    finally:
        app.quit()
    frame, gt = extract_pivot(path, "피벗", None)
    assert float(frame["value"].sum()) == 42.0  # 소계/총합계 셀은 제외됨
    assert gt == 42.0
    assert set(frame["region"]) == {"서울", "부산"}
