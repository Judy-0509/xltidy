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


def _hide_row_item(app, path, field, item):
    """저장된 피벗에서 한 행 항목을 숨겨(필터) 둔다."""
    wb = app.books.open(path)
    try:
        pt = wb.sheets["피벗"].api.PivotTables(1)
        pt.PivotFields(field).PivotItems(item).Visible = False
        wb.save()
    finally:
        wb.close()


@pytest.mark.excel
def test_extract_pivot_long_and_grand_total(tmp_path):
    from moa.pivot import extract_pivot
    path = str(tmp_path / "p.xlsx")
    app = xw.App(visible=False, add_book=False)
    try:
        _build_pivot(app, path)
    finally:
        app.quit()
    frame, gt = extract_pivot(path, "피벗", None)
    assert float(frame["value"].sum()) == 42.0  # 소계/총합계 셀은 제외됨
    assert gt == 42.0
    assert set(frame["지역"]) == {"서울", "부산"}  # 행필드 이름 원형 유지(rename 없음)


@pytest.mark.excel
def test_extract_pivot_multi_data_field_keeps_first_only(tmp_path):
    # 데이터 필드 2개: "Σ값" 의사필드가 축 카운트를 부풀려도 v1 셀만 정확히
    # 추출돼야 한다(보정 없으면 전부 소계로 오인되어 빈 결과, 필터 없으면 v2 혼입).
    from moa.pivot import extract_pivot
    path = str(tmp_path / "p2.xlsx")
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.add()
        src = wb.sheets[0]; src.name = "src"
        src.range("A1").value = [["지역", "구분", "v1", "v2"],
                                 ["서울", "인구", 10, 1], ["서울", "가구", 4, 2],
                                 ["부산", "인구", 20, 3], ["부산", "가구", 8, 4]]
        pv = wb.sheets.add("피벗", after=src)
        cache = wb.api.PivotCaches().Create(SourceType=1, SourceData="src!A1:D5")
        pt = cache.CreatePivotTable(TableDestination=pv.range("A3").api, TableName="PT1")
        pt.PivotFields("지역").Orientation = 1
        d1 = pt.PivotFields("v1"); d1.Orientation = 4; d1.Function = -4157
        d2 = pt.PivotFields("v2"); d2.Orientation = 4; d2.Function = -4157
        wb.save(path); wb.close()
    finally:
        app.quit()
    frame, gt = extract_pivot(path, "피벗", None)
    assert float(frame["value"].sum()) == 42.0   # v1만 (v2 혼입 시 52)
    assert len(frame) == 2                        # 지역 2행, 빈 결과 회귀 방지
    assert set(frame["지역"]) == {"서울", "부산"}
    assert gt == 42.0                             # v1의 총합계


@pytest.mark.excel
def test_apply_workbook_real_single_open(tmp_path):
    # 실제 open-once 경로(open_excel_session -> ExcelSession.grid -> apply_session) 검증
    from moa.apply import apply_workbook
    from moa.spec import TemplateSpec, sample_spec_dict
    path = str(tmp_path / "tbl.xlsx")
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.add()
        sht = wb.sheets[0]; sht.name = "데이터"
        sht.range("B5").value = [["산업", "2024년 1월", "2024년 2월"],
                                 ["제조업", 100, 110],
                                 ["서비스업", 200, 210],
                                 ["합계", 300, 320]]
        wb.save(path); wb.close()
    finally:
        app.quit()
    spec = TemplateSpec.model_validate(sample_spec_dict())
    res = apply_workbook(path, spec, version="2024Q1")  # 주입 없이 = 파일 1회 열기
    assert set(res.tables) == {"by_industry"}
    assert len(res.tables["by_industry"]) == 4   # 산업 2 x 월 2
    assert res.reconcile.ok                        # 합계 소계 == 성분 합


@pytest.mark.excel
def test_extract_pivot_clears_filters_to_get_all_data(tmp_path):
    # 핵심 검증: 피벗에 필터(서울 숨김)가 걸려 있어도 기본(clear_filters=True)이면
    # 전체 데이터를 돌려주고, clear_filters=False면 보이는(필터된) 부분만 준다.
    from moa.pivot import extract_pivot
    path = str(tmp_path / "pf.xlsx")
    app = xw.App(visible=False, add_book=False)
    try:
        _build_pivot(app, path)
        _hide_row_item(app, path, "지역", "서울")  # 서울 행 숨김 -> 부산만 보임
    finally:
        app.quit()

    # 화면 그대로(필터 유지): 부산만 (20+8 = 28)
    f_filt, _ = extract_pivot(path, "피벗", None, clear_filters=False)
    assert set(f_filt["지역"]) == {"부산"}
    assert float(f_filt["value"].sum()) == 28.0

    # 기본값(필터 해제): 전체 복원 (42), 두 지역 모두
    f_all, gt_all = extract_pivot(path, "피벗", None)
    assert set(f_all["지역"]) == {"서울", "부산"}
    assert float(f_all["value"].sum()) == 42.0
    assert gt_all == 42.0
