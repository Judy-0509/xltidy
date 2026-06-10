import pytest
xw = pytest.importorskip("xlwings")


@pytest.mark.excel
def test_workbook_table_and_pivot(tmp_path):
    from moa.extract import extract, list_sheets
    from moa.pivot import extract_pivot
    from moa.apply import apply_workbook
    from moa.spec import TemplateSpec

    path = str(tmp_path / "survey_2024Q1.xlsx")
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.add()
        d = wb.sheets[0]; d.name = "데이터"
        d.range("B2").value = [["지역", "인구"], ["서울", 10], [None, 20], ["합계", 30]]
        d.range("B3:B4").api.Merge()  # 서울 세로 병합
        # 피벗
        src = wb.sheets.add("src", after=d)
        # data-field header must be ASCII ("val"); this Excel rejects
        # Orientation=xlDataField on a non-ASCII field name. Value/label
        # semantics unchanged.
        src.range("A1").value = [["지역", "구분", "val"], ["서울", "인구", 10], ["부산", "인구", 20]]
        pv = wb.sheets.add("피벗", after=src)
        cache = wb.api.PivotCaches().Create(SourceType=1, SourceData="src!A1:C3")
        pt = cache.CreatePivotTable(TableDestination=pv.range("A3").api, TableName="PT1")
        pt.PivotFields("지역").Orientation = 1
        f = pt.PivotFields("val"); f.Orientation = 4; f.Function = -4157
        wb.save(path); wb.close()
    finally:
        app.quit()

    spec = TemplateSpec.model_validate({
        "template_id": "e2e", "version": 1, "sheets": [
            {"sheet_match": {"by": "name", "value": "데이터"}, "tables": [{
                "name": "region_pop", "kind": "table",
                "region": {"start": "B2", "end": "C5"},
                "header": {"orientation": "top", "levels": 1, "rows": [2]},
                "index_columns": [{"col": "B", "name": "region", "type": "str"}],
                "value_block": {"cols": ["C", "C"]},
                "unpivot": {"var_name": "metric", "value_name": "value"},
                "column_semantics": [{"source": "C2", "name": "인구", "type": "number"}],
                "period": {"source": {"from": "filename", "pattern": r"(\d{4})Q([1-4])"}, "name": "period"},
                "totals": [{"kind": "row_subtotal", "label": "합계", "over": "region"}]}]},
            {"sheet_match": {"by": "name", "value": "피벗"}, "tables": [{
                "name": "pivot_pop", "kind": "pivot", "pivot_name": None,
                "period": {"source": {"from": "filename", "pattern": r"(\d{4})Q([1-4])"}, "name": "period"}}]},
        ]})

    res = apply_workbook(path, spec, sheet_extractor=extract, pivot_extractor=extract_pivot,
                         list_sheets_fn=list_sheets, filename=path)
    assert set(res.tables) == {"region_pop", "pivot_pop"}
    rp = res.tables["region_pop"]
    assert len(rp) == 2 and set(rp["region"]) == {"서울"} and set(rp["period"]) == {"2024-1"}
    assert res.reconcile.ok is True
    assert float(res.tables["pivot_pop"]["value"].sum()) == 30.0
