import pytest
xw = pytest.importorskip("xlwings")


@pytest.mark.excel
def test_list_sheets_includes_hidden(tmp_path):
    from xltidy.extract import list_sheets
    path = tmp_path / "t.xlsx"
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.add()
        wb.sheets[0].name = "데이터"
        h = wb.sheets.add("숨김", after=wb.sheets[0])
        h.range("A1").value = "x"
        h.api.Visible = 0  # xlSheetHidden
        wb.save(str(path)); wb.close()
    finally:
        app.quit()
    infos = list_sheets(str(path))
    names = {i.name: i for i in infos}
    assert "데이터" in names and "숨김" in names
    assert names["숨김"].visibility in ("hidden", "very_hidden")


@pytest.mark.excel
def test_extract_values(tmp_path):
    from xltidy.extract import extract
    path = tmp_path / "t.xlsx"
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.add()
        sht = wb.sheets[0]; sht.name = "데이터"
        sht.range("B5").value = "산업"; sht.range("C5").value = "2024년 1월"
        sht.range("B6").value = "제조업"; sht.range("C6").value = 100
        wb.save(str(path)); wb.close()
    finally:
        app.quit()
    grid = extract(str(path), sheet="데이터")
    assert grid.sheet == "데이터" and grid.at(5, 2) == "산업" and grid.at(6, 3) == 100
