from __future__ import annotations

import contextlib

from ._xl import open_book
from .extract import grid_from_sheet, sheet_infos
from .models import CellGrid, SheetInfo
from .pivot import pivot_from_sheet


class ExcelSession:
    """열린 워크북 1개를 감싸 모든 읽기가 이를 재사용하게 한다(파일당 1회 open).

    시트별 grid는 캐시한다 — 같은 시트에 표가 여러 개여도 1회만 읽는다. COM 객체는
    절대 밖으로 내보내지 않고 SheetInfo/CellGrid/DataFrame/원시값만 반환한다
    (랭글링되지 않은 COM 참조가 EXCEL.EXE 를 좀비로 만든다).
    """

    def __init__(self, wb):
        self._wb = wb
        self._grids: dict[str, CellGrid] = {}

    def sheet_infos(self) -> list[SheetInfo]:
        return sheet_infos(self._wb)

    def grid(self, sheet) -> CellGrid:
        key = str(sheet)
        if key not in self._grids:
            self._grids[key] = grid_from_sheet(self._wb, sheet)
        return self._grids[key]

    def pivot(self, sheet, pivot_name=None, *, clear_filters: bool = True):
        return pivot_from_sheet(self._wb, sheet, pivot_name, clear_filters=clear_filters)


@contextlib.contextmanager
def open_excel_session(path: str, *, read_only: bool = True):
    """파일을 1회 열어 ExcelSession 을 제공하고, 끝나면 반드시 닫고 종료한다.

    피벗 필터 해제는 워크북을 쓰기 가능 상태로 요구하므로, 피벗이 있는 스펙은
    read_only=False 로 연다. 저장하지 않으므로 디스크 원본은 보존된다.
    """
    with open_book(path, read_only=read_only) as wb:
        yield ExcelSession(wb)


class FnSession:
    """주입된 path 기반 콜러블을 ExcelSession 처럼 보이게 하는 어댑터.

    테스트/`--grid` 경로에서 실제 Excel 을 열지 않고 같은 인터페이스를 제공한다.
    """

    def __init__(self, path, list_sheets_fn=None, sheet_extractor=None, pivot_extractor=None):
        self._p = path
        self._ls = list_sheets_fn
        self._se = sheet_extractor
        self._pe = pivot_extractor

    def sheet_infos(self):
        return self._ls(self._p) if self._ls else []

    def grid(self, sheet):
        return self._se(self._p, sheet)

    def pivot(self, sheet, pivot_name=None, *, clear_filters: bool = True):
        return self._pe(self._p, sheet, pivot_name)
