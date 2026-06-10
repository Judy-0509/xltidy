from __future__ import annotations

import xlwings as xw


def new_app():
    """헤드리스(보이지 않는) Excel 인스턴스를 만든다.

    display_alerts / screen_updating 을 꺼서 링크 업데이트·저장 확인 같은
    **모달 대화상자가 종료(quit)를 막는 일**을 방지한다. 모달이 뜨면 EXCEL.EXE
    가 응답 대기 상태로 남아 종료되지 않는데, 헤드리스에선 보이지도 않아 좀비가
    된다. 이 인스턴스는 사용자가 따로 연 Excel과 **별도 프로세스**다.
    """
    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        app.screen_updating = False
    except Exception:
        pass
    return app


def quit_app(app) -> None:
    """헤드리스 Excel 인스턴스를 **확실히** 종료한다.

    ``app.quit()`` 은 미해제 COM 참조가 있으면 EXCEL.EXE 를 남기는 일이 잦다
    (알려진 xlwings/COM 문제). 그래서 ``app.kill()`` 로 백스톱한다 — kill 은
    이 인스턴스의 프로세스만 강제 종료하므로 사용자가 따로 열어둔 워크북에는
    영향이 없다. 두 호출 모두 예외를 삼켜 정리 자체가 실패를 던지지 않게 한다.
    """
    try:
        app.quit()
    except Exception:
        pass
    try:
        app.kill()  # quit 이 프로세스를 남겼으면 강제 종료(이 인스턴스만)
    except Exception:
        pass
