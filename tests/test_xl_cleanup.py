from xltidy._xl import quit_app


class _FakeApp:
    """quit()/kill()을 흉내내는 더미(실제 Excel 없이 정리 로직만 검증)."""

    def __init__(self, quit_raises: bool = False, kill_raises: bool = False):
        self.quit_called = False
        self.kill_called = False
        self._qr, self._kr = quit_raises, kill_raises

    def quit(self):
        self.quit_called = True
        if self._qr:
            raise RuntimeError("quit failed")

    def kill(self):
        self.kill_called = True
        if self._kr:
            raise RuntimeError("kill failed")


def test_quit_then_kill_backstop():
    a = _FakeApp()
    quit_app(a)
    assert a.quit_called and a.kill_called  # kill은 항상 시도되는 백스톱


def test_kills_even_when_quit_fails():
    a = _FakeApp(quit_raises=True)
    quit_app(a)
    assert a.kill_called  # quit이 실패해도 kill로 프로세스 종료 보장


def test_swallows_all_errors():
    quit_app(_FakeApp(quit_raises=True, kill_raises=True))  # 정리가 예외를 던지면 안 됨
