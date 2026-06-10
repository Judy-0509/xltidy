from moa.config import load_settings


def test_env(monkeypatch):
    monkeypatch.setenv("MOA_QWEN_BASE_URL", "http://qwen.example.internal/v1")
    monkeypatch.setenv("MOA_QWEN_MODEL", "qwen2.5-72b")
    s = load_settings()
    assert s.base_url == "http://qwen.example.internal/v1" and s.model == "qwen2.5-72b"


def test_defaults(monkeypatch):
    for k in ("MOA_QWEN_BASE_URL", "MOA_QWEN_API_KEY", "MOA_QWEN_MODEL"):
        monkeypatch.delenv(k, raising=False)
    assert load_settings().base_url is None
