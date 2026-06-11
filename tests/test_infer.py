import json
from tests.fixtures import sample_grid
from moa.encode import encode
from moa.infer import build_agent_prompt, infer_with_qwen


def test_agent_prompt():
    p = build_agent_prompt(encode(sample_grid()), filename="report_2024Q1.xlsx")
    assert "SHEET: 데이터" in p and "TemplateSpec" in p
    assert "값(숫자)을 전사하지" in p
    assert "피벗" in p  # 피벗은 직접 kind:pivot로 적으라는 안내


def _resp(content: str):
    return type("R", (), {"choices": [type("Ch", (), {
        "message": type("M", (), {"content": content})()})()]})()


def test_infer_with_qwen_injected_client():
    payload = {"template_id": "x", "version": 1, "sheets": []}
    class C:
        class chat:
            class completions:
                @staticmethod
                def create(**kw):
                    return _resp(json.dumps(payload))
    spec = infer_with_qwen(encode(sample_grid()), filename="f.xlsx", model="qwen", client=C())
    assert spec.template_id == "x"


def test_infer_with_qwen_retries_on_invalid_json():
    # 1번째 응답은 깨진 JSON → 에러를 되먹여 재요청, 2번째에 성공해야 한다.
    payload = {"template_id": "retry-ok", "version": 1, "sheets": []}
    calls: list[list] = []
    class C:
        class chat:
            class completions:
                @staticmethod
                def create(messages=None, **kw):
                    calls.append(messages)
                    return _resp("{broken" if len(calls) == 1 else json.dumps(payload))
    spec = infer_with_qwen("enc", filename="f.xlsx", model="qwen", client=C())
    assert spec.template_id == "retry-ok" and len(calls) == 2
    assert any("유효한 TemplateSpec JSON이 아니다" in m["content"] for m in calls[1])


def test_infer_with_qwen_raises_after_retries_exhausted():
    import pytest
    class C:
        class chat:
            class completions:
                @staticmethod
                def create(**kw):
                    return _resp("not json at all")
    with pytest.raises(ValueError, match="did not return a valid TemplateSpec"):
        infer_with_qwen("enc", filename="f.xlsx", model="qwen", client=C(), retries=1)
