import json
from tests.fixtures import sample_grid
from moa.encode import encode
from moa.infer import build_agent_prompt, infer_with_qwen


def test_agent_prompt():
    p = build_agent_prompt(encode(sample_grid()), filename="report_2024Q1.xlsx")
    assert "SHEET: 데이터" in p and "TemplateSpec" in p
    assert "값(숫자)을 전사하지" in p
    assert "피벗" in p  # 피벗은 직접 kind:pivot로 적으라는 안내


def test_infer_with_qwen_injected_client():
    payload = {"template_id": "x", "version": 1, "sheets": []}
    class C:
        class chat:
            class completions:
                @staticmethod
                def create(**kw):
                    return type("R", (), {"choices": [type("Ch", (), {
                        "message": type("M", (), {"content": json.dumps(payload)})()})()]})()
    spec = infer_with_qwen(encode(sample_grid()), filename="f.xlsx", model="qwen", client=C())
    assert spec.template_id == "x"
