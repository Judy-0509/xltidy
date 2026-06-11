from __future__ import annotations

import json
from textwrap import dedent

from .spec import TemplateSpec

_INSTRUCTION = dedent(
    """
    너는 Excel 시트 인코딩을 보고 moa `TemplateSpec`(JSON)을 작성한다.
    - 표는 region/header/index_columns/value_block/column_semantics(source/name/source_text)/
      unpivot/version/totals 를 채운다. 다중헤더는 값열마다 column_semantics 한 항목으로 평탄화한다.
    - **실제 값(숫자)을 전사하지 마라.** 좌표·구조만(값은 #num 으로 가려져 있다).
    - 인코딩에 `DATA ROWS: a..b` 와 `... omitted ...` 가 있으면 표가 그만큼 길다는 뜻이다.
      region 끝 행은 **보이는 마지막 행이 아니라 b**(DATA ROWS의 끝)로 잡아라.
    - **피벗 테이블 시트는 추론하지 말고** kind:pivot + pivot_name(없으면 null) + version 만 적는다.
    - 출력은 유효한 JSON 하나.
    파일명: {filename}

    === 시트 인코딩 ===
    {encoded}
    """
).strip()


def build_agent_prompt(encoded: str, *, filename: str = "") -> str:
    schema = json.dumps(TemplateSpec.model_json_schema(), ensure_ascii=False)
    return f"{_INSTRUCTION.format(filename=filename, encoded=encoded)}\n\n=== TemplateSpec JSON Schema ===\n{schema}"


def infer_with_qwen(encoded: str, *, filename: str, model: str, client=None,
                    base_url: str | None = None, api_key: str | None = None,
                    max_tokens: int | None = None, guided_json: bool = False,
                    retries: int = 2) -> TemplateSpec:
    """사내 Qwen(OpenAI 호환)으로 TemplateSpec 추론.

    - 잘못된 JSON/스키마 불일치면 에러를 되먹여 `retries`회까지 재요청한다
      (json_object 모드는 스키마를 강제하지 못해 온프렘 모델에선 흔한 실패).
    - guided_json=True: vLLM 계열의 스키마 강제 디코딩(extra_body.guided_json).
      지원 서버에선 json_object 보다 훨씬 안정적이다.
    - max_tokens 미지정 시 서버 기본값 — 긴 스펙이 잘리면 명시할 것.
    """
    from pydantic import ValidationError
    if client is None:
        from openai import OpenAI
        client = OpenAI(base_url=base_url, api_key=api_key or "not-needed")
    kwargs: dict = {"model": model, "response_format": {"type": "json_object"}, "temperature": 0}
    if max_tokens:
        kwargs["max_tokens"] = max_tokens
    if guided_json:
        kwargs["extra_body"] = {"guided_json": TemplateSpec.model_json_schema()}
    messages = [{"role": "user", "content": build_agent_prompt(encoded, filename=filename)}]
    last_err: Exception | None = None
    for _ in range(max(retries, 0) + 1):
        resp = client.chat.completions.create(messages=messages, **kwargs)
        content = resp.choices[0].message.content
        try:
            return TemplateSpec.model_validate(json.loads(content))
        except (json.JSONDecodeError, ValidationError) as e:
            last_err = e
            messages = messages + [
                {"role": "assistant", "content": content},
                {"role": "user", "content": f"이전 출력이 유효한 TemplateSpec JSON이 아니다: {e}\n"
                                            f"수정한 유효 JSON 하나만 다시 출력하라."}]
    raise ValueError(f"qwen did not return a valid TemplateSpec after {max(retries, 0) + 1} "
                     f"attempt(s): {last_err}")
