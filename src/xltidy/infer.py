from __future__ import annotations

import json
from textwrap import dedent

from .spec import TemplateSpec

_INSTRUCTION = dedent(
    """
    너는 Excel 시트 인코딩을 보고 xltidy `TemplateSpec`(JSON)을 작성한다.
    - 표는 region/header/index_columns/value_block/column_semantics(source/name/source_text)/
      unpivot/period/totals 를 채운다. 다중헤더는 값열마다 column_semantics 한 항목으로 평탄화한다.
    - **실제 값(숫자)을 전사하지 마라.** 좌표·구조만(값은 #num 으로 가려져 있다).
    - **피벗 테이블 시트는 추론하지 말고** kind:pivot + pivot_name(없으면 null) + period 만 적는다.
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
                    base_url: str | None = None, api_key: str | None = None) -> TemplateSpec:
    if client is None:
        from openai import OpenAI
        client = OpenAI(base_url=base_url, api_key=api_key or "not-needed")
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": build_agent_prompt(encoded, filename=filename)}],
        response_format={"type": "json_object"}, temperature=0)
    return TemplateSpec.model_validate(json.loads(resp.choices[0].message.content))
