---
name: moa
description: Use when turning complex survey-institution Excel files into tidy DB-ready tables (one folder of CSV/Parquet per workbook), or when consolidating monthly/quarterly versions of the same template. Reads via xlwings only (no openpyxl). Lists all sheets (incl hidden) for the user to choose, infers a reusable TemplateSpec for messy tables, extracts pivot tables natively via COM, applies deterministically with reconciliation and drift detection.
license: MIT
compatibility: opencode
metadata:
  toolset: moa-cli
---

## 무엇을 하는가
Moa(모아)는 xlwings로만 복잡한 Excel을 읽어 워크북 1개를 출력 폴더 1개(시트/표마다 tidy CSV/Parquet)로 정규화하고 월/분기 버전을 통합한다. **일반 표는 내가(LLM) 구조만 추론하고 값은 전사하지 않는다. 피벗 테이블은 COM 네이티브 추출이라 내가 추론하지 않는다.** (CLI 명령은 `moa`; 구버전 `xltidy`도 동일하게 동작)

## 워크플로
0. **시트 선택(필수, 먼저)**: `moa sheets <대표파일.xlsx>` → 모든 시트(숨김·very_hidden 포함)를 사용자에게 보여주고, **어느 시트를 DB화할지** 먼저 물어 확정한다. (DB화 = 워크북 → 폴더, 시트/표마다 테이블)
1. **스펙 작성(양식당 1회)**:
   - **이름 규칙**: `<id>` = 엑셀 파일명에서 **기간/날짜 토큰을 뺀** kebab-case 슬러그(예: `Market Pulse - Flagship Model Sales, April 2026.xlsx` → `market-pulse-flagship-model-sales`). 기간을 빼야 월/분기 버전이 같은 스펙을 공유해 consolidate된다. 이 값을 `template_id`와 파일명(`specs/<id>.yaml`) 둘 다에 쓴다.
   - 표 시트: `moa infer <파일> --sheet <시트> --backend agent` → 인코딩을 읽고 `specs/<id>.yaml`의 해당 sheets[] 항목을 작성. 골격이 필요하면 `moa sample-spec`로 유효한 YAML 예시를 받아 채운다. 다중헤더는 값열마다 column_semantics 한 항목으로 평탄화. 숫자는 #num이니 좌표·구조만. **대형 시트는 자동으로 값-있는 행 기준 머리 40 + 꼬리 10행만 프롬프트에 싣는다(`--head/--tail`로 조절). 인코딩의 `DATA ROWS: a..b`가 표의 진짜 끝이니 region 끝 행은 b로 잡을 것.**
   - 피벗 시트: 추론하지 말고 `kind: pivot` + `pivot_name`(없으면 null) + `period`만 적는다. (추출 시 보고서 필터·숨김 항목·슬라이서를 **모두 해제하고 전체 데이터**를 가져온다.)
2. **검증**: `moa spec-validate specs/<id>.yaml --against <파일> --sheet <시트>`
3. **단일 적용(검증 기본 on)**: `moa apply specs/<id>.yaml --file <파일> --out-dir out/<엑셀-stem-슬러그> --format csv` (검증은 기본 수행; 끄려면 `--no-verify`. 출력 폴더는 워크북별 = 엑셀 파일명 기준, 1엑셀=1DB) → reconcile/verify FAIL이면 스펙(소계/영역/피벗명) 수정.
4. **★ 다른 버전 물어보기**: 사용자에게 *"이 엑셀의 다른 기간(월/분기) 버전 파일이 더 있나요?"* 를 묻는다.
   - **있으면**: 파일 경로/글롭을 받아 `moa consolidate specs/<id>.yaml "data/2024*.xlsx" --out-dir merged --format csv --on-drift stop` → 모든 버전을 `period` 축으로 누적한 **하나의 DB**. 드리프트 파일은 제외(양식 변경이면 version↑). **버전마다 period가 유일해야 한다** — 겹치거나 None이면 `period FAIL`로 신고되니(예: 파일명이 월 이름이면 연도만 잡혀 충돌) period 패턴/셀을 고친다.
   - **없으면**: 3단계 폴더가 최종 DB. 끝.

## 주의
- openpyxl/pd.read_excel 절대 제안 금지(정책 하드 금지).
- 무인 자동화면 `--backend qwen`(env `XLTIDY_QWEN_*`).
