---
name: xltidy
description: Use when turning complex survey-institution Excel files into tidy DB-ready tables (one folder of CSV/Parquet per workbook), or when consolidating monthly/quarterly versions of the same template. Reads via xlwings only (no openpyxl). Lists all sheets (incl hidden) for the user to choose, infers a reusable TemplateSpec for messy tables, extracts pivot tables natively via COM, applies deterministically with reconciliation and drift detection.
license: MIT
compatibility: opencode
metadata:
  toolset: xltidy-cli
---

## 무엇을 하는가
xlwings로만 복잡한 Excel을 읽어 워크북 1개를 출력 폴더 1개(시트/표마다 tidy CSV/Parquet)로 정규화하고 월/분기 버전을 통합한다. **일반 표는 내가(LLM) 구조만 추론하고 값은 전사하지 않는다. 피벗 테이블은 COM 네이티브 추출이라 내가 추론하지 않는다.**

## 워크플로
0. **시트 선택(필수, 먼저)**: `xltidy sheets <대표파일.xlsx>` → 모든 시트(숨김·very_hidden 포함)를 사용자에게 보여주고, **어느 시트를 DB화할지** 먼저 물어 확정한다. (DB화 = 워크북 → 폴더, 시트/표마다 테이블)
1. **스펙 작성(양식당 1회)**:
   - 표 시트: `xltidy infer <파일> --sheet <시트> --backend agent` → 인코딩을 읽고 `specs/<id>.yaml`의 해당 sheets[] 항목을 작성. 다중헤더는 값열마다 column_semantics 한 항목으로 평탄화. 숫자는 #num이니 좌표·구조만.
   - 피벗 시트: 추론하지 말고 `kind: pivot` + `pivot_name`(없으면 null) + `period`만 적는다.
2. **검증**: `xltidy spec-validate specs/<id>.yaml --against <파일> --sheet <시트>`
3. **단일 적용(검증 포함)**: `xltidy apply specs/<id>.yaml --file <파일> --out-dir out/<name> --format csv --verify --sample 100` → reconcile/verify ✗면 스펙(소계/영역/피벗명) 수정.
4. **★ 다른 버전 물어보기**: 사용자에게 *"이 엑셀의 다른 기간(월/분기) 버전 파일이 더 있나요?"* 를 묻는다.
   - **있으면**: 파일 경로/글롭을 받아 `xltidy consolidate specs/<id>.yaml "data/2024*.xlsx" --out-dir merged --format csv --on-drift stop --verify --sample 100` → 모든 버전을 `period` 축으로 누적한 **하나의 DB**. 드리프트 파일은 제외(양식 변경이면 version↑).
   - **없으면**: 3단계 폴더가 최종 DB. 끝.

## 주의
- openpyxl/pd.read_excel 절대 제안 금지(정책 하드 금지).
- 무인 자동화면 `--backend qwen`(env `XLTIDY_QWEN_*`).
