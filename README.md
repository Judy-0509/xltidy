# xltidy

xlwings-only로 복잡한 조사기관 Excel을 읽어 **워크북 1개 = 출력 폴더 1개**(시트/표마다 tidy CSV/Parquet)로 정규화하고, 월/분기 버전을 통합하는 도구. 재사용 가능한 `TemplateSpec`을 결정론적으로 적용한다. 일반 표는 LLM이 구조만 추론하고, 피벗 테이블은 COM 네이티브로 추출한다.

## 설치

```bash
python -m pip install -e ".[dev]"
# parquet 출력까지: python -m pip install -e ".[dev,parquet]"
# 사내 Qwen 자동 추론까지: python -m pip install -e ".[dev,parquet,qwen]"
```

요구사항: Python 3.10+, 그리고 COM 기능(`sheets`/`extract`/`apply`/`consolidate`의 피벗·시트 추출)에는 **Windows + 설치된 Microsoft Excel**이 필요하다.

## 빠른 시작

0. **시트 선택(먼저)** — 숨김·very_hidden 포함 모든 시트 확인:
   ```bash
   xltidy sheets <대표파일.xlsx>
   ```
1. **스펙 작성(양식당 1회)**:
   - 표 시트: `xltidy infer <파일> --sheet <시트> --backend agent` 로 인코딩을 받아 `specs/<id>.yaml`을 작성한다. 다중헤더는 값열마다 `column_semantics` 한 항목으로 평탄화. 숫자값은 `#num`으로 가려지므로 좌표·구조만.
   - 피벗 시트: 추론하지 말고 `kind: pivot` + `pivot_name`(없으면 null) + `period`만 적는다.
2. **검증**:
   ```bash
   xltidy spec-validate specs/<id>.yaml --against <파일> --sheet <시트>
   ```
3. **단일 적용**:
   ```bash
   xltidy apply specs/<id>.yaml --file <파일> --out-dir out/<period> --format csv
   ```
   reconcile(표 소계 == 합, 피벗 합 == 총합계)가 ✗면 스펙을 수정한다.
4. **버전 통합**:
   ```bash
   xltidy consolidate specs/<id>.yaml "data/2024*.xlsx" --out-dir merged --format parquet --on-drift stop
   ```
   드리프트(선택 시트 누락/헤더 변경 포함)가 보고된 파일은 적재되지 않는다. 양식 변경이면 스펙 `version`을 올린 뒤 재작성한다.

## 제약

- **xlwings 전용.** `openpyxl`, `pandas.read_excel`/`ExcelFile`은 하드 금지(`tests/test_no_openpyxl.py` 가드).
- 무인 자동 추론은 사내 Qwen 백엔드(`--backend qwen`) 옵션. 환경변수: `XLTIDY_QWEN_BASE_URL`, `XLTIDY_QWEN_API_KEY`, `XLTIDY_QWEN_MODEL`.
- 피벗 추출은 단일 데이터 필드(v1). 다중 데이터 필드는 WARNING 후 첫 필드만 사용.

## 테스트

```bash
python -m pytest -m "not excel"   # 코어(순수부) — Excel 불필요
python -m pytest -m excel         # COM(extract/pivot/e2e) — 데스크톱 Excel 필요
```

## Skill 설치 (opencode / Claude)

`skills/excel-to-db/`를 에이전트 skills 경로로 복사한다:
- opencode: `~/.config/opencode/skills/xltidy/`
- Claude: `.claude/skills/`

## 참고

- 설계 스펙: `docs/superpowers/specs/2026-06-10-xltidy-excel-to-db-design.md`
- 구현 계획: `docs/superpowers/plans/2026-06-10-xltidy.md`
