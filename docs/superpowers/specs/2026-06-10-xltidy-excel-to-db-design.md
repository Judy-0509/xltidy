# xltidy — 설계 문서 (Design Spec)

- **작성일**: 2026-06-10 (rev2: 다중시트·피벗·폴더출력·시트선택 반영)
- **상태**: 승인됨, 구현 진행
- **한 줄 요약**: 사내 환경에서 **xlwings로만** 복잡한 조사기관 Excel을 읽어, 시작 시 **사용자가 고른 시트(숨김 포함)** 를 대상으로, (agent/Qwen이 만든) **재사용 TemplateSpec**을 결정론적으로 적용해 **워크북 1개 = 출력 폴더 1개(시트/표마다 tidy 테이블)** 로 정규화하고, 동일 양식의 **월/분기 버전을 통합**한다. 일반 표는 LLM이 구조를 추론하고, **피벗 테이블은 COM으로 네이티브 추출**(LLM 불필요)한다.

---

## 1. 배경 & 목표

조사기관 Excel은 병합셀·다중헤더·소계행이 섞인 복잡한 보고서이고, **한 워크북에 데이터 시트가 여러 개**이며(숨김 시트 포함), **피벗 테이블**도 자주 쓰인다. 또한 동일 양식이 **월/분기로 반복**된다. 이를 tidy 관계형 데이터로 만들어 **워크북 하나를 하나의 DB(=출력 폴더, 표마다 파일)** 로 정리하고 버전 간 통합하는 것이 목표다.

레퍼런스 [exstruct](https://github.com/harumiWeb/exstruct)(Excel→구조화 JSON; openpyxl/COM 백엔드)의 데이터 모델·MCP **개념에서 영감만** 받되(코드 복사 없음; exstruct는 BSD-3-Clause) **사내 제약**에 맞춰 처음부터 재설계한다.

### 1.1 성공 기준
1. openpyxl을 **직접·전이 모두** 쓰지 않는다 (CI가 강제).
2. 시작 시 **모든 시트(visible/hidden/very_hidden)를 나열**하고 사용자가 DB화할 시트를 고른다. 선택은 스펙에 기록되어 통합·배치에서 재질문 없이 재사용된다.
3. 양식당 한 번 TemplateSpec을 만들면 같은 양식 N개 기간 파일을 일관되게 변환한다.
4. **워크북 1개 → 출력 폴더 1개**(시트/표마다 tidy CSV/Parquet).
5. **피벗 테이블**은 COM PivotTables로 네이티브 tidy 추출한다.
6. 결과가 원본과 정합한다(표: 소계==구성합 / 피벗: 데이터셀 합==총합계).
7. 양식·시트가 어긋나면 드리프트로 플래그한다(컬럼/영역 + **선택 시트 누락/이름변경**).
8. opencode/Claude에서 `skill`로 호출되어 agent가 워크플로를 수행한다.

### 1.2 비목표 (v1)
- SQL DB 직접 적재(SQLite/Postgres) — 출력은 CSV/Parquet 폴더. (후속)
- 자연어→SQL, MCP 서버, 서버/헤드리스 배치 — 후속.
- C(휴리스틱 자동 표 감지) — 후속. v1은 접근법 A(템플릿-스펙)만.
- **다중 데이터 필드 피벗** 완전 지원 — v1은 단일 데이터 필드. 다중은 명시 로그/제한.

---

## 2. 핵심 제약 & 불변식

| # | 불변식 | 강제 |
|---|--------|------|
| I1 | **읽기는 오직 xlwings.** openpyxl 하드 금지 → `pd.read_excel`/`ExcelFile`, Docling/unstructured 배제 | `tests/test_no_openpyxl.py`(import 가드 + 소스 스캔) |
| I2 | **LLM은 구조·좌표만, 값은 절대 전사 금지.** 표 값은 결정론 `apply`가 좌표에서 직접 읽음. **피벗은 LLM 미사용**(COM이 구조를 앎) | TemplateSpec에 값 필드 없음; pivot 경로는 encode/infer 건너뜀 |
| I3 | **정합성 검증 1급.** 표=소계==구성합, 피벗=데이터셀 합==총합계 | `reconcile_table`/`reconcile_pivot` 내장 |
| I4 | 실행 환경 = **Windows 데스크톱 인터랙티브**(Excel 설치). 헤드리스/동시성 비고려 | 코어는 Excel 없이 테스트(§5); COM은 `@pytest.mark.excel` |
| I5 | **병합셀 라벨은 앵커값으로 해소.** 인덱스/라벨/헤더 읽기는 `value_filled` 사용 | `CellGrid.value_filled` + 단위/e2e 테스트 |

---

## 3. 아키텍처

```
0) xltidy sheets <file> ─▶ 모든 시트(숨김 포함) 나열 ─▶ [사용자가 DB화할 시트 선택] ─▶ 스펙 sheets[] 기록
                                                                   │
선택 시트별로:                                                      ▼
  kind=table:  extract(xlwings) ─▶ CellGrid ─ encode ─▶ [agent/Qwen] ─▶ TableSpec(구조)
  kind=pivot:  extract_pivot(COM PivotTables) ───────────────────────▶ TableSpec(피벗명+period만; LLM 건너뜀)
                                                                   │
각 월/분기 파일 + TemplateSpec ── apply_workbook(결정론) ─▶ {표이름: tidy long(+period)} + reconcile + drift
                                                                   │
여러 파일 ── consolidate ─▶ 표이름별 period 누적 + 드리프트(시트누락 포함) ─▶ 출력 폴더(<표>.csv/.parquet)
```

- **추론 경계**: 일반 표만 LLM(encode→infer). 피벗은 COM이 구조를 알기에 **LLM을 건너뛰어 더 안정적**.
- 추론 백엔드 2종: **agent(기본)** = opencode/Claude가 직접 작성 / **qwen(옵션)** = OpenAI 호환 사내 Qwen.

---

## 4. 컴포넌트 명세

### 4.1 `extract.py` — xlwings 리더 + 시트 탐색
- `list_sheets(path) -> list[SheetInfo]`: 모든 시트 나열. `SheetInfo{name, index, visibility: visible|hidden|very_hidden, used_rows, used_cols, n_pivots}`. 빈 시트의 used_range는 0으로 안전 처리.
- `extract(path, sheet) -> CellGrid`: 한 시트의 used range를 값·수식·병합과 함께 읽음. **openpyxl/pd.read_excel 금지.**
- visibility: `sht.api.Visible`(-1/0/2). n_pivots: `sht.api.PivotTables().Count`.

### 4.2 `pivot.py` — 피벗 네이티브 추출 (COM, 최고위험)
- `extract_pivot(path, sheet, pivot_name=None) -> pd.DataFrame`(tidy long): `DataBodyRange` 셀을 순회하며 `cell.PivotCell.RowItems`/`.ColumnItems`로 축 튜플을 얻고, `PivotCell.PivotCellType`로 **소계/총합계 셀을 건너뜀**(데이터셀=`xlPivotCellValue`). 행필드 컬럼들 + 열필드 var + value.
- v1: **단일 데이터 필드**. 다중 데이터 필드는 명시 로그 후 제한.
- `grand_total(path, sheet, pivot_name) -> float|None`: 총합계 셀 값(정합성용).

### 4.3 `encode.py` — LLM용 압축 인코딩
- `encode(grid) -> str`: 동일/빈 영역 압축, 주소 그리드 유지, **숫자값은 `#num` 마스킹**(I2). 순수 함수. **kind=table 에만 사용**(피벗 미사용).

### 4.4 `infer.py` — TemplateSpec 추론(2 백엔드, 표 전용)
- agent: CLI가 인코딩+JSON Schema를 출력 → agent가 `spec.yaml` 작성 → `spec-validate`.
- qwen: `infer_with_qwen(encoded, ...) -> TemplateSpec` (openai extras, structured output).
- **피벗 테이블은 추론 대상 아님** — 스펙에 직접 `kind: pivot` + `pivot_name`만 적음.

### 4.5 `spec.py` — TemplateSpec(다중시트)
```yaml
template_id: kosis-emp-monthly
version: 1
sheets:
  - sheet_match: { by: name, value: "데이터" }
    tables:
      - name: emp_by_industry        # 출력 파일 stem
        kind: table
        region: { start: "B5", end: "M40" }
        header: { orientation: top, levels: 2, rows: [5, 6] }   # levels=메타; agent가 평탄화
        index_columns: [ { col: "B", name: industry, type: str } ]
        value_block: { cols: ["C", "M"] }
        unpivot: { var_name: month, value_name: value }
        column_semantics:
          - { source: "C5", name: "2024-01", type: number, source_text: "1월" }
        period: { source: { from: filename, pattern: "(\\d{4})[._-]?Q?([1-4])" }, name: period }
        totals: [ { kind: row_subtotal, label: "합계", over: industry } ]
  - sheet_match: { by: name, value: "피벗요약" }
    tables:
      - name: pivot_summary
        kind: pivot
        pivot_name: null            # null=시트 첫 피벗
        period: { source: { from: filename, pattern: "(\\d{4})Q([1-4])" }, name: period }
```
- 모델: `TemplateSpec{template_id, version, sheets:[SheetSpec]}`, `SheetSpec{sheet_match, tables:[TableSpec]}`, `TableSpec{name, kind:"table"|"pivot", period, (표필드들 Optional), pivot_name}`.
- `from_yaml/to_yaml`, `validate_against(grids)`: kind=table 필드 필수성, region 경계 검사.

### 4.6 `apply.py` — 적용(결정론) — 순수부 + DI 오케스트레이션
- 순수: `apply_table(grid: CellGrid, table) -> DataFrame`(병합 라벨은 `value_filled`로 해소, 소계행 제외, unpivot), `finalize_pivot(raw: DataFrame, table) -> DataFrame`(period 부착/정리).
- `resolve_period(period, filename, grid) -> str|None`(파일명 우선→셀).
- 오케스트레이션(Excel 접촉, DI로 테스트): `apply_workbook(path, spec, *, sheet_extractor=extract, pivot_extractor=extract_pivot, list_sheets_fn=list_sheets, period=None) -> ApplyResult`.
- `ApplyResult{tables: dict[str, DataFrame], reconcile: ReconcileReport, drift: list[str]}`.

### 4.7 `reconcile.py` — 무결성 (I3)
- `reconcile_table(grid, table)`: 소계행 값 == 구성행 합(허용오차). 라벨은 `value_filled`.
- `reconcile_pivot(data_cells_sum, grand_total)`: 합 == 총합계.
- 집계 `reconcile(...) -> ReconcileReport{ok, issues}`.

### 4.8 `consolidate.py` — 버전 통합 + 드리프트
- `detect_drift(path, spec, *, list_sheets_fn, sheet_extractor)`: ① **선택 시트 누락/이름변경**(스펙 sheet_match가 available 시트에 없음) ② 표: column_semantics.source_text 불일치/영역 초과 ③ 피벗: 해당 피벗 부재.
- `consolidate(files, spec, *, extractor_bundle, on_drift="stop") -> ConsolidateResult{tables: dict[str, DataFrame], drift_by_file}`: 파일별 apply_workbook → 표이름별 누적. 드리프트 파일은 stop 시 제외.

### 4.9 `dbio.py` — 출력 ("1 엑셀 = 1 폴더")
- `write_tables(tables: dict[str, DataFrame], out_dir, fmt="csv") -> list[str]`: `<out_dir>/<name>.csv|.parquet`. parquet는 pyarrow extras. csv는 표준.

### 4.10 `config.py`
- env: `XLTIDY_QWEN_BASE_URL/API_KEY/MODEL` (qwen 백엔드 전용).

### 4.11 `cli.py` — Typer
| 명령 | 동작 |
|------|------|
| `xltidy sheets <file> [--json]` | **모든 시트(숨김 포함) 나열** |
| `xltidy extract <file> --sheet S [--out grid.json]` | 한 시트 CellGrid |
| `xltidy encode <file|grid.json> --sheet S` | 표 인코딩 |
| `xltidy infer <file> --sheet S --backend agent|qwen` | 표 스펙 초안(피벗은 직접 작성) |
| `xltidy spec-validate <spec> [--against file]` | 검증 |
| `xltidy apply <spec> --file F [--period V] --out-dir DIR [--format csv|parquet]` | {표→파일} + reconcile |
| `xltidy consolidate <spec> "glob" --out-dir DIR [--format ...] [--on-drift stop|warn]` | 통합 + 드리프트 |

### 4.12 `skills/excel-to-db/SKILL.md`
- frontmatter `name: xltidy`. 워크플로 **0단계 = 시트 선택**:
  0. `xltidy sheets <대표파일>` → 사용자에게 **숨김 포함 전 시트** 제시 → DB화할 시트 확인.
  1. 표 시트: `infer ... --backend agent`로 인코딩 받아 `spec.yaml` 작성(다중헤더는 column_semantics로 평탄화). 피벗 시트: `kind: pivot`만 적음.
  2. `spec-validate` → `apply`(폴더 출력, reconcile) → `consolidate`(드리프트 처리).

---

## 5. 데이터 모델 & 테스트 전략

- **순수/COM 분리**가 테스트 핵심: `apply_table`/`finalize_pivot`/`reconcile_*`/`encode`/`detect_drift`(주입형)는 픽스처로 **Excel 없이** 단위 테스트. `list_sheets`/`extract`/`extract_pivot`만 COM(`@pytest.mark.excel`).
- **COM 테스트 공백 방지**: ① 대표적 피벗 e2e(행필드 2개 + 소계 + 총합계) 작성 ② **데스크톱에서 `-m excel` 스위트 실행 체크포인트** 태스크로 "작성됐지만 미실행" 방지.

---

## 6. 의존성 (openpyxl-free)

- 코어: `xlwings`, `pandas`, `pydantic>=2`, `pyyaml`, `typer`, `rich`
- extras: `[parquet]→pyarrow`, `[qwen]→openai`
- dev: `pytest`
- 금지: openpyxl, `pd.read_excel`/`ExcelFile`, Docling/unstructured

---

## 7. 리스크 & 완화

| 리스크 | 완화 |
|--------|------|
| **피벗 COM 추출(최고위험)**: 다중필드/소계/총합계 | `PivotCell.RowItems/ColumnItems`+`PivotCellType`로 소계·총합계 스킵, 단일 데이터필드 한정, 대표 e2e + 총합계 reconcile |
| COM(list_sheets/extract/pivot) 단위테스트 불가 | 순수/COM 분리 + 데스크톱 `-m excel` 실행 체크포인트 |
| 다중시트 양식 변경 | 시트 단위 드리프트(누락/이름변경) + `--on-drift` |
| 거대 시트 성능/토큰 | used range만, 숫자 `#num` 마스킹 |

---

## 8. v1 범위

- **포함**: sheets(숨김 포함 나열·선택) / extract / encode / infer(agent+qwen, 표) / **pivot 네이티브** / spec(다중시트, kind) / apply_workbook(표+피벗) / reconcile(표+피벗) / consolidate(시트 드리프트 포함) / dbio(폴더 CSV·Parquet) / CLI / SKILL.md / openpyxl-ban CI / `-m excel` 체크포인트
- **후속**: C 휴리스틱, SQL 어댑터, MCP, 서버 배치, 자연어→SQL, 다중 데이터필드 피벗

---

## 9. 결정 사항 (Resolved)

1. **period 출처**: 파일명 패턴 우선 → 실패 시 스펙 지정 셀.
2. **출력 형태**: 완전 long. **워크북 1개 = 출력 폴더 1개**, 시트/표마다 `<name>.csv|.parquet`.
3. **드리프트**: 기본 `stop`(어긋난 파일 제외 + diff 리포트). `--on-drift {stop|warn}`. **선택 시트 누락/이름변경도 드리프트.**
4. **다중시트**: v1 포함. 시작 시 `xltidy sheets`로 **숨김 포함 전 시트** 제시 후 사용자가 DB화할 시트 선택 → 스펙에 기록.
5. **피벗**: v1에서 **COM 네이티브 추출**(LLM 미사용). 단일 데이터필드 한정.
