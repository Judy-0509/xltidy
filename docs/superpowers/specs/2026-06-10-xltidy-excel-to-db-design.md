# xltidy — 설계 문서 (Design Spec)

- **작성일**: 2026-06-10
- **상태**: 승인 대기 (Draft)
- **한 줄 요약**: 사내 환경에서 **xlwings로만** 복잡한 조사기관 Excel을 읽어, (opencode/Claude agent 또는 사내 Qwen이 만든) **재사용 가능한 TemplateSpec**을 결정론적으로 적용해 **tidy 시계열 테이블(CSV/Parquet)** 로 정규화하고, 동일 양식의 **월/분기 버전을 통합**하는 도구.

---

## 1. 배경 & 목표

조사기관에서 받는 Excel은 병합셀·다중헤더·소계행이 섞인 복잡한 보고서이며, **동일 양식이 월별/분기별로 반복**된다. 이를 LLM/RAG·분석에 쓸 수 있는 깔끔한 관계형(tidy) 데이터로 만들고, 버전 간 일관되게 통합하는 것이 목표다.

레퍼런스인 [exstruct](https://github.com/harumiWeb/exstruct)(Excel→구조화 JSON, openpyxl/COM 백엔드)의 데이터 모델·MCP 아이디어는 차용하되, **사내 제약**에 맞춰 다시 설계한다.

### 1.1 성공 기준 (Success Criteria)
1. openpyxl을 **직접·전이 의존 모두** 사용하지 않고 동작한다 (CI가 이를 강제).
2. 하나의 복잡한 양식에 대해 TemplateSpec을 한 번 만들면, 같은 양식의 N개 기간 파일을 추가 추론 없이 일관되게 tidy 변환한다.
3. 변환 결과가 원본과 정합한다 (행/열 개수, 소계==합계 검증 통과).
4. 새 기간 파일이 양식과 어긋나면(컬럼 추가/이름변경/영역 이동) **드리프트로 플래그**되어 조용한 데이터 오염이 없다.
5. opencode에서 `skill` 툴로 호출되어 agent가 전체 워크플로를 수행할 수 있다.

### 1.2 비목표 (Non-Goals, v1)
- SQL DB(SQLite/Postgres) 직접 적재 — v1은 CSV/Parquet/DataFrame까지. (어댑터는 후속)
- 자연어→SQL 질의, MCP 서버, 서버/헤드리스 배치 — 후속.
- C(휴리스틱 자동 표 감지) 보조 경로 — 후속. v1은 접근법 **A(템플릿-스펙)** 만.

---

## 2. 핵심 제약 & 불변식 (Invariants)

| # | 불변식 | 강제 방법 |
|---|--------|-----------|
| I1 | **읽기는 오직 xlwings(살아있는 Excel).** openpyxl 하드 금지 → `pd.read_excel`/`pd.ExcelFile`(내부 openpyxl), Docling/unstructured도 배제 | `tests/test_no_openpyxl.py`: `xltidy` import 후 `sys.modules`에 `openpyxl` 없음을 assert. 의존성 목록에서 제외 |
| I2 | **LLM은 구조·좌표·규칙(TemplateSpec)만 산출. 실제 숫자값은 절대 전사하지 않음.** 값은 결정론적 `apply`가 좌표에서 직접 읽음 | TemplateSpec 스키마에 데이터 값 필드 없음. `apply`는 LLM 비호출 |
| I3 | **정합성 검증은 1급 시민.** `apply`는 항상 reconciliation 리포트를 동반 | 행/열 카운트, 소계==합계(허용오차) 검사 내장. 실패 시 비0 종료/플래그 |
| I4 | 실행 환경은 **Windows 데스크톱 인터랙티브**(Excel 설치 가정). 헤드리스/동시성 비고려(v1) | 문서화. 코어 로직은 Excel 없이 테스트되도록 분리(§5) |

---

## 3. 아키텍처 개요

```
                    ┌─────────────── 추론 경계(LLM) ────────────────┐
Excel(.xlsx/.xlsm)  │                                               │   결정론(LLM 없음)
   │                │                                               │
   ├─ extract ─▶ CellGrid ─ encode ─▶ 압축텍스트 ─▶ [agent 또는 Qwen] ─▶ TemplateSpec.yaml ─┐
   │   (xlwings)    (pydantic)         (LLM 입력)    "구조만 출력"        (검증된 인터페이스)  │
   │                                                                                       ▼
   └────────────────── 각 월/분기 파일 + TemplateSpec ── apply ─▶ tidy DataFrame(+period) + reconcile 리포트
                                                          │
                          여러 tidy ── consolidate ─▶ 통합 시계열 롱테이블 + 드리프트 리포트 ─▶ CSV/Parquet
```

- **추론 경계 왼쪽(extract/encode)**: 결정론. **오른쪽 끝(apply/consolidate)**: 결정론. **가운데(infer)만 LLM**.
- LLM 단계는 두 백엔드 중 택1 (§4.3):
  - **agent 백엔드(기본)** — opencode/Claude agent가 자기 LLM(사내 Qwen 등)으로 spec 작성. 별도 API 호출 없음.
  - **qwen 백엔드(옵션)** — OpenAI 호환 사내 Qwen을 프로그램적으로 호출. `pip install xltidy[qwen]`.

---

## 4. 컴포넌트 명세

각 컴포넌트는 **무엇을 하는가 / 인터페이스 / 의존**을 갖는다.

### 4.1 `extract.py` — xlwings 리더
- **무엇**: 살아있는 Excel을 열어 지정 시트의 used range를 읽어 `CellGrid`로 변환. 값·수식·표시서식(number_format)·병합영역을 포함.
- **인터페이스**: `extract(path: str, sheet: str | int | None = None) -> CellGrid`
- **의존**: xlwings (Excel COM). **openpyxl/pd.read_excel 금지.**
- **주의**: 값은 수식 재계산된 표시값을 취득(조사기관 수식 다수 대응). used range만 읽어 토큰·성능 관리.

### 4.2 `encode.py` — LLM용 압축 인코딩
- **무엇**: `CellGrid`를 LLM이 구조를 파악하기 쉬운 **압축 텍스트**로 직렬화 (Microsoft *SheetCompressor* 아이디어 차용: 동일/빈 셀 영역 런렝스 압축, 주소 그리드 유지, 숫자값은 형식만/샘플만 노출해 토큰 절감 + I2 강화).
- **인터페이스**: `encode(grid: CellGrid, max_cells: int = ...) -> str`
- **의존**: 없음(순수 함수). → 단위 테스트 용이.

### 4.3 `infer.py` — TemplateSpec 추론 (2 백엔드)
- **무엇**: 인코딩된 시트 + TemplateSpec JSON 스키마를 주고 spec 초안을 산출.
- **인터페이스**:
  - `agent` 백엔드: CLI가 인코딩 + 스키마 스캐폴드를 stdout로 출력 → **agent가 `spec.yaml`을 직접 작성** → `spec validate`. (LLM 호출은 opencode/Claude 런타임이 수행)
  - `qwen` 백엔드: `infer_with_qwen(encoded: str, *, base_url, model) -> TemplateSpec`. openai SDK, structured output(pydantic 스키마)로 강제.
- **의존**: 기본 경로는 없음. qwen 경로만 `openai`(extras).

### 4.4 `spec.py` — TemplateSpec 모델
- **무엇**: pydantic v2 모델 + YAML 로드/저장 + 검증(`validate`).
- TemplateSpec 핵심 필드(개념):
  ```yaml
  template_id: kosis-employment-monthly      # 양식 식별자
  version: 1                                  # 스펙 버전(드리프트 시 증가)
  sheet_match: { by: name|index, value: "데이터" }
  tables:
    - name: employment_by_industry
      region: { start: "B5", end: "M40" }     # 또는 anchor 기반
      header: { orientation: top, levels: 2, rows: [5, 6] }  # 다중헤더
      index_columns:                           # 행 라벨(차원)
        - { col: "B", name: industry, type: str }
      value_block: { cols: ["C", "M"] }        # 숫자 매트릭스 범위
      unpivot: { var_name: metric, value_name: value }  # wide→long melt
      column_semantics:                        # LLM이 부여한 의미
        - { source: "C5", name: "manufacturing", type: number }
      period:                                  # 기간 차원 출처
        source: { from: filename, pattern: "(\\d{4})[._-]?Q?([1-4]|\\d{2})" }
        name: period
      totals:                                  # reconciliation 훅
        - { kind: row_subtotal, label: "합계", over: industry }
  ```
- **인터페이스**: `TemplateSpec.from_yaml(path)`, `.to_yaml(path)`, `.validate(grid: CellGrid | None) -> list[Issue]`
- **의존**: pydantic, pyyaml.

### 4.5 `apply.py` — 스펙 적용(결정론)
- **무엇**: `CellGrid` + `TemplateSpec` → tidy long `DataFrame`(+`period` 컬럼). 좌표에서 실제 값을 읽어 unpivot·타입 캐스팅. **항상 reconciliation 동반(I3)**.
- **인터페이스**: `apply(grid: CellGrid, spec: TemplateSpec, *, period: str|None=None) -> ApplyResult` (`ApplyResult.frame`, `.reconcile`, `.drift`)
- **의존**: pandas. **LLM 비호출.**

### 4.6 `reconcile.py` — 무결성 검사
- **무엇**: 추출 행/열 개수, 소계==구성합(허용오차), 비어있지 않은 값 비율 등 검사 → `ReconcileReport`.
- **인터페이스**: `reconcile(grid, spec, frame) -> ReconcileReport`
- **의존**: pandas, numpy.

### 4.7 `consolidate.py` — 버전 통합 + 드리프트
- **무엇**: 같은 스펙으로 변환된 N개 기간 결과를 `period` 축으로 스택해 통합 시계열 생성. 각 파일 적용 시 헤더 라벨·영역 형태가 스펙과 어긋나면 **드리프트**(추가/삭제/이름변경/영역 이동)를 diff로 리포트.
- **인터페이스**: `consolidate(files: list[str], spec: TemplateSpec) -> ConsolidateResult` (`.frame`, `.drift_by_file`)
- **의존**: extract→apply 재사용.

### 4.8 `cli.py` — Typer CLI
| 명령 | 동작 |
|------|------|
| `xltidy extract <file> [--sheet S] [--out grid.json]` | CellGrid JSON 출력 |
| `xltidy encode <file|grid.json> [--sheet S]` | LLM용 압축 인코딩 출력 |
| `xltidy infer <file> --backend {agent|qwen} [--out spec.yaml]` | spec 초안 |
| `xltidy spec validate <spec.yaml> [--against file]` | 스키마/영역 검증 |
| `xltidy apply <file> --spec spec.yaml [--period auto|VAL] [--out out.parquet]` | tidy 변환 + reconcile |
| `xltidy consolidate <dir|glob> --spec spec.yaml [--out merged.parquet] [--on-drift stop\|warn]` | 통합(완전 long) + 드리프트(기본 `stop`) |

### 4.9 `config.py`
- env: `XLTIDY_QWEN_BASE_URL`, `XLTIDY_QWEN_API_KEY`, `XLTIDY_QWEN_MODEL` (qwen 백엔드 전용). `.env` 선택 지원.

### 4.10 `skills/excel-to-db/SKILL.md` — opencode/Claude 호환 skill
- frontmatter: `name: xltidy`, `description: ...` (lowercase-hyphen 규칙 준수).
- 본문이 agent에게 워크플로를 지시:
  1. `xltidy extract` + `xltidy encode`로 템플릿 인코딩
  2. 인코딩을 읽고 문서화된 스키마대로 `spec.yaml` **작성**(= agent의 LLM 추론 단계)
  3. `xltidy spec validate`
  4. 한 파일에 `xltidy apply` → reconcile 확인
  5. 기간 파일들에 `xltidy consolidate` → 드리프트 리포트 처리
- 배치: `~/.config/opencode/skills/xltidy/SKILL.md`(전역) 또는 레포 `.opencode/skills/`. opencode가 `.claude/skills/`도 읽으므로 동일 파일 재사용.

---

## 5. 데이터 모델 & 테스트 전략

- **CellGrid가 핵심 디커플링 지점**: `encode/apply/reconcile/consolidate`는 모두 `CellGrid`(직렬화 가능 pydantic)에서 동작 → **Excel/xlwings 없이 JSON 픽스처로 단위 테스트 가능**. `extract`만 Excel 필요(Windows 통합 테스트).
- **테스트**:
  - 단위: encode 압축 정확성, apply unpivot/타입/period, reconcile 소계 검증, consolidate 스택·드리프트 — 모두 픽스처 CellGrid.
  - 정책: `test_no_openpyxl.py` (I1 강제).
  - 통합(Windows+Excel): extract 라운드트립.

---

## 6. 기술 스택 & 의존성

- Python 3.11+
- **코어**: `xlwings`, `pandas`, `pydantic>=2`, `pyyaml`, `typer`, `rich`
- **extras**: `[parquet]`→`pyarrow`, `[qwen]`→`openai`
- **dev**: `pytest`
- **금지**: `openpyxl`, `pandas.read_excel`/`ExcelFile`, Docling/unstructured(전이 openpyxl)

---

## 7. 리스크 & 완화

| 리스크 | 완화 |
|--------|------|
| xlwings는 Excel+Windows 필요, CI에서 extract 불가 | 코어 로직을 CellGrid로 분리(§5)해 Excel 없이 대부분 테스트. extract는 로컬 통합 테스트 |
| 다중헤더 추론 품질이 LLM에 의존 | `spec validate`(영역 경계/타입) + reconciliation(I3) + 드리프트 사람 검토로 3중 방어 |
| 거대 시트 COM 성능/토큰 | used range만 인코딩, 인코딩 시 값 샘플링(I2) |
| 사내 Qwen 미사용 시에도 동작해야 | agent 백엔드가 기본 경로 → openai는 옵션 의존 |

---

## 8. v1 범위 요약

- **포함**: extract / encode / infer(agent+qwen) / spec(+validate) / apply / reconcile / consolidate(+드리프트) / CLI / SKILL.md / CSV·Parquet / openpyxl-ban CI
- **후속**: C 휴리스틱 보조, SQL 어댑터, MCP 서버, 서버 배치, 자연어→SQL

---

## 9. 결정 사항 (Resolved Decisions)

1. **`period` 출처**: **파일명 패턴 우선, 매칭 실패 시 스펙 지정 셀**에서 읽음. (둘 다 미해결이면 reconcile 경고)
2. **통합 출력 형태**: **완전 long** — `(차원…, metric, value, period)` 롱포맷으로 강제. tidy/DB/시계열·RAG 최적.
3. **드리프트 시 기본 동작**: **플래그 후 중단(엄격)**. 어긋난 파일은 적재하지 않고 diff 리포트를 내보내 사람이 검토. `--on-drift {stop|warn}` 플래그로 변경 가능(기본 `stop`).
