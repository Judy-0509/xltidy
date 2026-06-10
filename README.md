<h1 align="center">xltidy</h1>

<p align="center">
  <b>Turn messy survey-institution Excel into tidy, DB-ready tables — using <code>xlwings</code> only.</b><br>
  <b>복잡한 조사기관 Excel을 <code>xlwings</code>만으로 깔끔한 DB용 테이블로.</b>
</p>

<p align="center">
  <img alt="python" src="https://img.shields.io/badge/python-3.10%2B-blue">
  <img alt="excel io" src="https://img.shields.io/badge/Excel%20I%2FO-xlwings%20only-success">
  <img alt="no openpyxl" src="https://img.shields.io/badge/openpyxl-banned-critical">
  <img alt="platform" src="https://img.shields.io/badge/platform-Windows%20%2B%20Excel-lightgrey">
  <img alt="license" src="https://img.shields.io/badge/license-MIT-green">
</p>

---

## 🇬🇧 What is this?

**xltidy** converts complex, real-world Excel workbooks from research/survey institutions into **tidy, DB-ready tables** — without ever touching `openpyxl`.

Survey Excel files are hard: merged cells, multi-level headers, subtotal rows, pivot tables, and the *same template repeated every month/quarter*. xltidy reads them with a **live Excel via `xlwings`** (so formula values, formats, and pivots are real), then applies a **reusable `TemplateSpec`** deterministically:

- **One workbook → one output folder** (one tidy CSV/Parquet file per sheet/table) — *"1 Excel = 1 DB"*.
- **Regular tables**: an LLM (your in-house Qwen, or an [opencode](https://opencode.ai)/Claude agent) infers the *structure only* once per template; the actual numbers are read deterministically from the spreadsheet — **the LLM never transcribes values**.
- **Pivot tables**: extracted natively via Excel COM (`PivotTables`) — no LLM needed, and more reliable than scraping rendered cells.
- **Version consolidation**: stack monthly/quarterly files into one time series with a `period` dimension, and **flag drift** when a new file no longer matches the template (renamed columns, missing sheets, shifted regions).
- **Integrity first**: every run reconciles (table subtotal == sum of components; pivot data sum == grand total).

Designed for **on-prem / in-house** use (data never leaves the building), so it pairs with a self-hosted Qwen and ships as an **opencode/Claude skill**.

## 🇰🇷 이게 뭔가요?

**xltidy**는 조사기관에서 받는 **복잡한 Excel**(병합셀·다중헤더·소계행·피벗·월/분기 반복 양식)을 `openpyxl` 없이 **깔끔한 DB용 tidy 테이블**로 변환하는 도구입니다.

`xlwings`로 **살아있는 Excel을 직접** 읽어(수식 계산값·서식·피벗이 모두 실제값) 재사용 가능한 **`TemplateSpec`** 을 결정론적으로 적용합니다.

- **워크북 1개 → 출력 폴더 1개** (시트/표마다 tidy CSV·Parquet 파일) — *"엑셀 1개 = DB 1개"*.
- **일반 표**: LLM(사내 Qwen 또는 [opencode](https://opencode.ai)/Claude 에이전트)이 양식당 **한 번 구조만** 추론하고, 실제 숫자값은 좌표에서 **결정론적으로** 읽습니다 — **LLM은 값을 전사하지 않습니다**.
- **피벗 테이블**: Excel COM(`PivotTables`)으로 **네이티브 추출** — LLM 불필요, 렌더 셀 스크래핑보다 안정적.
- **버전 통합**: 월/분기 파일을 `period` 차원으로 한 시계열에 쌓고, 새 파일이 양식과 어긋나면(컬럼 이름변경·시트 누락·영역 이동) **드리프트로 플래그**.
- **무결성 우선**: 매 실행마다 정합성 검증(표 소계 == 구성합, 피벗 데이터합 == 총합계).

데이터가 **사내 밖으로 나가지 않는** 온프렘 환경을 전제로 설계되어, 자체 호스팅 Qwen과 함께 쓰고 **opencode/Claude 스킬**로 제공됩니다.

---

## How it works · 동작 원리

```
0) xltidy sheets <file>  ─▶  모든 시트(숨김 포함) 나열  ─▶  [사용자가 DB화할 시트 선택]
                                                               │
   선택 시트별로 · per selected sheet:                          ▼
   kind=table:  extract (xlwings) ─▶ CellGrid ─ encode ─▶ [agent/Qwen] ─▶ TableSpec   (구조만/structure only)
   kind=pivot:  extract_pivot (COM PivotTables) ───────────────────────▶ TableSpec   (LLM 건너뜀/LLM-free)
                                                               │
   각 월/분기 파일 + TemplateSpec ─ apply ─▶ {표이름: tidy long(+period)} + reconcile + drift
                                                               │
   여러 파일 ─ consolidate ─▶ 표별 period 누적 + 드리프트 ─▶ 출력 폴더 (<table>.csv/.parquet)
```

The LLM only ever produces the `TemplateSpec` (coordinates and structure). Deterministic code reads the real values. · LLM은 `TemplateSpec`(좌표·구조)만 만들고, 실제 값은 결정론 코드가 읽습니다.

> All commands below are **PowerShell** (Windows). · 아래 모든 명령은 **PowerShell**(Windows) 기준입니다.

## Install · 설치

```powershell
python -m pip install -e ".[dev]"
# + Parquet output:        python -m pip install -e ".[dev,parquet]"
# + in-house Qwen backend: python -m pip install -e ".[dev,parquet,qwen]"
```

Requirements · 요구사항: **Python 3.10+**, and for COM features (`sheets` / `extract` / pivot in `apply`·`consolidate`) a **Windows machine with Microsoft Excel installed**.

## In-house setup · 사내 설치 (PowerShell)

```powershell
# 1) Clone the private repo (you must be authenticated to GitHub) · private 레포 클론
git clone https://github.com/Judy-0509/xltidy.git
Set-Location xltidy

# 2) (recommended) isolated environment · 가상환경 권장
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 3) Install · 설치 (Parquet + in-house Qwen backend 포함)
python -m pip install -e ".[dev,parquet,qwen]"

# 4) Point at your in-house Qwen (OpenAI-compatible) · 사내 Qwen 연결
#    this session only · 이 세션만:
$env:XLTIDY_QWEN_BASE_URL = "http://qwen.your-company.internal/v1"
$env:XLTIDY_QWEN_API_KEY  = "your-internal-key"
$env:XLTIDY_QWEN_MODEL    = "qwen2.5-72b-instruct"
#    persist for your user (new shells) · 사용자 환경변수로 영구 저장:
[Environment]::SetEnvironmentVariable("XLTIDY_QWEN_BASE_URL", "http://qwen.your-company.internal/v1", "User")
[Environment]::SetEnvironmentVariable("XLTIDY_QWEN_MODEL", "qwen2.5-72b-instruct", "User")

# 5) Smoke-test the install · 설치 점검
python -m pytest -m "not excel" -q     # core, no Excel
python -m pytest -m excel -q           # COM (needs desktop Excel)
```

> Using it through an **opencode/Claude agent** instead? You don't need the `qwen` backend or the env vars — the agent's own LLM authors the spec. Just install (step 3 without `,qwen`) and register the skill (see below). · opencode/Claude 에이전트로 쓰면 `qwen` 백엔드·환경변수가 필요 없습니다(에이전트 LLM이 스펙 작성). 설치 후 스킬만 등록하세요.

## Quickstart · 빠른 시작

```powershell
# 0) Choose sheets — lists ALL sheets incl. hidden / very-hidden
#    시트 선택 — 숨김·very_hidden 포함 전체 시트 확인
xltidy sheets report_2024Q1.xlsx

# 1) Author a TemplateSpec once per template · 양식당 1회 스펙 작성
#    table sheets: get an encoding for the agent to fill the spec
#    표 시트: 에이전트가 spec.yaml을 채우도록 인코딩 출력
xltidy infer report_2024Q1.xlsx --sheet 데이터 --backend agent
#    pivot sheets: just write `kind: pivot` + pivot_name + period in the spec
#    피벗 시트: 스펙에 kind:pivot + pivot_name + period 만 작성

# 2) Validate · 검증
xltidy spec-validate specs/employment.yaml --against report_2024Q1.xlsx --sheet 데이터

# 3) Apply one workbook → one folder of tidy tables · 단일 적용(폴더 출력)
xltidy apply specs/employment.yaml --file report_2024Q1.xlsx --out-dir out/2024Q1 --format csv

# 4) Consolidate monthly/quarterly versions · 월/분기 버전 통합
xltidy consolidate specs/employment.yaml "data/2024*.xlsx" --out-dir merged --format parquet --on-drift stop
```

`reconcile` mismatches (subtotal ≠ sum, pivot data ≠ grand total) and `drift` (renamed headers, missing selected sheets) are reported; drifted files are **excluded** under `--on-drift stop`. · 정합성 불일치와 드리프트가 보고되며, `--on-drift stop`이면 어긋난 파일은 적재되지 않습니다.

## Use as an opencode / Claude skill · 스킬로 사용

Copy the skill into your agent's skills path · 에이전트 skills 경로로 복사 (PowerShell):

```powershell
# opencode (global) · opencode 전역
$dst = "$env:USERPROFILE\.config\opencode\skills\xltidy"
New-Item -ItemType Directory -Force -Path $dst | Out-Null
Copy-Item ".\skills\excel-to-db\SKILL.md" "$dst\SKILL.md" -Force

# Claude (per-project) · Claude 프로젝트별
New-Item -ItemType Directory -Force -Path ".\.claude\skills\xltidy" | Out-Null
Copy-Item ".\skills\excel-to-db\SKILL.md" ".\.claude\skills\xltidy\SKILL.md" -Force
```

Then in opencode the agent calls `skill({ name: "xltidy" })`. · 이후 opencode에서 에이전트가 `skill({ name: "xltidy" })`로 호출.

The skill drives the whole workflow (sheet selection → spec authoring → apply → consolidate). Since the agent already runs on an LLM (e.g., your in-house Qwen), no separate Qwen API call is needed. · 스킬이 전체 워크플로를 수행하며, 에이전트 자체가 LLM(사내 Qwen 등) 위에서 돌므로 별도 Qwen 호출이 필요 없습니다.

## Constraints · 제약

- **xlwings only.** `openpyxl`, `pandas.read_excel`/`ExcelFile` are hard-banned and enforced by `tests/test_no_openpyxl.py`. · 하드 금지, 가드 테스트로 강제.
- Unattended inference uses the in-house **Qwen** backend (`--backend qwen`), configured via `XLTIDY_QWEN_BASE_URL`, `XLTIDY_QWEN_API_KEY`, `XLTIDY_QWEN_MODEL`. · 무인 추론은 사내 Qwen 옵션.
- Pivot extraction supports a **single data field** in v1 (multi-field pivots warn and use the first). · 피벗은 v1에서 단일 데이터 필드.

## Testing · 테스트

```powershell
python -m pytest -m "not excel"   # core (pure) — no Excel needed · 코어, Excel 불필요
python -m pytest -m excel         # COM (extract/pivot/e2e) — desktop Excel required · 데스크톱 Excel 필요
```

## Project structure · 프로젝트 구조

```
src/xltidy/
  coords.py       A1 <-> (row, col)
  models.py       Cell, MergedRange, CellGrid (value_filled = merge->anchor), SheetInfo
  encode.py       CellGrid -> compact text for the LLM (numbers masked as #num)
  spec.py         TemplateSpec / SheetSpec / TableSpec (kind: table | pivot)
  reconcile.py    table subtotal==sum · pivot data==grand total
  apply.py        apply_table / finalize_pivot / apply_workbook (DI)
  dbio.py         write_tables -> per-workbook folder of CSV/Parquet
  consolidate.py  detect_drift (sheet/column/region) + consolidate
  config.py       XLTIDY_QWEN_* env
  infer.py        agent prompt builder + optional Qwen backend
  extract.py      xlwings: list_sheets (incl. hidden) + extract
  pivot.py        native pivot extraction via COM PivotTables
  cli.py          typer CLI
skills/excel-to-db/SKILL.md   opencode / Claude skill
docs/superpowers/             design spec + implementation plan
```

## Roadmap · 로드맵

SQL DB adapters (SQLite/Postgres), heuristic table auto-detection, MCP server, server/headless batch, natural-language → SQL, multi-data-field pivots. · SQL 어댑터, 휴리스틱 표 감지, MCP 서버, 서버 배치, 자연어→SQL, 다중 데이터필드 피벗.

## References · 참고

- Design spec · 설계 스펙: [`docs/superpowers/specs/2026-06-10-xltidy-excel-to-db-design.md`](docs/superpowers/specs/2026-06-10-xltidy-excel-to-db-design.md)
- Implementation plan · 구현 계획: [`docs/superpowers/plans/2026-06-10-xltidy.md`](docs/superpowers/plans/2026-06-10-xltidy.md)
- Inspired by [exstruct](https://github.com/harumiWeb/exstruct) (Excel → structured JSON for LLM/RAG).

## License

[MIT](LICENSE) © 2026 Judy-0509
