<p align="center">
  <img alt="Moa — scattered, messy spreadsheet cells gathering into one clean, tidy table" src="docs/assets/moa-banner.png" width="100%">
</p>

<p align="center"><b>English</b> · <a href="README.ko.md">한국어</a></p>

<p align="center">
  <b>Turn messy survey-institution Excel into tidy, DB-ready tables — using <code>xlwings</code> only.</b>
</p>

<p align="center">
  <a href="https://github.com/Judy-0509/moa/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/Judy-0509/moa/actions/workflows/ci.yml/badge.svg"></a>
  <img alt="python" src="https://img.shields.io/badge/python-3.10%2B-blue">
  <img alt="excel io" src="https://img.shields.io/badge/Excel%20I%2FO-xlwings%20only-success">
  <img alt="no openpyxl" src="https://img.shields.io/badge/openpyxl-banned-critical">
  <img alt="platform" src="https://img.shields.io/badge/platform-Windows%20%2B%20Excel-lightgrey">
  <img alt="license" src="https://img.shields.io/badge/license-MIT-green">
  <a href="https://github.com/Judy-0509/moa/releases"><img alt="release" src="https://img.shields.io/github/v/release/Judy-0509/moa"></a>
</p>

<p align="center"><sub><b>Moa</b> — from the Korean verb "to gather": gather scattered Excel into one clean DB.<br>
CLI command is <code>moa</code> (the legacy <code>xltidy</code> command still works; the package imports as <code>moa</code>).</sub></p>

<p align="center"><sub>🚧 <b>Status:</b> early stage (v0.1.x). The CLI works end-to-end today; the API may still change. Issues &amp; feedback welcome.</sub></p>

---

## What is this?

**Moa** converts complex, real-world Excel workbooks from research/survey institutions into **tidy, DB-ready tables** — without ever touching `openpyxl`.

Survey Excel files are hard: merged cells, multi-level headers, subtotal rows, pivot tables, and the *same template repeated every month/quarter*. Moa reads them with a **live Excel via `xlwings`** (so formula values, formats, and pivots are real), then applies a **reusable `TemplateSpec`** deterministically:

- **One workbook → one output folder** (one tidy CSV/Parquet file per sheet/table) — *"1 Excel = 1 DB"*.
- **Regular tables**: an LLM (your in-house Qwen, or an [opencode](https://opencode.ai)/Claude agent) infers the *structure only* once per template; the actual numbers are read deterministically from the spreadsheet — **the LLM never transcribes values**.
- **Pivot tables**: extracted natively via Excel COM (`PivotTables`) — no LLM needed. **All filters are cleared first** (report/page filters, hidden row/column items, slicers/timelines) so you get the **full dataset**, not the filtered view.
- **Version consolidation**: stack monthly/quarterly files into one time series with a `period` dimension, and **flag drift** when a new file no longer matches the template (renamed columns, missing sheets, shifted regions) or when two versions resolve to the **same / unresolved period**.
- **Integrity first**: every run reconciles (table subtotal == sum of components; pivot data sum == grand total), and **output verification runs by default** (row-count + random sample round-trip).
- **One open per workbook**: each file is opened once and the headless Excel process is always terminated — no orphaned `EXCEL.EXE`.

Designed for **on-prem / in-house** use (data never leaves the building), so it pairs with a self-hosted Qwen and ships as an **opencode/Claude skill**.

<p align="center">
  <img alt="Before: survey Excel with merged cells, multi-level headers and subtotal rows. After: tidy long-format table, reconciled and verified." src="docs/assets/moa-before-after.svg" width="100%">
</p>

## Why

I'm a market-intelligence analyst. Every month, research firms send Excel files that are beautiful to look at and terrible to analyze: merged cells, three-level headers, subtotal rows mixed into the data — and a slightly different layout each quarter. Cleaning them by hand took hours, and one mis-copied cell could silently corrupt a whole time series. Moa is the tool I wished existed: describe the template **once**, then turn every future file into the same clean, **verified** table.

---

## How it works

```
0) moa sheets <file>  ─▶  list ALL sheets (incl. hidden)  ─▶  [you pick the sheets to turn into DB tables]
                                                            │
   per selected sheet:                                      ▼
   kind=table:  extract (xlwings) ─▶ CellGrid ─ encode ─▶ [agent/Qwen] ─▶ TableSpec   (structure only)
   kind=pivot:  pivot (COM, filters cleared) ──────────────────────────▶ TableSpec   (LLM-free)
                                                            │
   each monthly/quarterly file + TemplateSpec ─ apply ─▶ {table: tidy long(+period)} + reconcile + verify
                                                            │
   many files ─ consolidate ─▶ per-table period stacking + drift/period checks ─▶ output folder (<table>.csv/.parquet)
```

The LLM only ever produces the `TemplateSpec` (coordinates and structure). Deterministic code reads the real values.

> All commands below are **PowerShell** (Windows).

## Install

```powershell
python -m pip install -e ".[dev]"
# + Parquet output:        python -m pip install -e ".[dev,parquet]"
# + in-house Qwen backend: python -m pip install -e ".[dev,parquet,qwen]"
```

This installs the **`moa`** command (and the legacy `xltidy` alias). Requirements: **Python 3.10+**, and for COM features (`sheets` / `extract` / pivot in `apply`·`consolidate`) a **Windows machine with Microsoft Excel installed**.

## In-house setup (PowerShell)

```powershell
# 1) Clone the repo
git clone https://github.com/Judy-0509/moa.git
Set-Location moa

# 2) (recommended) isolated environment
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 3) Install (with Parquet + in-house Qwen backend)
python -m pip install -e ".[dev,parquet,qwen]"

# 4) Point at your in-house Qwen (OpenAI-compatible)
#    this session only:
$env:MOA_QWEN_BASE_URL = "http://qwen.example.internal/v1"
$env:MOA_QWEN_API_KEY  = "your-internal-key"
$env:MOA_QWEN_MODEL    = "qwen2.5-72b-instruct"
#    persist for your user (new shells):
[Environment]::SetEnvironmentVariable("MOA_QWEN_BASE_URL", "http://qwen.example.internal/v1", "User")
[Environment]::SetEnvironmentVariable("MOA_QWEN_MODEL", "qwen2.5-72b-instruct", "User")

# 5) Smoke-test the install
python -m pytest -m "not excel" -q     # core, no Excel
python -m pytest -m excel -q           # COM (needs desktop Excel)
```

> Using it through an **opencode/Claude agent** instead? You don't need the `qwen` backend or the env vars — the agent's own LLM authors the spec. Just install (step 3 without `,qwen`) and register the skill (see below).

## Quickstart

```powershell
# 0) Choose sheets — lists ALL sheets incl. hidden / very-hidden
moa sheets report_2024Q1.xlsx

# 1) Author a TemplateSpec once per template
#    table sheets: get an encoding for the agent to fill the spec
moa infer report_2024Q1.xlsx --sheet 데이터 --backend agent
moa sample-spec            # need the YAML shape? prints a valid skeleton
#    pivot sheets: just write `kind: pivot` + pivot_name + period in the spec
#    (filters are cleared automatically at extraction time)

# 2) Validate
moa spec-validate specs/employment.yaml --against report_2024Q1.xlsx --sheet 데이터

# 3) Apply one workbook → one folder of tidy tables
moa apply specs/employment.yaml --file report_2024Q1.xlsx --out-dir out/2024Q1 --format csv

# 4) Consolidate monthly/quarterly versions
moa consolidate specs/employment.yaml "data/2024*.xlsx" --out-dir merged --format parquet --on-drift stop
```

`reconcile` mismatches (subtotal ≠ sum, pivot data ≠ grand total), `drift` (renamed headers, missing selected sheets), and `period` collisions (two versions resolving to the same / unresolved period) are reported; drifted files are **excluded** under `--on-drift stop`.

**Output verification runs by default** — a row-count check plus a random **sample round-trip** (source cells → output), which works even when the sheet has no subtotals. `--sample N` sets the sampled cell count (`0` = check all); `--no-verify` skips it.

## Use as an opencode / Claude skill

The skill ships at [`.opencode/skills/moa/SKILL.md`](.opencode/skills/moa/SKILL.md) (folder name = skill `name`), so opening this repo in opencode **auto-loads** it as a project skill. To use it elsewhere, copy the folder (PowerShell):

```powershell
# opencode (global)
$dst = "$env:USERPROFILE\.config\opencode\skills\moa"
New-Item -ItemType Directory -Force -Path $dst | Out-Null
Copy-Item ".\.opencode\skills\moa\SKILL.md" "$dst\SKILL.md" -Force

# Claude (per-project)
New-Item -ItemType Directory -Force -Path ".\.claude\skills\moa" | Out-Null
Copy-Item ".\.opencode\skills\moa\SKILL.md" ".\.claude\skills\moa\SKILL.md" -Force
```

Then in opencode the agent calls `skill({ name: "moa" })`.

### `/moa` slash command

A **skill** is auto-invoked by the agent (you ask in natural language, the model loads it). To trigger the workflow explicitly by typing **`/moa`**, use a **command** instead — a self-contained one ships at [`.opencode/commands/moa.md`](.opencode/commands/moa.md).

- Open this repo in opencode → `/moa <file>` works out of the box (no skill copy, no Qwen API — opencode's own model runs it).
- Make it available everywhere:

```powershell
$cmd = "$env:USERPROFILE\.config\opencode\commands"
New-Item -ItemType Directory -Force -Path $cmd | Out-Null
Copy-Item ".\.opencode\commands\moa.md" "$cmd\moa.md" -Force
```

The skill drives the whole workflow (sheet selection → spec authoring → apply → consolidate). Since the agent already runs on an LLM (e.g., your in-house Qwen), no separate Qwen API call is needed.

## Constraints

- **xlwings only.** `openpyxl`, `pandas.read_excel`/`ExcelFile` are hard-banned and enforced by `tests/test_no_openpyxl.py`.
- Unattended inference uses the in-house **Qwen** backend (`--backend qwen`), configured via `MOA_QWEN_BASE_URL`, `MOA_QWEN_API_KEY`, `MOA_QWEN_MODEL` (legacy `XLTIDY_*` names still work).
- Pivot extraction clears all filters and supports a **single data field** in v1 (multi-field pivots warn and use the first).
- Merge detection covers label/header anchors (text/date); a merge anchored by a **bare number**, and **merged numeric body** cells, are not supported (value cells are read per-cell).

## Testing

```powershell
python -m pytest -m "not excel"   # core (pure) — no Excel needed
python -m pytest -m excel         # COM (extract/pivot/e2e) — desktop Excel required
```

## Project structure

```
src/moa/                        (package imports as `moa`; CLI command is `moa`)
  coords.py       A1 <-> (row, col)
  models.py       Cell, MergedRange, CellGrid (value_filled = merge->anchor), SheetInfo
  encode.py       CellGrid -> compact text for the LLM (numbers masked as #num; large sheets head/tail sampled)
  spec.py         TemplateSpec / SheetSpec / TableSpec (kind: table | pivot)
  reconcile.py    table subtotal==sum · pivot data==grand total
  verify.py       independent output check (row count + random sample round-trip)
  apply.py        apply_table / finalize_pivot / apply_session / apply_workbook
  dbio.py         write_tables -> per-workbook folder of CSV/Parquet
  consolidate.py  detect_drift (sheet/column/region) + period-collision guard + consolidate
  config.py       MOA_QWEN_* env
  infer.py        agent prompt builder + optional Qwen backend
  _xl.py          headless Excel lifecycle: new_app / quit_app (kill backstop) / open_book
  session.py      ExcelSession (one open workbook, per-sheet grid cache) + FnSession adapter
  extract.py      xlwings: list_sheets (incl. hidden) + extract / grid_from_sheet
  pivot.py        native pivot extraction via COM PivotTables (clears all filters)
  cli.py          typer CLI
.opencode/skills/moa/SKILL.md      opencode skill (auto-loaded; folder = skill name)
.opencode/commands/moa.md          /moa slash command
docs/superpowers/                  design spec + implementation plan
```

## Roadmap

SQL DB adapters (SQLite/Postgres), heuristic table auto-detection, MCP server, server/headless batch, natural-language → SQL, multi-data-field pivots.

## References

- Design spec: [`docs/superpowers/specs/2026-06-10-xltidy-excel-to-db-design.md`](docs/superpowers/specs/2026-06-10-xltidy-excel-to-db-design.md)
- Implementation plan: [`docs/superpowers/plans/2026-06-10-xltidy.md`](docs/superpowers/plans/2026-06-10-xltidy.md)
- Inspired by [exstruct](https://github.com/harumiWeb/exstruct) (Excel → structured JSON for LLM/RAG).

## License

[MIT](LICENSE) © 2026 Judy-0509
