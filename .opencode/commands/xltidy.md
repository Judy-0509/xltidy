---
description: Excel 하나를 가리키면 알아서 시트 선택→스펙 작성→CSV DB 생성, 다른 버전 있으면 통합 (xltidy, xlwings only)
---

You turn the Excel workbook the user points at into a tidy, **CSV** DB using the **xltidy** CLI installed in the environment (`xltidy ...`). **You (this model) author the spec yourself — no external LLM/API needed.** Never use `openpyxl` or `pandas.read_excel`. Use **PowerShell** syntax for shell commands.

The target Excel path is: **$ARGUMENTS**
If that is empty, first ask the user for the Excel file path.

Do this conversationally, asking the user at the two decision points (★):

1. **List sheets and ★ ask which to use.** Run `xltidy sheets "<path>"` and show the user **every** sheet (including hidden / very_hidden) with its size and pivot count. Ask **which sheet(s) to DB-ify** (one workbook → one output folder, one CSV per sheet/table). Wait for their answer.

2. **Author the spec automatically.** Name it after the Excel **template**: `<id>` = the workbook filename with the **period/date dropped**, kebab-cased (e.g. `Market Pulse - Flagship Model Sales, April 2026.xlsx` → `market-pulse-flagship-model-sales`). Dropping the date lets every monthly/quarterly version share one spec for consolidation. Use that `<id>` for both `template_id` and the path `specs/<id>.yaml`.
   - **Table sheets**: run `xltidy infer "<path>" --sheet <sheet> --backend agent`, read the encoded sheet, and fill the spec entry yourself. Need the YAML shape? Run `xltidy sample-spec` for a valid skeleton to copy. Flatten multi-level headers into one `column_semantics` item per value column. Numbers are masked as `#num` — emit coordinates/structure only, **never transcribe values**. Large sheets are auto-sampled to the first 40 + last 10 value-rows (`--head/--tail` to adjust); the encoding's `DATA ROWS: a..b` is the table's true end, so set the region end to **b**, not the last visible row.
   - **Pivot sheets**: do NOT infer — just add `kind: pivot` + `pivot_name` (null = first) + `period`.

3. **Validate** the spec: `xltidy spec-validate specs/<id>.yaml --against "<path>" --sheet <sheet>`. Fix any issue.

4. **Build the CSV DB for this file (with verification).** Name the output folder after the workbook (one Excel = one DB folder) — use the file's stem, kebab-cased, e.g. `out/market-pulse-flagship-model-sales-april-2026`:
   ```powershell
   xltidy apply specs/<id>.yaml --file "<path>" --out-dir "out/<workbook-stem>" --format csv --verify --sample 100
   ```
   Report the output folder and any `reconcile FAIL` / `verify FAIL`. If something fails, fix the spec and re-run.

5. **★ Ask about other versions.** Ask the user: *"이 엑셀의 다른 기간(월별/분기별) 버전 파일이 더 있나요?"*
   - **If yes**: get their other file paths or a glob (e.g. `data\survey_2024Q*.xlsx`), then build **one combined DB across all versions** (each row tagged with its `period`):
     ```powershell
     xltidy consolidate specs/<id>.yaml "data\survey_2024Q*.xlsx" --out-dir "merged" --format csv --on-drift stop --verify --sample 100
     ```
     Report drift / verify per file. Drifted files (renamed headers, missing selected sheet) are excluded — if the template genuinely changed, bump the spec `version` and re-author for those.
   - **If no**: you're done — the single-file CSV folder from step 4 is the DB.

Always honor: xlwings only, LLM emits structure not values, COM features need Windows + Excel.
