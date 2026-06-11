---
description: Excel 하나를 가리키면 알아서 시트 선택→스펙 작성→CSV DB 생성, 다른 버전 있으면 통합 (moa, xlwings only)
---

You turn the Excel workbook the user points at into a tidy, **CSV** DB using the **moa** CLI installed in the environment (`moa ...`; the legacy `xltidy` command runs the same thing). **You (this model) author the spec yourself — no external LLM/API needed.** Never use `openpyxl` or `pandas.read_excel`. Use **PowerShell** syntax for shell commands.

The target Excel path is: **$ARGUMENTS**
If that is empty, first ask the user for the Excel file path.

Do this conversationally, asking the user at the two decision points (★):

1. **List sheets and ★ ask which to use.** Run `moa sheets "<path>"` and show the user **every** sheet (including hidden / very_hidden) with its size and pivot count. Ask **which sheet(s) to DB-ify** (one workbook → one output folder with one combined `db.csv`/`db.parquet` by default; `table` distinguishes tables). Wait for their answer.

2. **Author the spec automatically.** Name it after the Excel **template**: `<id>` = the workbook filename with the **version/date dropped**, kebab-cased (e.g. `Market Pulse - Flagship Model Sales, April 2026.xlsx` → `market-pulse-flagship-model-sales`). Dropping the date lets every monthly/quarterly version share one spec for consolidation. Use that `<id>` for both `template_id` and the path `specs/<id>.yaml`.
   - **Table sheets**: run `moa infer "<path>" --sheet <sheet> --backend agent`, read the encoded sheet, and fill the spec entry yourself. Need the YAML shape? Run `moa sample-spec` for a valid skeleton to copy. Flatten multi-level headers into one `column_semantics` item per value column. Numbers/formulas are masked as `#num` / `#num=FORMULA` — emit coordinates/structure only, **never transcribe values**. The encoding also includes `REGIONS`, `DATA COLS`, `COL FORMATS`, and structure hints (`LISTOBJECT`/`PIVOT`/`NAME`); `REGIONS` means multiple tables may be present on one sheet. Large sheets are auto-sampled to the first 40 + last 10 value-rows (`--head/--tail` to adjust); the encoding's `DATA ROWS: a..b` is the table's true end, so set the region end to **b**, not the last visible row.
   - **Pivot sheets**: do NOT infer — just add `kind: pivot` + `pivot_name` (null = first) + `version`. (Extraction clears all filters — report/page filters, hidden items, slicers — to return the FULL dataset.)
   - In every table spec, `version:` defines row provenance; the output column defaults to `version`, and `moa apply --version <value>` can override resolution for a single file.

3. **Validate** the spec: `moa spec-validate specs/<id>.yaml --against "<path>" --sheet <sheet>`. Fix any issue.

4. **Build the CSV DB for this file (verification runs by default).** Name the output folder after the workbook (one Excel = one DB folder) — use the file's stem, kebab-cased, e.g. `out/market-pulse-flagship-model-sales-april-2026`:
   ```powershell
   moa apply specs/<id>.yaml --file "<path>" --out-dir "out/<workbook-stem>" --format csv
   ```
   Report the output folder and any `reconcile FAIL` / `verify FAIL`. If something fails, fix the spec and re-run. (`--no-verify` to skip the output check.)

5. **★ Ask about other versions.** Ask the user: *"이 엑셀의 다른 기간(월별/분기별) 버전 파일이 더 있나요?"*
   - **If yes**: get their other file paths or a glob (e.g. `data\survey_2024Q*.xlsx`), then build **one combined DB across all versions** (each row tagged with its `version`; all tables go into one `db.csv`/`db.parquet` by default with a `table` column and unioned columns, while `--per-table` restores one file per table):
     ```powershell
     moa consolidate specs/<id>.yaml "data\survey_2024Q*.xlsx" --out-dir "merged" --format csv --on-drift stop
     ```
     Relay the visual per-file progress and final summary table, then report drift / verify / **version** issues per file. Each file must yield a **unique** `version` — filename regex uses the FULL match (e.g. pattern `(\d{4})Q([1-4])` on `survey_2024Q1.xlsx` yields `2024Q1`), and collisions or `None` are reported as `version FAIL`; fix the version pattern/cell. Drifted files (renamed headers, missing selected sheet) are excluded — if the template genuinely changed, bump the spec `version` and re-author for those.
   - **If no**: you're done — the single-file CSV folder from step 4 is the DB.

For unattended Qwen inference, prefer `MOA_QWEN_*`; `XLTIDY_QWEN_*` is only a legacy alias.

Always honor: xlwings only, LLM emits structure not values, COM features need Windows + Excel.
