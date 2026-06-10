---
description: Excel을 xltidy로 DB화 — 시트 선택 → 스펙 작성 → apply → consolidate (xlwings only)
---

You convert the user's Excel workbook(s) into tidy, DB-ready tables using the **xltidy** CLI available in this project (`xltidy ...`). **You (this model) author the structure yourself — never call any external LLM/API.** Never use `openpyxl` or `pandas.read_excel` (hard policy). Use **PowerShell** syntax for any shell commands.

Workflow:

1. **Sheet selection first.** Run `xltidy sheets <file>` and show the user **all** sheets, including hidden / very_hidden. Ask which sheet(s) to DB-ify. One workbook → one output folder, with one table file per sheet/table.

2. **Author a reusable spec (once per template)** → save to `specs/<template_id>.yaml`:
   - **Table sheets**: run `xltidy infer <file> --sheet <sheet> --backend agent`, read the encoded sheet, and write the spec entry. Flatten multi-level headers into one `column_semantics` item per value column. Numbers are masked as `#num` — emit coordinates/structure only, **never transcribe values**.
   - **Pivot sheets**: do NOT infer — just add `kind: pivot` + `pivot_name` (null = first pivot) + `period`. Pivots are extracted natively via COM.

3. **Validate**: `xltidy spec-validate specs/<id>.yaml --against <file> --sheet <sheet>`.

4. **Apply one file**: `xltidy apply specs/<id>.yaml --file <file> --out-dir out/<period> --format csv`. Report any `reconcile` mismatch (table subtotal ≠ sum, pivot data ≠ grand total) and fix the spec.

5. **Consolidate monthly/quarterly versions**: `xltidy consolidate specs/<id>.yaml "<glob>" --out-dir merged --format parquet --on-drift stop`. Files flagged for drift (renamed headers, missing selected sheets) are excluded — bump the spec `version` if the template changed.

Target (file/dir/glob): $ARGUMENTS
