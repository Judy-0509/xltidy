from __future__ import annotations

import glob as _glob
import json
from pathlib import Path

import typer
from rich import print as rprint
from rich.table import Table

from .apply import apply_workbook
from .consolidate import consolidate as _consolidate
from .dbio import write_tables
from .encode import encode as _encode
from .infer import build_agent_prompt
from .models import CellGrid
from .spec import TemplateSpec

app = typer.Typer(help="xlwings-only Excel -> tidy per-workbook folder of tables")


def _grid_from(file: str | None, grid: str | None, sheet: str | None) -> CellGrid:
    if grid:
        return CellGrid.model_validate_json(Path(grid).read_text(encoding="utf-8"))
    if file:
        from .extract import extract
        return extract(file, sheet=sheet)
    raise typer.BadParameter("--file 또는 --grid 필요")


@app.command()
def sheets(file: str, as_json: bool = typer.Option(False, "--json")):
    from .extract import list_sheets
    infos = list_sheets(file)
    if as_json:
        typer.echo(json.dumps([i.model_dump() for i in infos], ensure_ascii=False, indent=2))
        return
    t = Table("idx", "name", "visibility", "rows", "cols", "pivots")
    for i in infos:
        t.add_row(str(i.index), i.name, i.visibility, str(i.used_rows), str(i.used_cols), str(i.n_pivots))
    rprint(t)


@app.command()
def extract(file: str, sheet: str = typer.Option(None), out: str = typer.Option(None)):
    from .extract import extract as ex
    data = ex(file, sheet=sheet).model_dump_json()
    Path(out).write_text(data, encoding="utf-8") if out else typer.echo(data)


@app.command()
def encode(file: str = typer.Option(None), grid: str = typer.Option(None), sheet: str = typer.Option(None)):
    typer.echo(_encode(_grid_from(file, grid, sheet)))


@app.command()
def infer(file: str = typer.Option(None), grid: str = typer.Option(None), sheet: str = typer.Option(None),
          backend: str = typer.Option("agent"), out: str = typer.Option(None)):
    g = _grid_from(file, grid, sheet)
    enc = _encode(g)
    fname = file or grid or ""
    if backend == "agent":
        typer.echo(build_agent_prompt(enc, filename=fname))
        rprint("[yellow]# agent가 위로 spec.yaml 작성 후 `xltidy spec-validate`로 검증. 피벗 시트는 kind:pivot로 직접.[/]")
    elif backend == "qwen":
        from .config import load_settings
        from .infer import infer_with_qwen
        s = load_settings()
        spec = infer_with_qwen(enc, filename=fname, model=s.model or "qwen", base_url=s.base_url, api_key=s.api_key)
        text = spec.model_dump_json(indent=2)
        Path(out).write_text(text, encoding="utf-8") if out else typer.echo(text)
    else:
        raise typer.BadParameter("backend는 agent|qwen")


@app.command("spec-validate")
def spec_validate(spec: str, against: str = typer.Option(None), sheet: str = typer.Option(None)):
    sp = TemplateSpec.from_yaml(spec)
    issues: list[str] = []
    if against:
        issues = sp.validate_against({sheet or "": _grid_from(against, None, sheet)})
    if issues:
        for i in issues:
            rprint(f"[red]✗[/] {i}")
        raise typer.Exit(1)
    rprint("[green]✓ spec valid[/]")


@app.command()
def apply(spec: str, file: str = typer.Option(None), grid: str = typer.Option(None),
          sheet: str = typer.Option(None), period: str = typer.Option(None),
          out_dir: str = typer.Option(...), fmt: str = typer.Option("csv", "--format")):
    sp = TemplateSpec.from_yaml(spec)
    if grid:  # 단일 grid 주입(테스트/단일시트 편의)
        g = _grid_from(None, grid, sheet)
        res = apply_workbook(grid, sp, sheet_extractor=lambda p, s: g,
                             pivot_extractor=lambda p, s, n: (_empty_pivot(), None),
                             list_sheets_fn=lambda p: [], period=period, filename=grid)
    else:
        res = apply_workbook(file, sp, period=period, filename=file)
    paths = write_tables(res.tables, out_dir, fmt=fmt)
    if not res.reconcile.ok:
        for i in res.reconcile.issues:
            rprint(f"[red]reconcile ✗[/] {i}")
    rprint(f"[green]wrote[/] {len(paths)} tables → {out_dir}")
    raise typer.Exit(0 if res.reconcile.ok else 2)


@app.command()
def consolidate(spec: str, files: str, out_dir: str = typer.Option(...),
                fmt: str = typer.Option("csv", "--format"), on_drift: str = typer.Option("stop")):
    from .extract import extract as ex, list_sheets
    from .pivot import extract_pivot
    sp = TemplateSpec.from_yaml(spec)
    paths = sorted(_glob.glob(files))
    res = _consolidate(paths, sp, list_sheets_fn=list_sheets, sheet_extractor=ex,
                       pivot_extractor=extract_pivot, on_drift=on_drift)
    for path, drift in res.drift_by_file.items():
        rprint(f"[yellow]drift[/] {path}: {drift}")
    written = write_tables(res.tables, out_dir, fmt=fmt)
    rprint(f"[green]wrote[/] {len(written)} tables → {out_dir}")


def _empty_pivot():
    import pandas as pd
    return pd.DataFrame({"value": []})


if __name__ == "__main__":
    app()
