from __future__ import annotations

import glob as _glob
import json
from pathlib import Path

import typer
import yaml
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
def encode(file: str = typer.Option(None), grid: str = typer.Option(None), sheet: str = typer.Option(None),
           head: int = typer.Option(40, "--head", help="머리 N행만 인코딩(0=전체)"),
           tail: int = typer.Option(10, "--tail", help="꼬리 N행만 인코딩(0=전체)")):
    typer.echo(_encode(_grid_from(file, grid, sheet), head_rows=head, tail_rows=tail))


@app.command()
def infer(file: str = typer.Argument(None), grid: str = typer.Option(None), sheet: str = typer.Option(None),
          backend: str = typer.Option("agent"), out: str = typer.Option(None),
          head: int = typer.Option(40, "--head", help="구조 추론용 머리 N행만 프롬프트에 실음(0=전체)"),
          tail: int = typer.Option(10, "--tail", help="꼬리 N행만 프롬프트에 실음(0=전체)")):
    g = _grid_from(file, grid, sheet)
    enc = _encode(g, head_rows=head, tail_rows=tail)
    fname = file or grid or ""
    if backend == "agent":
        typer.echo(build_agent_prompt(enc, filename=fname))
        rprint("[yellow]# 위 인코딩으로 spec.yaml 작성(골격: `xltidy sample-spec`) 후 `xltidy spec-validate`로 검증. 피벗 시트는 kind:pivot.[/]")
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
    from pydantic import ValidationError
    try:
        sp = TemplateSpec.from_yaml(spec)
    except (ValidationError, yaml.YAMLError, FileNotFoundError) as e:
        rprint(f"[red]FAIL[/] 스펙 형식 오류: {e}")
        raise typer.Exit(1)
    issues: list[str] = []
    if against:
        issues = sp.validate_against({sheet or "": _grid_from(against, None, sheet)})
    if issues:
        for i in issues:
            rprint(f"[red]FAIL[/] {i}")
        raise typer.Exit(1)
    rprint("[green]OK spec valid[/]")


@app.command("sample-spec")
def sample_spec():
    """유효한 TemplateSpec YAML 골격을 출력(에이전트가 복사해 채우는 용도)."""
    from .spec import sample_spec_dict
    typer.echo(yaml.safe_dump(sample_spec_dict(), allow_unicode=True, sort_keys=False))


@app.command()
def apply(spec: str, file: str = typer.Option(None), grid: str = typer.Option(None),
          sheet: str = typer.Option(None), period: str = typer.Option(None),
          out_dir: str = typer.Option(...), fmt: str = typer.Option("csv", "--format"),
          verify: bool = typer.Option(True, "--verify/--no-verify",
                                      help="count check + random sample round-trip"),
          sample: int = typer.Option(50, "--sample", help="sampled cells per table (0 = all)")):
    sp = TemplateSpec.from_yaml(spec)
    vs = None if sample <= 0 else sample  # <=0 = check all
    if grid:  # 단일 grid 주입(테스트/단일시트 편의)
        g = _grid_from(None, grid, sheet)
        res = apply_workbook(grid, sp, sheet_extractor=lambda p, s: g,
                             pivot_extractor=lambda p, s, n: (_empty_pivot(), None),
                             list_sheets_fn=lambda p: [], period=period, filename=grid,
                             verify=verify, verify_sample=vs)
    else:
        res = apply_workbook(file, sp, period=period, filename=file,
                             verify=verify, verify_sample=vs)
    paths = write_tables(res.tables, out_dir, fmt=fmt)
    if not res.reconcile.ok:
        for i in res.reconcile.issues:
            rprint(f"[red]reconcile FAIL[/] {i}")
    for i in res.verify:
        rprint(f"[red]verify FAIL[/] {i}")
    rprint(f"[green]wrote[/] {len(paths)} tables -> {out_dir}")
    ok = res.reconcile.ok and not res.verify
    raise typer.Exit(0 if ok else 2)


@app.command()
def consolidate(spec: str, files: str, out_dir: str = typer.Option(...),
                fmt: str = typer.Option("csv", "--format"), on_drift: str = typer.Option("stop"),
                verify: bool = typer.Option(True, "--verify/--no-verify",
                                            help="count check + random sample round-trip"),
                sample: int = typer.Option(50, "--sample", help="sampled cells per table (0 = all)")):
    from .session import open_excel_session
    sp = TemplateSpec.from_yaml(spec)
    has_pivot = any(t.kind != "table" for s in sp.sheets for t in s.tables)
    paths = sorted(_glob.glob(files))
    res = _consolidate(paths, sp,
                       session_factory=lambda p: open_excel_session(p, read_only=not has_pivot),
                       on_drift=on_drift, verify=verify,
                       verify_sample=(None if sample <= 0 else sample))
    for path, drift in res.drift_by_file.items():
        rprint(f"[yellow]drift[/] {path}: {drift}")
    for path, vissues in res.verify_by_file.items():
        rprint(f"[red]verify FAIL[/] {path}: {vissues}")
    for issue in res.period_issues:
        rprint(f"[red]period FAIL[/] {issue}")
    written = write_tables(res.tables, out_dir, fmt=fmt)
    rprint(f"[green]wrote[/] {len(written)} tables -> {out_dir}")
    ok = not res.drift_by_file and not res.verify_by_file and not res.period_issues
    raise typer.Exit(0 if ok else 2)


def _empty_pivot():
    import pandas as pd
    return pd.DataFrame({"value": []})


if __name__ == "__main__":
    app()
