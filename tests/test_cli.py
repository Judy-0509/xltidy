from typer.testing import CliRunner

from tests.fixtures import sample_grid
from moa.cli import app

runner = CliRunner()


def test_encode_cmd(tmp_path):
    gp = tmp_path / "grid.json"; gp.write_text(sample_grid().model_dump_json(), encoding="utf-8")
    r = runner.invoke(app, ["encode", "--grid", str(gp)])
    assert r.exit_code == 0 and "SHEET:" in r.stdout


def test_sample_spec_cmd_roundtrips():
    import yaml
    from moa.spec import TemplateSpec
    r = runner.invoke(app, ["sample-spec"])
    assert r.exit_code == 0, r.stdout
    TemplateSpec.model_validate(yaml.safe_load(r.stdout))  # valid skeleton to copy


def test_infer_accepts_positional_file(tmp_path):
    # regression: `moa infer <path>` (positional) must not error with
    # "Got unexpected extra argument"; --grid still takes precedence.
    gp = tmp_path / "grid.json"; gp.write_text(sample_grid().model_dump_json(), encoding="utf-8")
    r = runner.invoke(app, ["infer", "Some Workbook, April 2026.xlsx",
                            "--grid", str(gp), "--backend", "agent"])
    assert r.exit_code == 0, r.stdout
    assert "unexpected extra argument" not in r.stdout.lower()
    assert "SHEET" in r.stdout  # agent prompt emitted from the grid


def test_apply_cmd_writes_folder(tmp_path):
    from moa.spec import TemplateSpec, sample_spec_dict
    gp = tmp_path / "grid.json"; gp.write_text(sample_grid().model_dump_json(), encoding="utf-8")
    sp = tmp_path / "spec.yaml"; TemplateSpec.model_validate(sample_spec_dict()).to_yaml(sp)
    out = tmp_path / "db"
    # no --verify flag: verification now runs by DEFAULT and must pass on clean data
    r = runner.invoke(app, ["apply", str(sp), "--grid", str(gp), "--sheet", "Data",
                            "--version", "2024Q1", "--out-dir", str(out)])
    assert r.exit_code == 0, r.stdout
    assert (out / "db.csv").exists()


def test_apply_cmd_per_table_writes_named_file(tmp_path):
    from moa.spec import TemplateSpec, sample_spec_dict
    gp = tmp_path / "grid.json"; gp.write_text(sample_grid().model_dump_json(), encoding="utf-8")
    sp = tmp_path / "spec.yaml"; TemplateSpec.model_validate(sample_spec_dict()).to_yaml(sp)
    out = tmp_path / "db_per_table"
    r = runner.invoke(app, ["apply", str(sp), "--grid", str(gp), "--sheet", "Data",
                            "--version", "2024Q1", "--out-dir", str(out), "--per-table"])
    assert r.exit_code == 0, r.stdout
    assert (out / "by_industry.csv").exists()


def test_consolidate_empty_glob_fails(tmp_path):
    # Empty glob must fail explicitly, not silently write zero tables.
    from moa.spec import TemplateSpec, sample_spec_dict
    sp = tmp_path / "spec.yaml"; TemplateSpec.model_validate(sample_spec_dict()).to_yaml(sp)
    r = runner.invoke(app, ["consolidate", str(sp), str(tmp_path / "nope_*.xlsx"),
                            "--out-dir", str(tmp_path / "db")])
    assert r.exit_code != 0
    assert "no files match" in (r.stdout + str(r.exception or ""))


def test_apply_no_verify_turns_it_off(tmp_path):
    # --verify is default-on; --no-verify must parse and still write the folder.
    from moa.spec import TemplateSpec, sample_spec_dict
    gp = tmp_path / "grid.json"; gp.write_text(sample_grid().model_dump_json(), encoding="utf-8")
    sp = tmp_path / "spec.yaml"; TemplateSpec.model_validate(sample_spec_dict()).to_yaml(sp)
    out = tmp_path / "db2"
    r = runner.invoke(app, ["apply", str(sp), "--grid", str(gp), "--sheet", "Data",
                            "--version", "2024Q1", "--out-dir", str(out), "--no-verify"])
    assert r.exit_code == 0, r.stdout
    assert (out / "db.csv").exists()
