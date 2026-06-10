from typer.testing import CliRunner
from xltidy.cli import app
from tests.fixtures import sample_grid

runner = CliRunner()


def test_encode_cmd(tmp_path):
    gp = tmp_path / "grid.json"; gp.write_text(sample_grid().model_dump_json(), encoding="utf-8")
    r = runner.invoke(app, ["encode", "--grid", str(gp)])
    assert r.exit_code == 0 and "SHEET: 데이터" in r.stdout


def test_apply_cmd_writes_folder(tmp_path):
    from xltidy.spec import TemplateSpec, sample_spec_dict
    gp = tmp_path / "grid.json"; gp.write_text(sample_grid().model_dump_json(), encoding="utf-8")
    sp = tmp_path / "spec.yaml"; TemplateSpec.model_validate(sample_spec_dict()).to_yaml(sp)
    out = tmp_path / "db"
    r = runner.invoke(app, ["apply", str(sp), "--grid", str(gp), "--sheet", "데이터",
                            "--period", "2024Q1", "--out-dir", str(out)])
    assert r.exit_code == 0, r.stdout
    assert (out / "by_industry.csv").exists()
