from moa.spec import TemplateSpec, sample_spec_dict
from tests.fixtures import sample_grid


def test_load_validate_ok():
    spec = TemplateSpec.model_validate(sample_spec_dict())
    assert spec.sheets[0].tables[0].index_columns[0].name == "industry"
    assert spec.validate_against({"데이터": sample_grid()}) == []


def test_yaml_roundtrip(tmp_path):
    spec = TemplateSpec.model_validate(sample_spec_dict())
    p = tmp_path / "s.yaml"
    spec.to_yaml(p)
    s2 = TemplateSpec.from_yaml(p)
    assert s2.sheets[0].tables[0].value_block.cols == ["C", "D"]


def test_validate_flags_out_of_bounds():
    d = sample_spec_dict()
    d["sheets"][0]["tables"][0]["region"]["end"] = "D999"
    spec = TemplateSpec.model_validate(d)
    issues = spec.validate_against({"데이터": sample_grid()})
    assert any("region" in i.lower() for i in issues)


def test_pivot_table_minimal():
    d = sample_spec_dict()
    d["sheets"].append({
        "sheet_match": {"by": "name", "value": "피벗"},
        "tables": [{"name": "p", "kind": "pivot", "pivot_name": None,
                    "period": {"source": {"from": "filename", "pattern": r"(\d{4})Q([1-4])"}, "name": "period"}}],
    })
    spec = TemplateSpec.model_validate(d)
    assert spec.sheets[1].tables[0].kind == "pivot"
