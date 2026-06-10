from tests.fixtures import sample_grid
from xltidy.apply import apply_table
from xltidy.spec import TemplateSpec, sample_spec_dict
from xltidy.verify import verify_table


def _table():
    return TemplateSpec.model_validate(sample_spec_dict()).sheets[0].tables[0]


def test_verify_ok_full_sample():
    t = _table()
    frame = apply_table(sample_grid(), t, period="2024Q1")
    # sample=None -> check every cell round-tripped
    assert verify_table(sample_grid(), t, frame, period="2024Q1", sample=None) == []


def test_verify_flags_count_mismatch():
    t = _table()
    frame = apply_table(sample_grid(), t, period="2024Q1").iloc[:2]  # drop rows
    issues = verify_table(sample_grid(), t, frame, period="2024Q1", sample=None)
    assert any("row count" in i for i in issues)


def test_verify_flags_value_mismatch():
    t = _table()
    frame = apply_table(sample_grid(), t, period="2024Q1").copy()
    frame.loc[0, "value"] = 99999.0  # corrupt one output value
    issues = verify_table(sample_grid(), t, frame, period="2024Q1", sample=None)
    assert any("mismatch" in i for i in issues)
