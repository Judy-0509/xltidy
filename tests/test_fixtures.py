from tests.fixtures import sample_grid, sample_pivot_raw


def test_fixtures():
    assert sample_grid().at(8, 3) == 300.0
    f, gt = sample_pivot_raw()
    assert f["value"].sum() == gt
