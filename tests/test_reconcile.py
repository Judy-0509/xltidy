from tests.fixtures import sample_grid
from xltidy.models import Cell
from xltidy.spec import TemplateSpec, sample_spec_dict
from xltidy.reconcile import reconcile_table, reconcile_pivot, ReconcileReport


def _table():
    return TemplateSpec.model_validate(sample_spec_dict()).sheets[0].tables[0]


def test_reconcile_table_ok():
    assert reconcile_table(sample_grid(), _table()) == []


def test_reconcile_table_detects_bad_subtotal():
    g = sample_grid()
    g.cells = [c for c in g.cells if not (c.row == 8 and c.col == 3)]
    g.cells.append(Cell(row=8, col=3, value=999.0))
    issues = reconcile_table(g, _table())
    assert any("합계" in i for i in issues)


def test_reconcile_pivot():
    assert reconcile_pivot(42.0, 42.0, "p") == []
    assert reconcile_pivot(40.0, 42.0, "p")  # 불일치 → 이슈
