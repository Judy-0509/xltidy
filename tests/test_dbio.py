import pandas as pd
from moa.dbio import write_tables


def test_write_tables_csv(tmp_path):
    tables = {"t1": pd.DataFrame({"a": [1, 2]}), "t2": pd.DataFrame({"b": [3]})}
    paths = write_tables(tables, str(tmp_path / "out"), fmt="csv")
    assert (tmp_path / "out" / "t1.csv").exists()
    assert (tmp_path / "out" / "t2.csv").exists()
    assert len(paths) == 2
    assert "a" in (tmp_path / "out" / "t1.csv").read_text(encoding="utf-8")
