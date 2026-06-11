import pandas as pd

from moa.dbio import write_tables


def test_write_tables_csv_single_table_default(tmp_path):
    t1 = pd.DataFrame({"a": [1, 2]})
    t2 = pd.DataFrame({"b": [3]})
    tables = {"t1": t1, "t2": t2}

    paths = write_tables(tables, str(tmp_path / "out"), fmt="csv")

    assert paths == [str(tmp_path / "out" / "db.csv")]
    out = pd.read_csv(tmp_path / "out" / "db.csv", encoding="utf-8-sig")
    assert list(out.columns) == ["table", "a", "b"]
    assert list(out["table"]) == ["t1", "t1", "t2"]
    assert pd.isna(out.loc[0, "b"])
    assert pd.isna(out.loc[2, "a"])
    assert list(t1.columns) == ["a"]
    assert list(t2.columns) == ["b"]


def test_write_tables_csv_per_table(tmp_path):
    tables = {"t1": pd.DataFrame({"a": [1, 2]}), "t2": pd.DataFrame({"b": [3]})}
    paths = write_tables(tables, str(tmp_path / "out"), fmt="csv", single_table=False)
    assert (tmp_path / "out" / "t1.csv").exists()
    assert (tmp_path / "out" / "t2.csv").exists()
    assert len(paths) == 2
    assert "a" in (tmp_path / "out" / "t1.csv").read_text(encoding="utf-8")


def test_write_tables_empty_returns_no_paths(tmp_path):
    assert write_tables({}, str(tmp_path / "out")) == []
    assert not (tmp_path / "out").exists()
