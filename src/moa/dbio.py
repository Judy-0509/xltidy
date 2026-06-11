from __future__ import annotations

from pathlib import Path

import pandas as pd


def write_tables(tables: dict[str, pd.DataFrame], out_dir: str, fmt: str = "csv",
                 single_table: bool = True) -> list[str]:
    """기본은 모든 표를 table 컬럼과 함께 db.csv|.parquet 하나로 쓴다."""
    if not tables:
        return []

    d = Path(out_dir)
    d.mkdir(parents=True, exist_ok=True)

    if single_table:
        frames: list[pd.DataFrame] = []
        for name, frame in tables.items():
            out = frame.copy()
            out.insert(0, "table", name)
            frames.append(out)
        db = pd.concat(frames, ignore_index=True, sort=False)
        if fmt == "parquet":
            p = d / "db.parquet"
            try:
                db.to_parquet(p, index=False)  # pyarrow extras
            except ImportError as e:
                raise RuntimeError(
                    "parquet 출력에는 pyarrow가 필요합니다: pip install moa[parquet] "
                    "또는 pip install pyarrow") from e
        else:
            p = d / "db.csv"
            db.to_csv(p, index=False, encoding="utf-8-sig")
        return [str(p)]

    written: list[str] = []
    for name, frame in tables.items():
        if fmt == "parquet":
            p = d / f"{name}.parquet"
            try:
                frame.to_parquet(p, index=False)  # pyarrow extras
            except ImportError as e:
                raise RuntimeError(
                    "parquet 출력에는 pyarrow가 필요합니다: pip install moa[parquet] "
                    "또는 pip install pyarrow") from e
        else:
            p = d / f"{name}.csv"
            frame.to_csv(p, index=False, encoding="utf-8-sig")
        written.append(str(p))
    return written
