from __future__ import annotations

from pathlib import Path

import pandas as pd


def write_tables(tables: dict[str, pd.DataFrame], out_dir: str, fmt: str = "csv") -> list[str]:
    """워크북 1개 = 폴더 1개. 표마다 <name>.csv|.parquet."""
    d = Path(out_dir)
    d.mkdir(parents=True, exist_ok=True)
    written: list[str] = []
    for name, frame in tables.items():
        if fmt == "parquet":
            p = d / f"{name}.parquet"
            frame.to_parquet(p, index=False)  # pyarrow extras
        else:
            p = d / f"{name}.csv"
            frame.to_csv(p, index=False, encoding="utf-8-sig")
        written.append(str(p))
    return written
