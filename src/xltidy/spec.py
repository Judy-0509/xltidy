from __future__ import annotations

from pathlib import Path
from typing import Literal

import yaml
from pydantic import BaseModel, Field

from .coords import parse_range
from .models import CellGrid


class SheetMatch(BaseModel):
    by: Literal["name", "index"] = "name"
    value: str | int


class HeaderSpec(BaseModel):
    orientation: Literal["top"] = "top"
    levels: int = 1  # 참고용 메타. 다중헤더는 agent가 column_semantics로 평탄화
    rows: list[int] = []


class IndexColumn(BaseModel):
    col: str
    name: str
    type: Literal["str", "int", "float", "date_period"] = "str"


class ValueBlock(BaseModel):
    cols: list[str]  # [start, end] 포함


class ColumnSemantic(BaseModel):
    source: str
    name: str
    type: Literal["number", "str", "date_period"] = "number"
    source_text: str | None = None  # 드리프트 감지용 원본 헤더 텍스트


class UnpivotSpec(BaseModel):
    var_name: str = "metric"
    value_name: str = "value"


class PeriodSource(BaseModel):
    from_: Literal["filename", "cell"] = Field("filename", alias="from")
    pattern: str | None = None
    cell: str | None = None
    model_config = {"populate_by_name": True}


class PeriodSpec(BaseModel):
    source: PeriodSource
    name: str = "period"


class TotalCheck(BaseModel):
    kind: Literal["row_subtotal"] = "row_subtotal"
    label: str
    over: str


class Region(BaseModel):
    start: str
    end: str


class TableSpec(BaseModel):
    name: str
    kind: Literal["table", "pivot"] = "table"
    period: PeriodSpec
    # kind=table
    region: Region | None = None
    header: HeaderSpec | None = None
    index_columns: list[IndexColumn] = []
    value_block: ValueBlock | None = None
    unpivot: UnpivotSpec = UnpivotSpec()
    column_semantics: list[ColumnSemantic] = []
    totals: list[TotalCheck] = []
    # kind=pivot
    pivot_name: str | None = None


class SheetSpec(BaseModel):
    sheet_match: SheetMatch
    tables: list[TableSpec]


class TemplateSpec(BaseModel):
    template_id: str
    version: int = 1
    sheets: list[SheetSpec]

    @classmethod
    def from_yaml(cls, path: str | Path) -> "TemplateSpec":
        return cls.model_validate(yaml.safe_load(Path(path).read_text(encoding="utf-8")))

    def to_yaml(self, path: str | Path) -> None:
        data = self.model_dump(by_alias=True, exclude_none=True)
        Path(path).write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding="utf-8")

    def validate_against(self, grids: dict[str, CellGrid]) -> list[str]:
        issues: list[str] = []
        for sheet in self.sheets:
            grid = grids.get(str(sheet.sheet_match.value))
            for t in sheet.tables:
                if t.kind == "table":
                    if not (t.region and t.header and t.value_block):
                        issues.append(f"table {t.name}: kind=table requires region/header/value_block")
                        continue
                    if grid is not None:
                        r1, c1, r2, c2 = parse_range(f"{t.region.start}:{t.region.end}")
                        if r2 > grid.n_rows or c2 > grid.n_cols:
                            issues.append(f"table {t.name}: region {t.region.start}:{t.region.end} out of bounds")
        return issues


def sample_spec_dict() -> dict:
    return {
        "template_id": "sample-survey-monthly",
        "version": 1,
        "sheets": [{
            "sheet_match": {"by": "name", "value": "데이터"},
            "tables": [{
                "name": "by_industry", "kind": "table",
                "region": {"start": "B5", "end": "D8"},
                "header": {"orientation": "top", "levels": 1, "rows": [5]},
                "index_columns": [{"col": "B", "name": "industry", "type": "str"}],
                "value_block": {"cols": ["C", "D"]},
                "unpivot": {"var_name": "month", "value_name": "value"},
                "column_semantics": [
                    {"source": "C5", "name": "2024-01", "type": "number", "source_text": "2024년 1월"},
                    {"source": "D5", "name": "2024-02", "type": "number", "source_text": "2024년 2월"},
                ],
                "period": {"source": {"from": "filename", "pattern": r"(\d{4})[._-]?Q?([1-4])"}, "name": "period"},
                "totals": [{"kind": "row_subtotal", "label": "합계", "over": "industry"}],
            }],
        }],
    }
