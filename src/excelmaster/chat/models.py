"""Pydantic models for chat-driven dashboard state."""
from __future__ import annotations

import copy
from enum import Enum
from typing import Any, Literal, Union

from pydantic import BaseModel, Field

from excelmaster.models import (
    AggFunc,
    ChartConfig,
    ChartType,
    ColorTheme,
    KPIConfig,
    NumberFormat,
)


# ─── Object Type Enum ────────────────────────────────────────────────────────

class ObjectType(str, Enum):
    TITLE = "title"
    FILTER_PANEL = "filter_panel"
    KPI_ROW = "kpi_row"
    SECTION_HEADER = "section_header"
    CHART = "chart"
    TABLE = "table"
    PIVOT = "pivot"
    TEXT = "text"


# ─── Placed Payloads ─────────────────────────────────────────────────────────

class PlacedTitle(BaseModel):
    text: str
    subtitle: str = ""


class PlacedFilterPanel(BaseModel):
    filter_columns: list[str] = Field(default_factory=list)


class PlacedKPIRow(BaseModel):
    kpis: list[KPIConfig] = Field(default_factory=list)


class PlacedSectionHeader(BaseModel):
    text: str
    color: str = ""  # hex override; empty = theme default


class PlacedChart(BaseModel):
    chart: ChartConfig
    width: Literal["full", "half"] = "half"
    side: Literal["left", "right"] = "left"


class PlacedTable(BaseModel):
    columns: list[str] = Field(default_factory=list)
    max_rows: int = 15
    show_conditional: bool = True


class PlacedPivot(BaseModel):
    index_col: str
    value_col: str
    columns_col: str = ""
    agg: AggFunc = AggFunc.SUM


class PlacedText(BaseModel):
    content: str
    style: Literal["body", "heading", "insight", "footnote"] = "body"


PayloadUnion = Union[
    PlacedTitle,
    PlacedFilterPanel,
    PlacedKPIRow,
    PlacedSectionHeader,
    PlacedChart,
    PlacedTable,
    PlacedPivot,
    PlacedText,
]


# ─── Cell-Level Operation Models ────────────────────────────────────────────

class CellWrite(BaseModel):
    cell: str              # "A1" or "B3"
    value: Any = None      # string, number, formula (=...), or None
    format: dict = Field(default_factory=dict)  # {bold: True, bg_color: "#FF0000", ...}


class CellFormatOp(BaseModel):
    range: str             # "A1:F20"
    format: dict = Field(default_factory=dict)


class ConditionalFormatOp(BaseModel):
    range: str
    rule_type: str = ""    # "3_color_scale", "data_bar", "icon_set", "cell_is"
    params: dict = Field(default_factory=dict)


class DataValidationOp(BaseModel):
    range: str
    validation_type: str = "list"  # "list", "whole", "decimal", "custom"
    params: dict = Field(default_factory=dict)


class MergeOp(BaseModel):
    range: str
    value: str = ""
    format: dict = Field(default_factory=dict)


class HyperlinkOp(BaseModel):
    cell: str
    url: str
    display_text: str = ""


class CommentOp(BaseModel):
    cell: str
    text: str
    author: str = "Excel Master"


class ImageOp(BaseModel):
    cell: str
    image_path: str
    x_scale: float = 1.0
    y_scale: float = 1.0


# ─── Placed Object ───────────────────────────────────────────────────────────

class PlacedObject(BaseModel):
    id: str                              # e.g. "chart_0", "kpi_row_0"
    type: ObjectType
    anchor_row: int = 0
    height_rows: int = 1
    payload: PayloadUnion

    @property
    def end_row(self) -> int:
        return self.anchor_row + self.height_rows


# ─── Sheet Layout ────────────────────────────────────────────────────────────

class SheetLayout(BaseModel):
    name: str = "Dashboard"
    objects: list[PlacedObject] = Field(default_factory=list)
    freeze_row: int = 0
    freeze_col: int = 0
    zoom: int = 100
    tab_color: str = ""
    hidden: bool = False
    # Cell-level operations (rendered after placed objects)
    cell_writes: list[CellWrite] = Field(default_factory=list)
    cell_formats: list[CellFormatOp] = Field(default_factory=list)
    conditional_formats: list[ConditionalFormatOp] = Field(default_factory=list)
    data_validations: list[DataValidationOp] = Field(default_factory=list)
    merges: list[MergeOp] = Field(default_factory=list)
    hyperlinks: list[HyperlinkOp] = Field(default_factory=list)
    comments: list[CommentOp] = Field(default_factory=list)
    images: list[ImageOp] = Field(default_factory=list)
    row_heights: dict[int, float] = Field(default_factory=dict)
    col_widths: dict[int, float] = Field(default_factory=dict)
    hidden_rows: list[int] = Field(default_factory=list)
    hidden_cols: list[int] = Field(default_factory=list)

    def next_free_row(self) -> int:
        if not self.objects:
            return 0
        return max(o.end_row for o in self.objects)

    def sorted_objects(self) -> list[PlacedObject]:
        return sorted(self.objects, key=lambda o: (o.anchor_row, 0 if o.type != ObjectType.CHART else (0 if getattr(o.payload, "side", "left") == "left" else 1)))

    def find_object(self, obj_id: str) -> PlacedObject | None:
        for o in self.objects:
            if o.id == obj_id:
                return o
        return None

    def remove_object(self, obj_id: str) -> PlacedObject | None:
        for i, o in enumerate(self.objects):
            if o.id == obj_id:
                return self.objects.pop(i)
        return None


# ─── Workbook State ──────────────────────────────────────────────────────────

class WorkbookState(BaseModel):
    title: str = "Dashboard"
    theme_key: str = "corporate_blue"
    sheets: list[SheetLayout] = Field(default_factory=lambda: [SheetLayout(name="Dashboard")])
    version: int = 0

    def dashboard_sheet(self) -> SheetLayout:
        for s in self.sheets:
            if s.name == "Dashboard":
                return s
        return self.sheets[0]

    def get_sheet(self, name: str) -> SheetLayout | None:
        for s in self.sheets:
            if s.name == name:
                return s
        return None

    def add_sheet(self, name: str) -> SheetLayout:
        existing = self.get_sheet(name)
        if existing:
            return existing
        sheet = SheetLayout(name=name)
        self.sheets.append(sheet)
        return sheet

    def snapshot(self) -> WorkbookState:
        return self.model_copy(deep=True)


# ─── Action Types ────────────────────────────────────────────────────────────

class ActionType(str, Enum):
    ADD_CHART = "add_chart"
    MODIFY_CHART = "modify_chart"
    ADD_TABLE = "add_table"
    MODIFY_TABLE = "modify_table"
    ADD_KPI_ROW = "add_kpi_row"
    MODIFY_KPI = "modify_kpi"
    ADD_PIVOT = "add_pivot"
    ADD_SECTION_HEADER = "add_section_header"
    ADD_TEXT = "add_text"
    REMOVE = "remove"
    MOVE = "move"
    ADD_SHEET = "add_sheet"
    CHANGE_THEME = "change_theme"
    CHANGE_TITLE = "change_title"
    AUTO_DASHBOARD = "auto_dashboard"


class ChatAction(BaseModel):
    action: ActionType
    target_sheet: str = "Dashboard"
    target_id: str = ""
    params: dict[str, Any] = Field(default_factory=dict)
