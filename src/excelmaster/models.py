"""Pydantic models for Excel Master."""
from __future__ import annotations
from enum import Enum
from typing import Any, Literal
from pydantic import BaseModel, Field


# ─── Enums ────────────────────────────────────────────────────────────────────

class ChartType(str, Enum):
    BAR = "bar"
    BAR_HORIZONTAL = "bar_horizontal"
    LINE = "line"
    PIE = "pie"
    DOUGHNUT = "doughnut"
    AREA = "area"
    SCATTER = "scatter"
    COMBO = "combo"


class AggFunc(str, Enum):
    SUM = "sum"
    AVG = "avg"
    COUNT = "count"
    MAX = "max"
    MIN = "min"
    MEDIAN = "median"
    DISTINCT = "distinct_count"


class NumberFormat(str, Enum):
    NUMBER = "number"
    CURRENCY = "currency"
    PERCENTAGE = "percentage"
    DECIMAL = "decimal"
    INTEGER = "integer"


class DashboardTemplate(str, Enum):
    EXECUTIVE_SUMMARY = "executive_summary"
    HR_ANALYTICS = "hr_analytics"
    DARK_OPERATIONAL = "dark_operational"
    FINANCIAL = "financial"
    SUPPLY_CHAIN = "supply_chain"
    MARKETING = "marketing"
    MINIMAL_CLEAN = "minimal_clean"


class ColorTheme(str, Enum):
    CORPORATE_BLUE = "corporate_blue"
    HR_PURPLE = "hr_purple"
    DARK_MODE = "dark_mode"
    SUPPLY_GREEN = "supply_green"
    FINANCE_GREEN = "finance_green"
    MARKETING_ORANGE = "marketing_orange"
    SLATE_MINIMAL = "slate_minimal"
    EXECUTIVE_NAVY = "executive_navy"


# ─── Data Profile ──────────────────────────────────────────────────────────────

class ColumnInfo(BaseModel):
    name: str
    dtype: str                    # "numeric" | "categorical" | "date" | "text" | "boolean"
    unique_values: int = 0
    null_pct: float = 0.0
    sample_values: list[Any] = Field(default_factory=list)
    min_val: Any = None
    max_val: Any = None


class DatasetProfile(BaseModel):
    """Profile of a dataset used to inform LLM decisions."""
    name: str
    file_path: str
    sheet_name: str = "Data"
    rows: int
    columns: list[ColumnInfo]
    industry: str = ""
    description: str = ""
    date_columns: list[str] = Field(default_factory=list)
    numeric_columns: list[str] = Field(default_factory=list)
    categorical_columns: list[str] = Field(default_factory=list)

    @property
    def column_names(self) -> list[str]:
        return [c.name for c in self.columns]


# ─── Deep Analysis (LLM Output) ───────────────────────────────────────────────

class CorrelationInsight(BaseModel):
    col_a: str = ""
    col_b: str = ""
    coefficient: float = 0.0
    interpretation: str = ""

class OutlierInsight(BaseModel):
    column: str = ""
    count: int = 0
    pct: float = 0.0
    description: str = ""

class PerformerEntry(BaseModel):
    dimension_value: str = ""
    metric_value: float = 0.0
    metric_column: str = ""

class TrendInsight(BaseModel):
    column: str = ""
    direction: str = ""       # "up", "down", "flat"
    description: str = ""
    pct_change: float = 0.0

class DeepAnalysis(BaseModel):
    """Structured deep analysis produced by LLM from pre-computed stats."""
    executive_summary: str = ""
    key_findings: list[str] = Field(default_factory=list)
    data_quality_score: int = 0
    data_quality_notes: list[str] = Field(default_factory=list)
    distribution_insights: list[str] = Field(default_factory=list)
    correlation_insights: list[CorrelationInsight] = Field(default_factory=list)
    outlier_insights: list[OutlierInsight] = Field(default_factory=list)
    top_performers: list[PerformerEntry] = Field(default_factory=list)
    bottom_performers: list[PerformerEntry] = Field(default_factory=list)
    dimension_analysis: str = ""
    trend_insights: list[TrendInsight] = Field(default_factory=list)
    trend_summary: str = ""
    near_term_outlook: str = ""
    long_term_outlook: str = ""
    recommendations: list[str] = Field(default_factory=list)
    industry_context: str = ""


# ─── Dashboard Config (LLM Output) ────────────────────────────────────────────

class KPIConfig(BaseModel):
    label: str
    column: str
    aggregation: AggFunc = AggFunc.SUM
    format: NumberFormat = NumberFormat.NUMBER
    prefix: str = ""
    suffix: str = ""
    icon: str = ""                # emoji or icon name
    trend_column: str = ""        # column for sparkline/trend


class ChartConfig(BaseModel):
    type: ChartType
    title: str
    x_column: str
    y_columns: list[str]
    color: str = ""
    position: Literal["top-left", "top-center", "top-right",
                       "mid-left", "mid-center", "mid-right",
                       "bot-left", "bot-center", "bot-right"] = "mid-center"
    aggregation: AggFunc = AggFunc.SUM
    top_n: int = 0                # limit categories to top N
    show_data_labels: bool = True


class FilterConfig(BaseModel):
    column: str
    filter_type: Literal["dropdown", "date_range", "checkbox"] = "dropdown"


class DashboardConfig(BaseModel):
    """Complete dashboard configuration produced by LLM."""
    template: DashboardTemplate
    title: str
    subtitle: str = ""
    theme: ColorTheme = ColorTheme.CORPORATE_BLUE
    kpis: list[KPIConfig] = Field(default_factory=list)
    charts: list[ChartConfig] = Field(default_factory=list)
    primary_dimension: str = ""   # main grouping/slice column
    time_column: str = ""
    filters: list[FilterConfig] = Field(default_factory=list)
    table_columns: list[str] = Field(default_factory=list)
    insights: list[str] = Field(default_factory=list)
    show_raw_data: bool = True
    deep_analysis: DeepAnalysis | None = None


# ─── Build Result ──────────────────────────────────────────────────────────────

class BuildResult(BaseModel):
    success: bool
    output_path: str
    dataset: str
    template_used: str
    kpi_count: int
    chart_count: int
    error: str = ""
