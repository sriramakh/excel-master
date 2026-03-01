"""LLM-powered template selector: picks template, charts, KPIs, filters."""
from __future__ import annotations
import json
from typing import Any

from ..models import (
    DatasetProfile, DashboardConfig, DashboardTemplate, ColorTheme,
    KPIConfig, ChartConfig, FilterConfig, AggFunc, ChartType, NumberFormat,
)
from ..data.data_engine import profile_to_prompt_text
from .llm_client import LLMClient
from .themes import TEMPLATE_DEFAULT_THEME


SYSTEM_PROMPT = """You are an expert Excel dashboard architect and data visualization specialist.

Given a dataset profile, you will design the optimal dashboard configuration.
Your output must be a valid JSON object matching this exact schema:

{
  "template": "<one of: executive_summary | hr_analytics | dark_operational | financial | supply_chain | marketing | minimal_clean>",
  "title": "<concise dashboard title>",
  "subtitle": "<one-line description>",
  "theme": "<one of: corporate_blue | hr_purple | dark_mode | supply_green | finance_green | marketing_orange | slate_minimal | executive_navy>",
  "kpis": [
    {
      "label": "<short KPI label>",
      "column": "<exact column name from dataset>",
      "aggregation": "<sum | avg | count | max | min | median | distinct_count>",
      "format": "<number | currency | percentage | decimal | integer>",
      "prefix": "<optional prefix like $ or #>",
      "suffix": "<optional suffix like % or K>",
      "icon": "<optional emoji icon>",
      "trend_column": "<optional date/period column for trend>",
    },
    ... (3-6 KPIs total)
  ],
  "charts": [
    {
      "type": "<bar | bar_horizontal | line | pie | doughnut | area | scatter | combo>",
      "title": "<chart title>",
      "x_column": "<exact column name>",
      "y_columns": ["<exact column name>"],
      "position": "<top-left | top-center | top-right | mid-left | mid-center | mid-right | bot-left | bot-center | bot-right>",
      "aggregation": "<sum | avg | count>",
      "top_n": <0 for all, or 5-15 for top N>,
      "show_data_labels": true
    },
    ... (3-6 charts total)
  ],
  "primary_dimension": "<main grouping column name>",
  "time_column": "<date or period column if exists, else ''>",
  "filters": [
    {"column": "<column name>", "filter_type": "<dropdown | date_range | checkbox>"},
    ...
  ],
  "table_columns": ["<col1>", "<col2>", "<col3>", "<col4>", "<col5>"],
  "insights": [
    "<Key business insight 1>",
    "<Key business insight 2>",
    "<Key business insight 3>"
  ],
  "show_raw_data": true
}

TEMPLATE SELECTION RULES:
- executive_summary: Board/C-suite KPIs, strategic metrics, mixed chart types, corporate theme
- hr_analytics: People data, turnover, headcount, satisfaction — use hr_purple theme
- dark_operational: Dense operational data, working hours, costs — use dark_mode theme
- financial: P&L, budget variance, cash flow, revenue — use finance_green theme
- supply_chain: Logistics, shipments, inventory, carriers — use supply_green theme
- marketing: Campaigns, leads, ROI, web analytics — use marketing_orange theme
- minimal_clean: Clean general-purpose, survey, research, small datasets — use slate_minimal

CHART SELECTION RULES:
- Use 'bar' for comparing categories (sales by region, revenue by product)
- Use 'bar_horizontal' when category names are long (department names, carrier names)
- Use 'line' for time series trends (monthly revenue, quarterly KPIs)
- Use 'pie' for composition with 5 or fewer categories (market share, top 5 segments)
- Use 'doughnut' for ratio/composition with a center metric (satisfaction breakdown)
- Use 'area' for cumulative trends (cumulative revenue, stacked growth)
- Use 'scatter' for correlation analysis (two numeric variables)
- Always choose columns that EXIST in the dataset

KPI RULES:
- Choose 3-6 KPIs that represent the most important business metrics
- For financial data: total revenue, total cost, margin, budget variance
- For HR data: headcount, turnover rate, avg salary, satisfaction score
- For supply chain: shipment count, on-time rate, total freight cost, avg transit
- For marketing: total spend, total leads, avg ROI, conversion rate
- Always verify column names exist in the dataset

Return ONLY valid JSON. No explanations, no markdown, no code blocks."""


def _build_user_prompt(profile: DatasetProfile) -> str:
    return f"""Design an Excel dashboard for this dataset:

{profile_to_prompt_text(profile)}

Requirements:
1. Title should be business-meaningful (e.g., "Q4 Sales Performance Dashboard")
2. Select 3-6 KPIs from the most impactful numeric columns
3. Select 3-6 diverse chart types covering different analytical angles
4. Choose the chart positions to create a logical visual flow
5. All column names in your response MUST exactly match column names listed above
6. If the dataset has a time/date column, always include a line chart for trend analysis
7. For the filters, pick 2-3 categorical columns with manageable cardinality

Return the JSON dashboard configuration."""


def _safe_parse_config(raw: dict[str, Any], profile: DatasetProfile) -> DashboardConfig:
    """Parse LLM JSON into a validated DashboardConfig with fallbacks."""
    col_names = set(profile.column_names)

    # Parse template
    try:
        template = DashboardTemplate(raw.get("template", "executive_summary"))
    except ValueError:
        template = DashboardTemplate.EXECUTIVE_SUMMARY

    # Parse theme
    try:
        theme = ColorTheme(raw.get("theme", TEMPLATE_DEFAULT_THEME.get(template, ColorTheme.CORPORATE_BLUE)))
    except ValueError:
        theme = TEMPLATE_DEFAULT_THEME.get(template, ColorTheme.CORPORATE_BLUE)

    # Parse KPIs (validate column names)
    kpis = []
    for k in raw.get("kpis", []):
        col = k.get("column", "")
        if col not in col_names:
            # Try fuzzy match
            matched = next((c for c in profile.numeric_columns if col.lower() in c.lower()), None)
            if matched:
                col = matched
            else:
                continue
        try:
            agg = AggFunc(k.get("aggregation", "sum"))
        except ValueError:
            agg = AggFunc.SUM
        try:
            fmt = NumberFormat(k.get("format", "number"))
        except ValueError:
            fmt = NumberFormat.NUMBER
        kpis.append(KPIConfig(
            label=k.get("label", col),
            column=col,
            aggregation=agg,
            format=fmt,
            prefix=k.get("prefix", ""),
            suffix=k.get("suffix", ""),
            icon=k.get("icon", ""),
            trend_column=k.get("trend_column", ""),
        ))

    # Fallback KPIs
    if not kpis and profile.numeric_columns:
        for nc in profile.numeric_columns[:4]:
            kpis.append(KPIConfig(label=nc.replace("_", " ").title(),
                                   column=nc, aggregation=AggFunc.SUM,
                                   format=NumberFormat.NUMBER))

    # Parse charts (validate column names)
    charts = []
    for c in raw.get("charts", []):
        x = c.get("x_column", "")
        y_cols = [yc for yc in c.get("y_columns", []) if yc in col_names]
        if x not in col_names:
            # Try fuzzy match on categorical
            x = next((c2 for c2 in profile.categorical_columns if x.lower() in c2.lower()),
                      profile.categorical_columns[0] if profile.categorical_columns else "")
        if not x or not y_cols:
            continue
        try:
            chart_type = ChartType(c.get("type", "bar"))
        except ValueError:
            chart_type = ChartType.BAR
        try:
            agg = AggFunc(c.get("aggregation", "sum"))
        except ValueError:
            agg = AggFunc.SUM
        charts.append(ChartConfig(
            type=chart_type,
            title=c.get("title", f"{x} Analysis"),
            x_column=x,
            y_columns=y_cols,
            position=c.get("position", "mid-center"),
            aggregation=agg,
            top_n=c.get("top_n", 0),
            show_data_labels=c.get("show_data_labels", True),
        ))

    # Fallback charts
    if not charts:
        if profile.categorical_columns and profile.numeric_columns:
            charts.append(ChartConfig(
                type=ChartType.BAR,
                title=f"{profile.categorical_columns[0]} Analysis",
                x_column=profile.categorical_columns[0],
                y_columns=[profile.numeric_columns[0]],
                position="mid-center",
                aggregation=AggFunc.SUM,
            ))
        if profile.date_columns and profile.numeric_columns:
            charts.append(ChartConfig(
                type=ChartType.LINE,
                title="Trend Over Time",
                x_column=profile.date_columns[0],
                y_columns=[profile.numeric_columns[0]],
                position="bot-center",
                aggregation=AggFunc.SUM,
            ))

    # Parse filters
    filters = []
    for f in raw.get("filters", []):
        col = f.get("column", "")
        if col in col_names:
            try:
                ftype = f.get("filter_type", "dropdown")
            except Exception:
                ftype = "dropdown"
            filters.append(FilterConfig(column=col, filter_type=ftype))

    # Table columns
    table_cols = [c for c in raw.get("table_columns", []) if c in col_names]
    if not table_cols:
        table_cols = list(col_names)[:6]

    return DashboardConfig(
        template=template,
        title=raw.get("title", f"{profile.name} Dashboard"),
        subtitle=raw.get("subtitle", profile.description),
        theme=theme,
        kpis=kpis,
        charts=charts,
        primary_dimension=raw.get("primary_dimension", ""),
        time_column=raw.get("time_column", ""),
        filters=filters,
        table_columns=table_cols,
        insights=raw.get("insights", []),
        show_raw_data=raw.get("show_raw_data", True),
    )


class TemplateSelector:
    """Uses LLM to select and configure the optimal dashboard."""

    def __init__(self):
        self.llm = LLMClient()

    def select(self, profile: DatasetProfile) -> DashboardConfig:
        """Call LLM to get dashboard configuration for the given dataset."""
        print(f"  Asking LLM to configure dashboard for: {profile.name}...")
        raw = self.llm.generate_json(
            system_prompt=SYSTEM_PROMPT,
            user_prompt=_build_user_prompt(profile),
        )
        config = _safe_parse_config(raw, profile)
        print(f"  → Template: {config.template.value}, Theme: {config.theme.value}")
        print(f"  → KPIs: {len(config.kpis)}, Charts: {len(config.charts)}")
        return config

    def select_with_override(self, profile: DatasetProfile,
                              template: str | None = None,
                              theme: str | None = None) -> DashboardConfig:
        """Select config but allow overriding template and theme."""
        config = self.select(profile)
        if template:
            try:
                config.template = DashboardTemplate(template)
            except ValueError:
                pass
        if theme:
            try:
                config.theme = ColorTheme(theme)
            except ValueError:
                pass
        return config
