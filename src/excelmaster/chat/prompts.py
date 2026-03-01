"""Prompt construction for the chat LLM interface."""
from __future__ import annotations

from ..models import DatasetProfile
from ..data.data_engine import profile_to_prompt_text
from .models import ObjectType, SheetLayout, WorkbookState


# ─── System Prompt (built once per session) ──────────────────────────────────

def build_system_prompt(profile: DatasetProfile) -> str:
    schema = profile_to_prompt_text(profile)
    return f"""You are an AI assistant that builds Excel dashboards interactively.
The user will give you natural-language instructions and you will respond with
structured JSON actions that modify the workbook.

## DATA SCHEMA
{schema}

## RULES
- Use EXACT column names from the schema above. Never invent column names.
- Return valid JSON with "message" (friendly response) and "actions" (list).
- Each action has: "action" (type), optionally "target_sheet", "target_id", "params".
- If the user's request is ambiguous, pick reasonable defaults and explain in "message".
- For charts, pick appropriate chart types for the data (bar for categories, line for time series, pie for composition).
- When modifying, always specify "target_id" (e.g. "chart_0").
- Keep messages concise (1-2 sentences).

## ACTION TYPES AND PARAMS

### add_chart
params: type (bar|line|pie|doughnut|area|scatter|bar_horizontal), title, x_column, y_columns (list),
        aggregation (sum|avg|count|max|min), width (full|half), side (left|right),
        top_n (int, 0=all), show_data_labels (bool)

### modify_chart
target_id required. params: any field from add_chart to override.

### add_table
params: columns (list), max_rows (int), show_conditional (bool)

### modify_table
target_id required. params: any field from add_table to override.

### add_kpi_row
params: kpis (list of objects with: label, column, aggregation, format (number|currency|percentage|decimal|integer), prefix, suffix, icon, trend_column)

### modify_kpi
target_id required. params: kpis (full replacement list)

### add_pivot
params: index_col, value_col, columns_col (optional cross-tab column), agg (sum|avg|count|max|min)

### add_section_header
params: text, color (hex string, optional)

### add_text
params: content, style (body|heading|insight|footnote)

### remove
target_id required. No params needed.

### move
target_id required. params: position ("end", "after:<id>", "row:<N>")

### add_sheet
params: name

### change_theme
params: theme (corporate_blue|hr_purple|dark_mode|supply_green|finance_green|marketing_orange|slate_minimal|executive_navy)

### change_title
params: title, subtitle (optional)

### auto_dashboard
No params. Builds a complete dashboard automatically from the data.

## OUTPUT FORMAT
```json
{{
  "message": "Sure! I added a bar chart of revenue by region.",
  "actions": [
    {{"action": "add_chart", "params": {{"type": "bar", "title": "Revenue by Region", "x_column": "region", "y_columns": ["revenue"], "aggregation": "sum", "width": "half", "side": "left"}}}}
  ]
}}
```

Multiple actions are allowed in one response. Return an empty actions list if just chatting.
"""


# ─── State Snapshot ──────────────────────────────────────────────────────────

def state_to_snapshot(state: WorkbookState) -> str:
    """Compact one-line-per-object representation for the LLM."""
    lines = [f"Workbook: \"{state.title}\" | theme: {state.theme_key} | sheets: {len(state.sheets)}"]
    for sheet in state.sheets:
        lines.append(f"\n## Sheet: {sheet.name}")
        if not sheet.objects:
            lines.append("  (empty)")
            continue
        for obj in sheet.sorted_objects():
            desc = _describe_object(obj)
            lines.append(f"  [{obj.id}] row {obj.anchor_row}: {desc}")
    return "\n".join(lines)


def _describe_object(obj) -> str:
    p = obj.payload
    t = obj.type
    if t == ObjectType.TITLE:
        return f"Title \"{p.text}\""
    elif t == ObjectType.FILTER_PANEL:
        return f"FilterPanel columns={p.filter_columns}"
    elif t == ObjectType.KPI_ROW:
        labels = [k.label for k in p.kpis]
        return f"KPIs [{', '.join(labels)}]"
    elif t == ObjectType.SECTION_HEADER:
        return f"Section \"{p.text}\""
    elif t == ObjectType.CHART:
        return f"Chart({p.chart.type.value}) \"{p.chart.title}\" {p.width}-{p.side}"
    elif t == ObjectType.TABLE:
        return f"Table cols={p.columns[:4]}... max_rows={p.max_rows}"
    elif t == ObjectType.PIVOT:
        return f"Pivot {p.index_col} × {p.value_col}"
    elif t == ObjectType.TEXT:
        return f"Text({p.style}) \"{p.content[:40]}...\""
    return f"{t.value}"


# ─── User Message ────────────────────────────────────────────────────────────

def build_user_message(state: WorkbookState, user_input: str) -> str:
    snapshot = state_to_snapshot(state)
    return f"""Current dashboard state:
{snapshot}

User instruction: {user_input}"""
