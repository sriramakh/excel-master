# Excel Master — Agent Skills

## Identity
- **name**: excel-master
- **version**: 0.1.0
- **description**: AI-powered Excel dashboard engine with agentic capabilities
- **runtime**: Python 3.11+ (venv at `.venv/`)

## Capabilities
- Build Excel dashboards from CSV/XLSX data via natural language
- Add/modify/remove charts (bar, line, pie, area, scatter, doughnut, horizontal bar)
- Add KPI metric tile rows with configurable aggregation and formatting
- Add data tables with conditional formatting (3-color scale, data bars)
- Add pivot tables with flexible aggregation
- Write to individual cells (values, formulas, formatting)
- Format cell ranges (bold, colors, borders, number format)
- Apply conditional formatting (color scales, data bars, icon sets, cell rules)
- Add data validation (dropdown lists, whole number, decimal, custom)
- Merge cells with text and formatting
- Insert hyperlinks, comments, images
- Manage worksheets (create, rename, delete, reorder, hide/show, tab color)
- Resize/hide rows and columns
- Freeze panes and set zoom level
- Full undo/redo with 30-level history
- Object registry: every artifact tracked by ID for later reference/modification
- Multi-sheet auto-join for complex Excel files

## Invocation

### CLI
```bash
# Interactive agentic REPL
excelmaster agent <data_file> [-o output.xlsx]

# Legacy chat mode (backward-compatible)
excelmaster chat <data_file> [-o output.xlsx]
```

### Python API
```python
from excelmaster.agent import AgentSession

session = AgentSession("data.csv")
session.load()

# Natural language (uses LLM + tool calling)
result = session.execute_instruction("Add a bar chart of revenue by region")

# Direct tool call (no LLM, instant)
result = session.execute_tool("add_chart", {
    "type": "bar",
    "x_column": "region",
    "y_columns": ["revenue"],
    "title": "Revenue by Region"
})

# Auto-build full dashboard
result = session.auto_dashboard()

# Save
session.save("output/dashboard.xlsx")
```

### Direct Tool Call (no LLM)
```python
# Cell-level operations
session.execute_tool("write_cells", {
    "writes": [
        {"cell": "A1", "value": "Summary", "bold": True, "font_size": 14},
        {"cell": "B1", "value": "=SUM(B2:B100)", "num_format": "#,##0.00"}
    ]
})

# Conditional formatting
session.execute_tool("add_excel_feature", {
    "feature": "conditional_format",
    "range": "C2:C50",
    "rule_type": "data_bar",
    "bar_color": "#4472C4"
})

# Sheet management
session.execute_tool("sheet_operation", {
    "operation": "create",
    "sheet": "Regional Analysis"
})
```

## Tools

### 1. add_chart
Add a chart to the dashboard.
- **Required**: `type` (bar|line|pie|doughnut|area|scatter|bar_horizontal|combo), `x_column`, `y_columns`
- **Optional**: `title`, `aggregation`, `width` (full|half), `side` (left|right), `top_n`, `show_data_labels`, `sheet`, `position`

### 2. modify_object
Modify any existing object by its ID.
- **Required**: `object_id`, `changes` (dict of fields to update)
- **Optional**: `sheet`

### 3. remove_object
Remove an object from the dashboard by ID.
- **Required**: `object_id`
- **Optional**: `sheet`

### 4. add_kpi_row
Add a row of KPI metric tiles.
- **Required**: `kpis` (list of {label, column, aggregation?, format?, prefix?, suffix?})
- **Optional**: `sheet`, `position`

### 5. add_table
Add a data table or pivot table.
- **Optional**: `table_type` (data|pivot), `columns`, `max_rows`, `show_conditional`, `index_col`, `value_col`, `columns_col`, `agg`, `sheet`, `position`

### 6. add_content
Add title bar, section header, or text block.
- **Required**: `content_type` (title|section_header|text), `text`
- **Optional**: `subtitle`, `style`, `color`, `sheet`, `position`

### 7. write_cells
Write values/formulas/formatting to individual cells.
- **Required**: `writes` (list of {cell, value?, bold?, italic?, font_size?, font_color?, bg_color?, num_format?, align?, border?})
- **Optional**: `sheet`

### 8. format_range
Apply formatting to a cell range.
- **Required**: `range`
- **Optional**: `bold`, `italic`, `font_size`, `font_color`, `bg_color`, `num_format`, `align`, `valign`, `border`, `text_wrap`, `sheet`

### 9. sheet_operation
Create, rename, delete, reorder sheets. Set tab color, hide/show.
- **Required**: `operation` (create|rename|delete|reorder|set_tab_color|hide|show)
- **Optional**: `sheet`, `new_name`, `position`, `tab_color`

### 10. row_col_operation
Resize or hide/show rows and columns.
- **Required**: `target` (row|column), `operation` (resize|hide|show), `index`
- **Optional**: `end_index`, `size`, `sheet`

### 11. add_excel_feature
Add conditional formatting, data validation, freeze panes, zoom, merge cells, hyperlinks, comments, images.
- **Required**: `feature` (conditional_format|data_validation|freeze_panes|zoom|merge|hyperlink|comment|image)
- **Optional**: varies by feature (see tool schema for details)

### 12. change_theme
Change the workbook color theme.
- **Required**: `theme` (corporate_blue|hr_purple|dark_mode|supply_green|finance_green|marketing_orange|slate_minimal|executive_navy)

### 13. query_workbook
Read-only inspection of the workbook state.
- **Required**: `query` (list_objects|object_details|data_summary|list_sheets|registry_snapshot)
- **Optional**: `object_id`, `sheet`

## State Management

### Undo/Redo
```python
session.undo()  # Reverts last action
session.redo()  # Re-applies reverted action
```

### State Inspection
```python
state = session.get_state()    # Full state dict
tools = session.get_tools()    # Tool schemas for discovery
```

## Themes
| Theme | Color | Best For |
|-------|-------|----------|
| corporate_blue | Blue | Executive, general |
| hr_purple | Purple | People/HR data |
| dark_mode | Dark | Dense operational |
| supply_green | Green | Supply chain |
| finance_green | Green | Financial reports |
| marketing_orange | Orange | Marketing/campaigns |
| slate_minimal | Gray | Clean, minimal |
| executive_navy | Navy | C-Suite |
