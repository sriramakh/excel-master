"""OpenAI function/tool schemas for the agentic Excel Master."""
from __future__ import annotations

TOOLS: list[dict] = [
    # ── 1. add_chart ──────────────────────────────────────────────────────────
    {
        "type": "function",
        "function": {
            "name": "add_chart",
            "description": "Add a chart to the dashboard. Supports bar, line, pie, area, scatter, doughnut, bar_horizontal, and combo types.",
            "parameters": {
                "type": "object",
                "properties": {
                    "type": {
                        "type": "string",
                        "enum": ["bar", "line", "pie", "doughnut", "area", "scatter", "bar_horizontal", "combo"],
                        "description": "Chart type",
                    },
                    "title": {"type": "string", "description": "Chart title"},
                    "x_column": {"type": "string", "description": "Column for x-axis / categories"},
                    "y_columns": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Column(s) for y-axis values",
                    },
                    "aggregation": {
                        "type": "string",
                        "enum": ["sum", "avg", "count", "max", "min", "median"],
                        "description": "Aggregation function (default: sum)",
                    },
                    "width": {
                        "type": "string",
                        "enum": ["full", "half"],
                        "description": "Chart width (default: half)",
                    },
                    "side": {
                        "type": "string",
                        "enum": ["left", "right"],
                        "description": "Side for half-width charts (default: left)",
                    },
                    "top_n": {
                        "type": "integer",
                        "description": "Show only top N categories (0 = all)",
                    },
                    "show_data_labels": {
                        "type": "boolean",
                        "description": "Show data labels on chart (default: true)",
                    },
                    "sheet": {
                        "type": "string",
                        "description": "Target sheet name (default: Dashboard)",
                    },
                    "position": {
                        "type": "string",
                        "description": "Position: 'end', 'after:<id>', 'row:<N>'",
                    },
                },
                "required": ["type", "x_column", "y_columns"],
            },
        },
    },

    # ── 2. modify_object ──────────────────────────────────────────────────────
    {
        "type": "function",
        "function": {
            "name": "modify_object",
            "description": "Modify any existing object by its ID. Pass only the fields you want to change.",
            "parameters": {
                "type": "object",
                "properties": {
                    "object_id": {"type": "string", "description": "ID of the object to modify (e.g. chart_0, table_0)"},
                    "changes": {
                        "type": "object",
                        "description": "Fields to update. For charts: type, title, x_column, y_columns, aggregation, width, side. For tables: columns, max_rows, show_conditional. For KPI rows: kpis (full list). For text: content, style.",
                    },
                    "sheet": {"type": "string", "description": "Sheet where the object lives (default: Dashboard)"},
                },
                "required": ["object_id", "changes"],
            },
        },
    },

    # ── 3. remove_object ──────────────────────────────────────────────────────
    {
        "type": "function",
        "function": {
            "name": "remove_object",
            "description": "Remove an object from the dashboard by its ID.",
            "parameters": {
                "type": "object",
                "properties": {
                    "object_id": {"type": "string", "description": "ID of the object to remove"},
                    "sheet": {"type": "string", "description": "Sheet name (default: searches all sheets)"},
                },
                "required": ["object_id"],
            },
        },
    },

    # ── 4. add_kpi_row ────────────────────────────────────────────────────────
    {
        "type": "function",
        "function": {
            "name": "add_kpi_row",
            "description": "Add a row of KPI metric tiles to the dashboard.",
            "parameters": {
                "type": "object",
                "properties": {
                    "kpis": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "label": {"type": "string"},
                                "column": {"type": "string"},
                                "aggregation": {
                                    "type": "string",
                                    "enum": ["sum", "avg", "count", "max", "min", "median", "distinct_count"],
                                },
                                "format": {
                                    "type": "string",
                                    "enum": ["number", "currency", "percentage", "decimal", "integer"],
                                },
                                "prefix": {"type": "string"},
                                "suffix": {"type": "string"},
                                "icon": {"type": "string"},
                                "trend_column": {"type": "string"},
                            },
                            "required": ["label", "column"],
                        },
                        "description": "List of KPI definitions",
                    },
                    "sheet": {"type": "string"},
                    "position": {"type": "string"},
                },
                "required": ["kpis"],
            },
        },
    },

    # ── 5. add_table ──────────────────────────────────────────────────────────
    {
        "type": "function",
        "function": {
            "name": "add_table",
            "description": "Add a data table or pivot table to the dashboard.",
            "parameters": {
                "type": "object",
                "properties": {
                    "table_type": {
                        "type": "string",
                        "enum": ["data", "pivot"],
                        "description": "Table type (default: data)",
                    },
                    "columns": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Columns to show (for data tables)",
                    },
                    "max_rows": {"type": "integer", "description": "Max rows to display (default: 15)"},
                    "show_conditional": {"type": "boolean", "description": "Show conditional formatting (default: true)"},
                    "index_col": {"type": "string", "description": "Row grouping column (for pivot tables)"},
                    "value_col": {"type": "string", "description": "Value column (for pivot tables)"},
                    "columns_col": {"type": "string", "description": "Cross-tab column (for pivot tables)"},
                    "agg": {
                        "type": "string",
                        "enum": ["sum", "avg", "count", "max", "min", "median"],
                        "description": "Aggregation for pivot (default: sum)",
                    },
                    "sheet": {"type": "string"},
                    "position": {"type": "string"},
                },
                "required": [],
            },
        },
    },

    # ── 6. add_content ────────────────────────────────────────────────────────
    {
        "type": "function",
        "function": {
            "name": "add_content",
            "description": "Add a title bar, section header, or text block to the dashboard.",
            "parameters": {
                "type": "object",
                "properties": {
                    "content_type": {
                        "type": "string",
                        "enum": ["title", "section_header", "text"],
                        "description": "Type of content to add",
                    },
                    "text": {"type": "string", "description": "Content text"},
                    "subtitle": {"type": "string", "description": "Subtitle (for title type only)"},
                    "style": {
                        "type": "string",
                        "enum": ["body", "heading", "insight", "footnote"],
                        "description": "Text style (for text type, default: body)",
                    },
                    "color": {"type": "string", "description": "Hex color override (for section headers)"},
                    "sheet": {"type": "string"},
                    "position": {"type": "string"},
                },
                "required": ["content_type", "text"],
            },
        },
    },

    # ── 7. write_cells ────────────────────────────────────────────────────────
    {
        "type": "function",
        "function": {
            "name": "write_cells",
            "description": "Write values, formulas, and/or formatting to individual cells.",
            "parameters": {
                "type": "object",
                "properties": {
                    "writes": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "cell": {"type": "string", "description": "Cell address like 'A1', 'B3'"},
                                "value": {
                                    "description": "Value to write: string, number, or formula starting with '='",
                                },
                                "bold": {"type": "boolean"},
                                "italic": {"type": "boolean"},
                                "font_size": {"type": "number"},
                                "font_color": {"type": "string", "description": "Hex color like '#FF0000'"},
                                "bg_color": {"type": "string", "description": "Background hex color"},
                                "num_format": {"type": "string", "description": "Number format like '#,##0.00'"},
                                "align": {"type": "string", "enum": ["left", "center", "right"]},
                                "border": {"type": "integer", "description": "Border style (1=thin, 2=medium, 5=thick)"},
                            },
                            "required": ["cell"],
                        },
                        "description": "List of cell writes",
                    },
                    "sheet": {"type": "string", "description": "Target sheet (default: Dashboard)"},
                },
                "required": ["writes"],
            },
        },
    },

    # ── 8. format_range ───────────────────────────────────────────────────────
    {
        "type": "function",
        "function": {
            "name": "format_range",
            "description": "Apply formatting to a range of cells (bold, color, borders, number format).",
            "parameters": {
                "type": "object",
                "properties": {
                    "range": {"type": "string", "description": "Range like 'A1:F20'"},
                    "bold": {"type": "boolean"},
                    "italic": {"type": "boolean"},
                    "font_size": {"type": "number"},
                    "font_color": {"type": "string"},
                    "bg_color": {"type": "string"},
                    "num_format": {"type": "string"},
                    "align": {"type": "string", "enum": ["left", "center", "right"]},
                    "valign": {"type": "string", "enum": ["top", "vcenter", "bottom"]},
                    "border": {"type": "integer"},
                    "text_wrap": {"type": "boolean"},
                    "sheet": {"type": "string"},
                },
                "required": ["range"],
            },
        },
    },

    # ── 9. sheet_operation ────────────────────────────────────────────────────
    {
        "type": "function",
        "function": {
            "name": "sheet_operation",
            "description": "Create, rename, delete, reorder sheets. Set tab color, hide/show.",
            "parameters": {
                "type": "object",
                "properties": {
                    "operation": {
                        "type": "string",
                        "enum": ["create", "rename", "delete", "reorder", "set_tab_color", "hide", "show"],
                        "description": "Sheet operation type",
                    },
                    "sheet": {"type": "string", "description": "Sheet name to operate on"},
                    "new_name": {"type": "string", "description": "New name (for rename)"},
                    "position": {"type": "integer", "description": "New position index (for reorder)"},
                    "tab_color": {"type": "string", "description": "Hex tab color (for set_tab_color)"},
                },
                "required": ["operation"],
            },
        },
    },

    # ── 10. row_col_operation ─────────────────────────────────────────────────
    {
        "type": "function",
        "function": {
            "name": "row_col_operation",
            "description": "Insert, delete, resize, or hide/show rows and columns.",
            "parameters": {
                "type": "object",
                "properties": {
                    "target": {
                        "type": "string",
                        "enum": ["row", "column"],
                    },
                    "operation": {
                        "type": "string",
                        "enum": ["resize", "hide", "show"],
                    },
                    "index": {"type": "integer", "description": "Row or column index (0-based)"},
                    "end_index": {"type": "integer", "description": "End index for range operations (inclusive)"},
                    "size": {"type": "number", "description": "Height (rows) or width (columns) in points"},
                    "sheet": {"type": "string"},
                },
                "required": ["target", "operation", "index"],
            },
        },
    },

    # ── 11. add_excel_feature ─────────────────────────────────────────────────
    {
        "type": "function",
        "function": {
            "name": "add_excel_feature",
            "description": "Add Excel features: conditional formatting, data validation, freeze panes, zoom, merge cells, hyperlinks, comments, images.",
            "parameters": {
                "type": "object",
                "properties": {
                    "feature": {
                        "type": "string",
                        "enum": [
                            "conditional_format", "data_validation",
                            "freeze_panes", "zoom",
                            "merge", "hyperlink", "comment", "image",
                        ],
                        "description": "Feature type to add",
                    },
                    "range": {"type": "string", "description": "Cell or range (e.g. 'A1:F20')"},
                    "cell": {"type": "string", "description": "Single cell address (for hyperlink, comment, image)"},
                    # Conditional format
                    "rule_type": {
                        "type": "string",
                        "enum": ["3_color_scale", "2_color_scale", "data_bar", "icon_set", "cell_is"],
                        "description": "Conditional format rule type",
                    },
                    "criteria": {"type": "string", "description": "Criteria for cell_is rules (e.g. '>', 'between')"},
                    "value": {"description": "Threshold value for cell_is rules"},
                    "min_color": {"type": "string"},
                    "mid_color": {"type": "string"},
                    "max_color": {"type": "string"},
                    "bar_color": {"type": "string", "description": "Data bar color"},
                    # Data validation
                    "validate": {
                        "type": "string",
                        "enum": ["list", "whole", "decimal", "custom"],
                    },
                    "source": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List values for validation",
                    },
                    # Freeze panes
                    "freeze_row": {"type": "integer"},
                    "freeze_col": {"type": "integer"},
                    # Zoom
                    "zoom_level": {"type": "integer", "description": "Zoom percentage (10-400)"},
                    # Merge
                    "merge_value": {"type": "string", "description": "Text to write in merged cell"},
                    "format": {"type": "object", "description": "Format dict for merge"},
                    # Hyperlink
                    "url": {"type": "string"},
                    "display_text": {"type": "string"},
                    # Comment
                    "comment_text": {"type": "string"},
                    "author": {"type": "string"},
                    # Image
                    "image_path": {"type": "string"},
                    "x_scale": {"type": "number"},
                    "y_scale": {"type": "number"},
                    "sheet": {"type": "string"},
                },
                "required": ["feature"],
            },
        },
    },

    # ── 12. change_theme ──────────────────────────────────────────────────────
    {
        "type": "function",
        "function": {
            "name": "change_theme",
            "description": "Change the workbook color theme.",
            "parameters": {
                "type": "object",
                "properties": {
                    "theme": {
                        "type": "string",
                        "enum": [
                            "corporate_blue", "hr_purple", "dark_mode",
                            "supply_green", "finance_green",
                            "marketing_orange", "slate_minimal", "executive_navy",
                        ],
                        "description": "Color theme to apply",
                    },
                },
                "required": ["theme"],
            },
        },
    },

    # ── 13. query_workbook ────────────────────────────────────────────────────
    {
        "type": "function",
        "function": {
            "name": "query_workbook",
            "description": "Read-only: list objects, get object details, get data summary, inspect registry. Use this to discover IDs before modifying objects.",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "enum": [
                            "list_objects", "object_details",
                            "data_summary", "list_sheets",
                            "registry_snapshot",
                        ],
                        "description": "What to query",
                    },
                    "object_id": {"type": "string", "description": "Object ID (for object_details)"},
                    "sheet": {"type": "string", "description": "Filter by sheet name"},
                },
                "required": ["query"],
            },
        },
    },
]


def get_tool_schemas() -> list[dict]:
    """Return the full list of tool schemas for OpenAI API."""
    return TOOLS


def get_tool_names() -> list[str]:
    """Return the list of tool function names."""
    return [t["function"]["name"] for t in TOOLS]
