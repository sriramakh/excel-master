# Excel Master

**AI-powered Excel dashboard generation engine** — transform any dataset into professional, interactive Excel dashboards using LLM-driven design decisions.

Excel Master profiles your data, selects the optimal dashboard template, configures KPIs and charts, and renders a fully styled `.xlsx` file — all in one command. It also offers an **agentic mode** with OpenAI tool calling for programmatic dashboard construction, and a legacy **chat mode** for conversational building.

---

## Features

- **Agentic Mode (NEW)** — OpenAI tool-calling interface with 13 tools, object registry, and full Python API (`AgentSession`)
- **Automatic Dashboard Generation** — Feed in any `.xlsx` or `.csv` file and get a complete dashboard with KPIs, charts, tables, filters, and deep analysis
- **LLM-Powered Design** — Uses OpenAI or MiniMax to select the best template, theme, chart types, and KPI configurations for your data
- **13 Agent Tools** — `add_chart`, `modify_object`, `remove_object`, `add_kpi_row`, `add_table`, `add_content`, `write_cells`, `format_range`, `sheet_operation`, `row_col_operation`, `add_excel_feature`, `change_theme`, `query_workbook`
- **Cell-Level Control** — Write to individual cells, apply formatting, conditional formatting, data validation, merge cells, insert hyperlinks/comments/images
- **7 Professional Templates** — Executive Summary, HR Analytics, Dark Operational, Financial, Supply Chain, Marketing, and Minimal Clean
- **8 Color Themes** — Corporate Blue, HR Purple, Dark Mode, Supply Green, Finance Green, Marketing Orange, Slate Minimal, Executive Navy
- **Interactive Chat Mode** — Build and modify dashboards conversationally with undo/redo support (legacy)
- **Deep Analysis** — LLM-interpreted statistical analysis with correlations, outliers, trends, and actionable recommendations
- **Dynamic Formulas** — Dashboards use SUMIFS formulas linked to filter dropdowns, so they stay interactive in Excel
- **Object Registry** — Every artifact tracked by ID for later reference, modification, or removal
- **9 Synthetic Data Generators** — Generate realistic test datasets across industries (Finance, HR, Supply Chain, Marketing, etc.)
- **Multi-Sheet Intelligence** — Automatically discovers relationships between sheets and joins dimension tables into a unified fact table
- **Rule-Based Fallback** — Works without an LLM using heuristic template and chart selection (`--no-llm`)

---

## Quick Start

### Prerequisites

- Python 3.11+
- An OpenAI API key (required for agent mode; MiniMax supported for legacy chat/dashboard)

### Installation

```bash
# Clone the repository
git clone <repo-url>
cd "Excel Master"

# Create virtual environment
python -m venv .venv
source .venv/bin/activate

# Install in editable mode
pip install -e .
```

### Configuration

Copy the example environment file and fill in your API key:

```bash
cp .env.example .env
```

Then edit `.env` with your credentials. At minimum, set your OpenAI key:

```env
OPENAI_API_KEY=sk-your-openai-api-key-here
```

See [`.env.example`](.env.example) for all available settings (LLM provider, model, temperature, paths, etc.).

### Usage

```bash
# Generate a synthetic dataset
excelmaster generate-data finance

# Generate all 9 datasets
excelmaster generate-data all

# Build a dashboard from any Excel/CSV file
excelmaster build-dashboard data/finance_data.xlsx

# Full pipeline: generate data + build dashboard
excelmaster run finance

# Run all datasets through the full pipeline
excelmaster run all

# Profile a dataset without building a dashboard
excelmaster profile data/finance_data.xlsx

# Agentic mode (OpenAI tool calling)
excelmaster agent data/finance_data.xlsx

# Legacy chat mode
excelmaster chat data/finance_data.xlsx

# Build without LLM (rule-based fallback)
excelmaster build-dashboard data/finance_data.xlsx --no-llm

# Override template and theme
excelmaster build-dashboard data/finance_data.xlsx --template financial --theme finance_green

# List all available datasets, templates, and themes
excelmaster list
```

---

## Available Datasets

| Type | Industry | Description |
|------|----------|-------------|
| `extreme_load` | Multi-Industry | 5 datasets, 10K+ rows, 15+ columns each |
| `moderate` | E-Commerce | Orders data, 2,500 rows, 14 columns |
| `feature_rich` | Investment | Portfolio analytics, 22+ calculated columns |
| `sparse` | Research | Survey data with 35-60% null rates |
| `finance` | Finance | P&L, cash flow, accounts receivable |
| `supply_chain` | Logistics | Shipments, carriers, inventory |
| `executive` | C-Suite | KPI scorecard, OKRs, quarterly trends |
| `hr_admin` | HR | Employee master, payroll, recruitment |
| `marketing` | Marketing | Campaigns, web analytics, content |

## Available Templates

| Template | Best For | Default Theme |
|----------|----------|---------------|
| `executive_summary` | Board/C-Suite KPIs | `corporate_blue` |
| `hr_analytics` | People & Workforce | `hr_purple` |
| `dark_operational` | Dense Operational Data | `dark_mode` |
| `financial` | P&L, Budget, Cash Flow | `finance_green` |
| `supply_chain` | Logistics & Freight | `supply_green` |
| `marketing` | Campaigns & ROI | `marketing_orange` |
| `minimal_clean` | Research, Survey, General | `slate_minimal` |

---

## Agent Mode

The agent mode uses OpenAI tool calling for precise, multi-step dashboard construction:

```
You: auto
  Assistant: Auto-built dashboard: Finance Dashboard
  Dashboard v1: 4 charts, 5 KPIs, 1 tables, 0 cell ops (11 total objects)

You: add a pie chart of revenue by region
  [ok] add_chart: Added pie chart 'Pie Chart'
  Dashboard v2: 5 charts, 5 KPIs, 1 tables, 0 cell ops (12 total objects)

You: write "Summary" in bold to cell A1 and format B1:F1 with blue background
  [ok] write_cells: Wrote 1 cell(s)
  [ok] format_range: Formatted range B1:F1

You: undo
  Saved: output/finance_agent_dashboard.xlsx
```

**Commands:** `auto` | `undo` | `redo` | `show` | `save` | `save as <name>` | `quit`

### Python API

```python
from excelmaster.agent import AgentSession

session = AgentSession("data.csv")
session.load()

# Natural language (uses LLM + tool calling)
result = session.execute_instruction("Add a bar chart of revenue by region")

# Direct tool call (no LLM, instant & free)
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

### Direct Tool Calls (No LLM)

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

---

## Legacy Chat Mode

The original chat mode uses JSON-based action parsing for conversational dashboard building:

```
You: auto
  -> Auto dashboard built (4 charts, 5 KPIs, 1 table)

You: add a pie chart of revenue by region
  -> Added pie chart 'Revenue by Region' [chart_3]

You: undo
  -> Undone. Re-rendering...
```

**Commands:** `auto` | `start` | `undo` | `redo` | `show` | `reset` | `save as <name>` | `quit`

---

## Project Structure

```
Excel Master/
├── src/excelmaster/
│   ├── __init__.py              # Package init, version
│   ├── config.py                # Settings (env vars, LLM config, paths)
│   ├── models.py                # Core Pydantic models (KPI, Chart, Dashboard configs)
│   ├── cli/
│   │   └── app.py               # Typer CLI (generate-data, build-dashboard, run, agent, chat)
│   ├── agent/                   # Agentic layer (NEW)
│   │   ├── __init__.py          # Exports AgentSession
│   │   ├── session.py           # AgentSession — Python API, REPL, undo/redo
│   │   ├── tools.py             # 13 OpenAI function/tool schemas
│   │   ├── tool_executor.py     # Dispatches tool calls to WorkbookState mutations
│   │   ├── llm_bridge.py        # ToolCallingBridge — wraps OpenAI tool-calling API
│   │   └── registry.py          # ObjectRegistry — tracks every artifact by ID
│   ├── chat/                    # Legacy chat mode
│   │   ├── engine.py            # Interactive chat REPL with JSON-based LLM actions
│   │   ├── models.py            # State models (WorkbookState, SheetLayout, cell ops)
│   │   ├── prompts.py           # System/user prompt construction for chat LLM
│   │   ├── layout.py            # Layout engine for object positioning on sheets
│   │   └── renderer.py          # FlexibleTemplate — renders WorkbookState to xlsx
│   ├── dashboard/
│   │   ├── dashboard_engine.py  # Main orchestrator (profile → LLM → render → save)
│   │   ├── llm_client.py        # Provider-agnostic LLM client with JSON repair
│   │   ├── template_selector.py # LLM-powered template/chart/KPI selection
│   │   ├── deep_analysis.py     # Statistical pre-computation + LLM interpretation
│   │   ├── themes.py            # Color theme definitions
│   │   ├── xl_chart.py          # xlsxwriter chart builders (bar, line, pie, etc.)
│   │   ├── xl_style.py          # Format factory for consistent Excel styling
│   │   ├── xl_dynamic.py        # SUMIFS formulas, data validation, sparklines
│   │   └── templates/           # 7 template implementations
│   │       ├── base_xl_template.py
│   │       ├── executive_xl.py
│   │       ├── hr_xl.py
│   │       ├── dark_operational_xl.py
│   │       ├── financial_xl.py
│   │       ├── supply_chain_xl.py
│   │       ├── marketing_xl.py
│   │       └── minimal_clean_xl.py
│   └── data/
│       ├── data_engine.py       # Dataset profiling, multi-sheet discovery & join
│       └── generators/          # 9 synthetic data generators
│           ├── base.py
│           ├── extreme_load.py
│           ├── moderate.py
│           ├── feature_rich.py
│           ├── sparse.py
│           ├── finance.py
│           ├── supply_chain.py
│           ├── executive.py
│           ├── hr_admin.py
│           └── marketing.py
├── skills.md                    # Agent skill manifest (capabilities & tool reference)
├── data/
│   ├── input/                   # Sample CSV datasets (IKEA, YouTube, Netflix, Nykaa)
│   └── output/                  # Generated dashboard outputs
├── output/                      # Default dashboard output directory
├── Dashboard Screenshots/       # Reference dashboard screenshots
├── pyproject.toml               # Build config (hatchling) and dependencies
└── .env                         # API keys and LLM configuration (gitignored)
```

---

## Dependencies

| Package | Purpose |
|---------|---------|
| `openpyxl` | Excel file reading |
| `pandas` | Data manipulation and profiling |
| `numpy` | Numerical computations |
| `openai` | LLM API client (OpenAI & MiniMax) |
| `pydantic` | Data validation and models |
| `pydantic-settings` | Environment-based configuration |
| `typer` | CLI framework |
| `rich` | Terminal formatting and progress |
| `python-dotenv` | `.env` file loading |
| `xlsxwriter` | Excel dashboard rendering |
| `pillow` | Image processing support |

---

## Agent Tools Reference

See [`skills.md`](skills.md) for the complete tool reference including all 13 tools, their parameters, and usage examples.

| # | Tool | Description |
|---|------|-------------|
| 1 | `add_chart` | Bar, line, pie, area, scatter, doughnut, horizontal bar, combo |
| 2 | `modify_object` | Modify any object by ID (charts, tables, KPIs, text) |
| 3 | `remove_object` | Remove any object by ID |
| 4 | `add_kpi_row` | Row of metric tiles with aggregation and formatting |
| 5 | `add_table` | Data table or pivot table |
| 6 | `add_content` | Title bar, section header, or text block |
| 7 | `write_cells` | Write values, formulas, formatting to individual cells |
| 8 | `format_range` | Bold, colors, borders, number format on cell ranges |
| 9 | `sheet_operation` | Create, rename, delete, reorder, hide/show sheets |
| 10 | `row_col_operation` | Resize or hide/show rows and columns |
| 11 | `add_excel_feature` | Conditional formatting, data validation, freeze panes, zoom, merge, hyperlinks, comments, images |
| 12 | `change_theme` | Switch color theme |
| 13 | `query_workbook` | Read-only inspection (list objects, details, data summary, registry) |

---

## License

This project is provided as-is for educational and internal use.
