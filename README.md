# Excel Master

**AI-powered Excel dashboard generation engine** — transform any dataset into professional, interactive Excel dashboards using LLM-driven design decisions.

Excel Master profiles your data, selects the optimal dashboard template, configures KPIs and charts, and renders a fully styled `.xlsx` file — all in one command. It also offers an interactive chat mode where you can build and modify dashboards with natural language.

---

## Features

- **Automatic Dashboard Generation** — Feed in any `.xlsx` or `.csv` file and get a complete dashboard with KPIs, charts, tables, filters, and deep analysis
- **LLM-Powered Design** — Uses OpenAI or MiniMax to select the best template, theme, chart types, and KPI configurations for your data
- **7 Professional Templates** — Executive Summary, HR Analytics, Dark Operational, Financial, Supply Chain, Marketing, and Minimal Clean
- **8 Color Themes** — Corporate Blue, HR Purple, Dark Mode, Supply Green, Finance Green, Marketing Orange, Slate Minimal, Executive Navy
- **Interactive Chat Mode** — Build and modify dashboards conversationally with undo/redo support
- **Deep Analysis** — LLM-interpreted statistical analysis with correlations, outliers, trends, and actionable recommendations
- **Dynamic Formulas** — Dashboards use SUMIFS formulas linked to filter dropdowns, so they stay interactive in Excel
- **9 Synthetic Data Generators** — Generate realistic test datasets across industries (Finance, HR, Supply Chain, Marketing, etc.)
- **Multi-Sheet Intelligence** — Automatically discovers relationships between sheets and joins dimension tables into a unified fact table
- **Rule-Based Fallback** — Works without an LLM using heuristic template and chart selection (`--no-llm`)

---

## Quick Start

### Prerequisites

- Python 3.10+
- An OpenAI API key (or MiniMax API token)

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

# Interactive chat mode
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

## Chat Mode

The interactive chat mode lets you build dashboards conversationally:

```
You: auto
  -> Auto dashboard built (4 charts, 5 KPIs, 1 table)

You: add a pie chart of revenue by region
  -> Added pie chart 'Revenue by Region' [chart_3]

You: change the theme to dark mode
  -> Changed theme to 'dark_mode'

You: remove chart_1
  -> Removed chart 'chart_1'

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
│   │   └── app.py               # Typer CLI (generate-data, build-dashboard, run, chat, etc.)
│   ├── chat/
│   │   ├── engine.py            # Interactive chat REPL with LLM-driven actions
│   │   ├── models.py            # Chat state models (WorkbookState, SheetLayout, PlacedObject)
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

## License

This project is provided as-is for educational and internal use.
