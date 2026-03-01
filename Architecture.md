# Architecture

This document describes the internal architecture of **Excel Master**, an AI-powered Excel dashboard generation engine.

---

## High-Level Overview

Excel Master follows a **pipeline architecture** with four major stages:

```
┌─────────────┐     ┌─────────────┐     ┌─────────────────┐     ┌──────────────┐
│  Data Input  │────▶│  Profiling   │────▶│  LLM Decision   │────▶│  Rendering   │
│  (xlsx/csv)  │     │  & Analysis  │     │  Engine          │     │  (xlsxwriter)│
└─────────────┘     └─────────────┘     └─────────────────┘     └──────────────┘
```

Additionally, a **Chat Mode** wraps the pipeline in an interactive REPL, allowing incremental dashboard construction via natural language.

---

## Module Architecture

```
src/excelmaster/
│
├── config.py                    # Centralized configuration
├── models.py                    # Shared Pydantic data models
│
├── cli/                         # User-facing CLI layer
│   └── app.py                   # Typer commands
│
├── data/                        # Data generation & profiling
│   ├── data_engine.py           # Profiling, multi-sheet join, orchestration
│   └── generators/              # 9 synthetic dataset generators
│       └── base.py              # Abstract base generator
│
├── dashboard/                   # Dashboard generation core
│   ├── dashboard_engine.py      # Main pipeline orchestrator
│   ├── llm_client.py            # LLM abstraction (OpenAI / MiniMax)
│   ├── template_selector.py     # LLM-driven template & config selection
│   ├── deep_analysis.py         # Statistical pre-computation + LLM analysis
│   ├── themes.py                # Color theme registry
│   ├── xl_chart.py              # Chart builder functions
│   ├── xl_style.py              # Format/style factory
│   ├── xl_dynamic.py            # SUMIFS formulas, sparklines, dropdowns
│   └── templates/               # Template implementations
│       ├── base_xl_template.py  # Abstract base with shared rendering logic
│       ├── executive_xl.py      # Executive Summary template
│       ├── hr_xl.py             # HR Analytics template
│       ├── dark_operational_xl.py
│       ├── financial_xl.py
│       ├── supply_chain_xl.py
│       ├── marketing_xl.py
│       └── minimal_clean_xl.py
│
└── chat/                        # Interactive chat mode
    ├── engine.py                # Chat REPL + action execution
    ├── models.py                # WorkbookState, SheetLayout, actions
    ├── prompts.py               # LLM prompt construction
    ├── layout.py                # Object positioning engine
    └── renderer.py              # FlexibleTemplate (state → xlsx)
```

---

## Core Pipeline (`dashboard_engine.py`)

The `DashboardEngine` class orchestrates the standard (non-chat) pipeline:

```
DashboardEngine.build(data_path)
│
├── 1. Profile Dataset
│   └── data_engine.profile_dataset()
│       ├── Read xlsx/csv
│       ├── Multi-sheet discovery & join (if >1 sheet)
│       │   └── discover_and_join() — fact table detection, dimension joins
│       └── Column classification (numeric, categorical, date, text)
│           └── Returns DatasetProfile
│
├── 2. LLM Template Selection
│   └── TemplateSelector.select(profile)
│       ├── Sends DatasetProfile as structured prompt to LLM
│       ├── LLM returns JSON: template, theme, KPIs, charts, filters
│       └── _safe_parse_config() validates & applies fallbacks
│           └── Returns DashboardConfig
│
├── 3. Deep Analysis (optional)
│   ├── compute_deep_stats(df, profile) — pre-computes:
│   │   ├── Data quality metrics (nulls, duplicates)
│   │   ├── Numeric statistics (mean, std, skew, outliers)
│   │   ├── Categorical distributions (top values)
│   │   ├── Correlations (Pearson matrix, |r| > 0.3)
│   │   ├── Trends (period-over-period changes)
│   │   └── Dimensional analysis (top/bottom performers)
│   ├── build_analysis_prompt() — structures stats for LLM
│   └── LLM interprets → DeepAnalysis model
│
├── 4. Template Rendering
│   ├── _get_template_class(config.template) — picks template
│   ├── template.build(df, output_path) → renders xlsx:
│   │   ├── Data sheet (raw data dump)
│   │   ├── Calculations sheet (SUMIFS formula tables)
│   │   ├── Dashboard sheet (KPIs, charts, tables, filters)
│   │   └── Deep Analysis sheet (if analysis available)
│   └── Returns Path to saved file
│
└── Returns BuildResult
```

---

## Data Layer

### DatasetProfile

The `DatasetProfile` model captures everything the LLM needs to make design decisions:

```
DatasetProfile
├── name, file_path, sheet_name
├── rows: int
├── columns: list[ColumnInfo]
│   └── ColumnInfo: name, dtype, unique_values, null_pct, sample_values, min/max
├── industry: str (auto-detected or user-provided)
├── date_columns, numeric_columns, categorical_columns
```

### Multi-Sheet Discovery (`discover_and_join`)

For multi-sheet Excel files, the engine:

1. **Reads all sheets** into DataFrames
2. **Scores each sheet** as a fact table candidate (rows × 0.5 + numeric cols × 100 + total cols × 10 + joinable dims × 500)
3. **Identifies dimension tables** — sheets with more text columns than numeric
4. **Finds join keys** — shared text/ID columns with high uniqueness (≥30%) and value overlap
5. **Left-joins** dimensions into the fact table (capped at 12 new columns per join)
6. **Reverts** if a join causes row explosion (>5% growth = many-to-many)

### Synthetic Data Generators

All 9 generators extend `BaseGenerator`:

```python
class BaseGenerator:
    def __init__(self, output_dir)
    def generate(self) -> dict[str, pd.DataFrame]  # sheet_name → DataFrame
    def save(self) -> Path                           # writes to xlsx
```

Each generator creates industry-specific realistic data (e.g., `FinanceGenerator` creates P&L statements, cash flow, and accounts receivable sheets).

---

## LLM Layer

### LLMClient (`llm_client.py`)

Provider-agnostic client supporting **OpenAI** and **MiniMax**:

```
LLMClient
├── generate_json(system_prompt, user_prompt) → dict
│   └── Retry loop (exponential backoff) → parse → repair → return
├── generate_chat_json(messages) → dict
│   └── Used by chat mode with full conversation history
├── generate_text(system_prompt, user_prompt) → str
└── JSON repair pipeline:
    ├── Strip <think>...</think> blocks
    ├── Extract from markdown code fences
    ├── Fix trailing commas, single quotes, Python booleans
    ├── Remove comments
    └── Auto-close truncated brackets
```

### Template Selector (`template_selector.py`)

Sends a structured prompt with:
- Dataset schema (column names, types, ranges, samples)
- Template selection rules (industry → template mapping)
- Chart selection rules (chart type guidelines)
- KPI rules (industry-specific metric recommendations)

Returns a `DashboardConfig` with all design decisions.

### Deep Analysis (`deep_analysis.py`)

Two-phase approach to minimize LLM token usage:

1. **Pre-compute** statistics in Python (correlations, outliers, trends, distributions)
2. **Send statistics** to LLM for **interpretation** (not raw data)

This produces the `DeepAnalysis` model with executive summary, key findings, quality score, trend insights, and recommendations.

---

## Rendering Layer

### Base Template (`base_xl_template.py`)

Abstract base class (~950 lines) providing shared rendering infrastructure:

```
BaseXLTemplate
├── _init_workbook()           # Create workbook, sheets, theme
├── _write_data_sheet()        # Raw data → "Data" sheet
├── _write_filter_slicer_panel() # Filter dropdowns with data validation
├── _build_engine()            # Initialize DynamicEngine for Calculations
├── _write_kpi_tile()          # Single KPI card with value + sparkline
├── _write_section_header()    # Colored divider row
├── _write_summary_table()     # Data table with striped rows
├── _write_insights_box()      # Findings display panel
├── _write_deep_analysis_sheet() # Full analysis report sheet
└── _close()                   # Finalize and save xlsx
```

### Dynamic Engine (`xl_dynamic.py`)

Creates **live Excel formulas** that respond to filter dropdown changes:

```
Dashboard Filter Cell (e.g., B3 = "North")
        │
        ▼
Calculations Sheet: SUMIFS formulas reference Data sheet
        │
        ▼
Charts reference Calculations ranges → auto-update
```

Key components:
- `make_kpi_formula()` — IF/SUMIFS formula for KPI cells
- `make_calc_formula()` — SUMIFS for chart data tables
- `DynamicEngine` — manages the Calculations sheet cursor and writes formula tables
- `write_sparkline()` — mini trend charts inside KPI cards
- `add_filter_dropdown()` — data validation lists linked to Calculations

### Chart Builders (`xl_chart.py`)

Each chart type has a dedicated builder function:

| Function | Chart Type |
|----------|-----------|
| `build_bar_chart()` | Vertical column or horizontal bar |
| `build_line_chart()` | Smooth/straight line with markers |
| `build_area_chart()` | Stacked area |
| `build_pie_chart()` | Pie with percentage labels |
| `build_doughnut_chart()` | Doughnut |
| `build_scatter_chart()` | Scatter with markers |

`build_xl_chart()` dispatches to the correct builder based on `ChartConfig.type`.

### Style Factory (`xl_style.py`)

Centralized format creation with caching:

```
StyleFactory(workbook, theme)
├── title(), subtitle()         # Header formats
├── filter_label/value()        # Filter bar
├── section_header()            # Colored dividers
├── kpi_bg/label/value/delta()  # KPI card cells
├── table_header/data/num()     # Data tables
├── insight_header/text()       # Findings boxes
├── badge_positive/negative()   # Status indicators
├── analysis_*()                # Deep Analysis sheet formats
└── Internal _cache dict avoids duplicate format creation
```

### Template Implementations

7 templates extend `BaseXLTemplate`, each with a custom `build()` method:

| Template | Key Layout |
|----------|-----------|
| **ExecutiveSummaryXL** | Title → Filters → KPIs → 4 charts (2×2) → Table → Insights |
| **HRAnalyticsXL** | Title → Filters → KPIs → Section headers → Charts → Table |
| **DarkOperationalXL** | Dark background → Dense metrics → Operational charts |
| **FinancialXL** | Title → Financial KPIs → P&L charts → Budget variance table |
| **SupplyChainXL** | Title → Logistics KPIs → Route/carrier charts → Shipment table |
| **MarketingXL** | Title → Campaign KPIs → ROI charts → Channel analysis |
| **MinimalCleanXL** | Clean layout → Minimal styling → General-purpose |

---

## Chat Mode Architecture

The chat system adds an interactive layer on top of the rendering pipeline:

```
ChatEngine
├── WorkbookState              # In-memory dashboard model
│   ├── title, theme_key
│   └── sheets: list[SheetLayout]
│       └── objects: list[PlacedObject]
│           ├── id: str (e.g., "chart_0")
│           ├── type: ObjectType (title, chart, kpi_row, table, etc.)
│           ├── anchor_row: int
│           └── payload: Union[PlacedChart, PlacedKPIRow, PlacedTable, ...]
│
├── REPL Loop
│   ├── Special commands: auto, start, undo, redo, show, reset, save as, quit
│   └── LLM-driven turns:
│       ├── build_user_message(state, input) → snapshot + instruction
│       ├── LLM returns JSON: { message, actions[] }
│       └── Execute actions sequentially:
│           ├── add_chart, modify_chart
│           ├── add_table, modify_table
│           ├── add_kpi_row, modify_kpi
│           ├── add_pivot, add_section_header, add_text
│           ├── remove, move
│           ├── add_sheet, change_theme, change_title
│           └── auto_dashboard
│
├── LayoutEngine               # Object positioning
│   ├── generate_id()          # Unique IDs per object type
│   ├── insert_object()        # Place at end, after:id, or row:N
│   ├── find_half_pair_row()   # Pair half-width charts side-by-side
│   └── reflow()               # Recompute all positions after remove/move
│
├── Undo/Redo Stack            # Deep copies of WorkbookState (max 30)
│
└── FlexibleTemplate           # Renders WorkbookState → xlsx
    ├── Extends BaseXLTemplate
    ├── Converts state to minimal DashboardConfig for base class
    └── Dispatches per-object rendering:
        ├── _render_title()
        ├── _render_kpi_row()
        ├── _render_section_header()
        ├── _render_chart()
        ├── _render_table()
        ├── _render_pivot()
        └── _render_text()
```

### Chat LLM Protocol

The chat system prompt includes:
- Full data schema from `DatasetProfile`
- Complete action type reference with parameter specs
- JSON output format specification

Each user turn sends a **state snapshot** (compact one-line-per-object representation) plus the user instruction. The LLM returns structured actions that are parsed and executed in sequence.

---

## Configuration (`config.py`)

Uses `pydantic-settings` with `.env` file support:

```python
class Settings:
    openai_api_key: str         # OpenAI API key
    minimax_api_token: str      # MiniMax API token
    llm_provider: str           # "openai" | "minimax"
    model: str                  # e.g., "gpt-4o"
    minimax_model: str          # e.g., "MiniMax-Text-01"
    max_tokens: int             # Default 4096
    temperature: float          # Default 0.3
    max_retries: int            # Default 3
    data_dir: Path              # Default "data"
    output_dir: Path            # Default "output"
```

Singleton access via `get_settings()`.

---

## Data Flow Diagram

```
                    ┌──────────────────────┐
                    │    CLI (app.py)       │
                    │  generate-data        │
                    │  build-dashboard      │
                    │  run (full pipeline)  │
                    │  chat                 │
                    │  profile              │
                    │  list                 │
                    └──────────┬───────────┘
                               │
              ┌────────────────┼────────────────┐
              ▼                ▼                 ▼
     ┌────────────┐   ┌──────────────┐   ┌───────────┐
     │  Generators │   │  Dashboard   │   │   Chat    │
     │  (9 types)  │   │  Engine      │   │   Engine  │
     └──────┬─────┘   └──────┬───────┘   └─────┬─────┘
            │                 │                  │
            ▼                 ▼                  ▼
     ┌────────────┐   ┌──────────────┐   ┌───────────────┐
     │  .xlsx     │   │  Profiling   │   │  WorkbookState│
     │  datasets  │   │  + LLM Call  │   │  + LLM Actions│
     └────────────┘   └──────┬───────┘   └───────┬───────┘
                              │                   │
                              ▼                   ▼
                       ┌──────────────┐   ┌───────────────┐
                       │  Template    │   │  Flexible     │
                       │  Rendering   │   │  Template     │
                       └──────┬───────┘   └───────┬───────┘
                              │                   │
                              ▼                   ▼
                       ┌──────────────────────────────┐
                       │  BaseXLTemplate (xlsxwriter)  │
                       │  ├── Data sheet               │
                       │  ├── Calculations (SUMIFS)    │
                       │  ├── Dashboard (KPIs/Charts)  │
                       │  └── Deep Analysis            │
                       └──────────────┬───────────────┘
                                      │
                                      ▼
                               ┌────────────┐
                               │  .xlsx      │
                               │  Dashboard  │
                               └────────────┘
```

---

## Key Design Decisions

1. **xlsxwriter over openpyxl for output** — xlsxwriter supports charts, sparklines, data validation, and conditional formatting natively. openpyxl is used only for reading input files.

2. **SUMIFS-based dynamic formulas** — Instead of static snapshots, dashboards contain live formulas that recalculate when filter dropdowns change in Excel.

3. **Two-phase deep analysis** — Statistics are pre-computed in Python to minimize LLM token usage; the LLM only interprets pre-computed numbers.

4. **Fact-dimension auto-join** — Multi-sheet files are automatically unified by detecting fact tables (most rows + most numeric columns) and joining dimension tables via shared text/ID keys.

5. **JSON repair pipeline** — LLM outputs are aggressively repaired (trailing commas, single quotes, Python booleans, truncated brackets) to maximize reliability.

6. **State-based chat model** — The chat engine maintains a `WorkbookState` model that is fully serializable, enabling undo/redo via deep copies and incremental modifications via structured actions.

7. **Universal theme baseline** — All 8 color themes currently share the same universal palette, designed as a single cohesive system. The theme registry is extensible for future differentiation.
