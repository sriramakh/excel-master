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

On top of this, an **Agent Layer** provides an OpenAI tool-calling interface with 13 tools, an object registry, and a full Python API (`AgentSession`). A legacy **Chat Mode** wraps the pipeline in an interactive JSON-action REPL.

---

## Module Architecture

```
src/excelmaster/
│
├── config.py                    # Centralized configuration
├── models.py                    # Shared Pydantic data models
│
├── cli/                         # User-facing CLI layer
│   └── app.py                   # Typer commands (incl. `agent` command)
│
├── agent/                       # Agentic layer (NEW)
│   ├── __init__.py              # Exports AgentSession
│   ├── session.py               # AgentSession: Python API, REPL, undo/redo
│   ├── tools.py                 # 13 OpenAI function/tool schemas
│   ├── tool_executor.py         # Dispatches tool calls → WorkbookState mutations
│   ├── llm_bridge.py            # ToolCallingBridge: OpenAI tool-calling wrapper
│   └── registry.py              # ObjectRegistry: tracks artifacts by ID
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
└── chat/                        # Legacy interactive chat mode
    ├── engine.py                # Chat REPL + JSON-action execution
    ├── models.py                # WorkbookState, SheetLayout, cell-level ops
    ├── prompts.py               # LLM prompt construction
    ├── layout.py                # Object positioning engine
    └── renderer.py              # FlexibleTemplate (state → xlsx + cell ops)
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

## Agent Layer Architecture (`agent/`)

The agent layer provides an OpenAI tool-calling interface on top of the shared `WorkbookState` and rendering pipeline.

### Component Overview

```
AgentSession (session.py)
├── load()                        # Read data, profile, init conversation
├── execute_instruction(text)     # NL → LLM tool calling → execute → result
├── execute_tool(name, args)      # Direct tool call (no LLM, instant)
├── auto_dashboard()              # LLM template selection → state
├── save(path)                    # Render state → xlsx via FlexibleTemplate
├── undo() / redo()               # State snapshots (max 30)
├── get_state()                   # Full state dict for external agents
├── get_tools()                   # Tool schemas for discovery
└── run_repl()                    # Interactive CLI loop
```

### Tool Calling Flow

```
User instruction
    │
    ▼
AgentSession.execute_instruction()
    ├── Build user message (state snapshot + registry + instruction)
    ├── Send to ToolCallingBridge.call_with_tools()
    │   └── OpenAI chat.completions.create(tools=TOOLS, tool_choice="auto")
    │       └── Returns: assistant text + tool_calls[]
    │
    ├── For each tool_call:
    │   ├── ToolExecutor.execute(tool_name, args)
    │   │   └── Dispatches to one of 13 handler methods
    │   │       └── Mutates WorkbookState + registers in ObjectRegistry
    │   └── Collect results
    │
    ├── Send tool results back to LLM (multi-round loop, max 5 rounds)
    │   └── LLM may issue more tool_calls or return final text
    │
    └── Return {text, actions[], object_ids[]}
```

### 13 Agent Tools (`tools.py` + `tool_executor.py`)

| # | Tool | State Mutation |
|---|------|----------------|
| 1 | `add_chart` | Creates `PlacedChart` → `LayoutEngine.insert_object()` → registry |
| 2 | `modify_object` | Finds object by ID, patches payload fields in-place |
| 3 | `remove_object` | Removes from sheet, reflows layout, removes from registry |
| 4 | `add_kpi_row` | Creates `PlacedKPIRow` with validated `KPIConfig` list |
| 5 | `add_table` | Creates `PlacedTable` (data) or `PlacedPivot` (pivot) |
| 6 | `add_content` | Creates `PlacedTitle`, `PlacedSectionHeader`, or `PlacedText` |
| 7 | `write_cells` | Appends `CellWrite` ops to `SheetLayout.cell_writes` |
| 8 | `format_range` | Appends `CellFormatOp` to `SheetLayout.cell_formats` |
| 9 | `sheet_operation` | Creates/renames/deletes/reorders sheets, sets tab color, hide/show |
| 10 | `row_col_operation` | Sets `row_heights`/`col_widths`/`hidden_rows`/`hidden_cols` |
| 11 | `add_excel_feature` | Appends `ConditionalFormatOp`, `DataValidationOp`, `MergeOp`, `HyperlinkOp`, `CommentOp`, `ImageOp`, or sets freeze/zoom |
| 12 | `change_theme` | Sets `WorkbookState.theme_key` |
| 13 | `query_workbook` | Read-only: lists objects, object details, data summary, sheets, registry |

### Object Registry (`registry.py`)

```
ObjectRegistry
├── register(op_type, sheet, location, description, turn, params) → entry_id
├── get(entry_id) → RegistryEntry | None
├── list_all(sheet?, op_type?) → list[RegistryEntry]
├── remove(entry_id)
├── to_snapshot() → str              # Compact text for LLM context injection
├── snapshot_dict() → list[dict]     # Serializable for undo snapshots
├── restore(data) → None             # Restore from snapshot
└── clear()
```

`RegistryEntry` tracks: `id`, `op_type`, `sheet`, `location`, `description`, `created_at`, `turn`, `params`.

The registry serves dual purposes:
1. **LLM context** — injected into each turn so the LLM knows what exists and can reference IDs
2. **Undo/redo** — snapshot and restore alongside `WorkbookState`

### ToolCallingBridge (`llm_bridge.py`)

Wraps OpenAI's `chat.completions.create` with tool definitions:

```
ToolCallingBridge
├── call_with_tools(messages) → (text, tool_calls[])
│   └── Retry loop with exponential backoff
├── send_tool_results(messages, results) → (text, more_tool_calls[])
│   └── Sends tool result messages for follow-up rounds
└── build_assistant_tool_call_message(text, tool_calls) → dict
    └── Builds proper assistant message with tool_calls for history
```

System preamble enforces: use exact column names, query before modifying, concise responses, appropriate chart types.

---

## Shared State Model (`chat/models.py`)

Both the agent layer and legacy chat mode share the same state model:

```
WorkbookState
├── title: str
├── theme_key: str
├── version: int
└── sheets: list[SheetLayout]
    ├── name, freeze_row, freeze_col, zoom, tab_color, hidden
    ├── objects: list[PlacedObject]       # High-level dashboard objects
    │   ├── id, type, anchor_row, height_rows
    │   └── payload: PlacedTitle | PlacedChart | PlacedKPIRow | PlacedTable
    │              | PlacedPivot | PlacedSectionHeader | PlacedText
    │              | PlacedFilterPanel
    ├── cell_writes: list[CellWrite]              # Individual cell values/formulas
    ├── cell_formats: list[CellFormatOp]          # Range formatting
    ├── conditional_formats: list[ConditionalFormatOp]
    ├── data_validations: list[DataValidationOp]
    ├── merges: list[MergeOp]
    ├── hyperlinks: list[HyperlinkOp]
    ├── comments: list[CommentOp]
    ├── images: list[ImageOp]
    ├── row_heights: dict[int, float]
    ├── col_widths: dict[int, float]
    ├── hidden_rows: list[int]
    └── hidden_cols: list[int]
```

The renderer (`FlexibleTemplate`) processes high-level objects first, then applies all cell-level operations in order: row heights → col widths → hidden rows/cols → cell writes → cell formats → merges → conditional formats → data validations → hyperlinks → comments → images.

---

## Legacy Chat Mode Architecture (`chat/`)

The legacy chat system uses JSON-based action parsing (not OpenAI tool calling):

```
ChatEngine
├── WorkbookState              # Shared state model (see above)
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
├── LayoutEngine               # Object positioning (shared with agent)
│   ├── generate_id()          # Unique IDs per object type
│   ├── insert_object()        # Place at end, after:id, or row:N
│   ├── find_half_pair_row()   # Pair half-width charts side-by-side
│   └── reflow()               # Recompute all positions after remove/move
│
└── Undo/Redo Stack            # Deep copies of WorkbookState (max 30)
```

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
                    │  agent (NEW)          │
                    │  chat (legacy)        │
                    │  profile / list       │
                    └──────────┬───────────┘
                               │
         ┌─────────────────────┼─────────────────────┐
         ▼                     ▼                      ▼
  ┌────────────┐      ┌──────────────┐      ┌─────────────────┐
  │  Generators │      │  Dashboard   │      │  AgentSession   │
  │  (9 types)  │      │  Engine      │      │  (tool calling) │
  └──────┬─────┘      └──────┬───────┘      └────────┬────────┘
         │                    │                       │
         ▼                    ▼                       ▼
  ┌────────────┐      ┌──────────────┐      ┌─────────────────┐
  │  .xlsx     │      │  Profiling   │      │  ToolExecutor   │
  │  datasets  │      │  + LLM Call  │      │  + Registry     │
  └────────────┘      └──────┬───────┘      └────────┬────────┘
                              │                       │
                              ▼                       ▼
                       ┌──────────────┐      ┌─────────────────┐
                       │  Template    │      │  WorkbookState  │
                       │  Rendering   │      │  + Cell Ops     │
                       └──────┬───────┘      └────────┬────────┘
                              │                       │
                              ▼                       ▼
                       ┌──────────────────────────────────┐
                       │  FlexibleTemplate (xlsxwriter)    │
                       │  ├── Data sheet                   │
                       │  ├── Calculations (SUMIFS)        │
                       │  ├── Dashboard (objects + cells)  │
                       │  ├── Extra sheets                 │
                       │  └── Deep Analysis                │
                       └──────────────┬───────────────────┘
                                      │
                                      ▼
                               ┌────────────┐
                               │  .xlsx      │
                               │  Dashboard  │
                               └────────────┘
```

---

## Key Design Decisions

1. **OpenAI tool calling over JSON parsing** — The agent layer uses native OpenAI function calling (`tool_choice: "auto"`) instead of the legacy approach of parsing raw JSON from LLM text output. This gives structured, validated tool invocations with multi-round follow-up.

2. **Dual API surface** — `execute_instruction()` uses LLM tool calling for natural language; `execute_tool()` bypasses LLM entirely for programmatic/free access. Both share the same `ToolExecutor` and `ObjectRegistry`.

3. **Object Registry** — Every artifact (chart, table, KPI row, cell write, merge, etc.) is tracked with a unique ID, operation type, sheet, location, and creation turn. The registry is injected into every LLM turn as context and is snapshot/restored alongside `WorkbookState` for undo/redo.

4. **Two-layer state model** — `WorkbookState` now has both high-level objects (`PlacedObject` list) and cell-level operations (`CellWrite`, `CellFormatOp`, `ConditionalFormatOp`, etc.). The renderer processes objects first, then overlays cell operations.

5. **xlsxwriter over openpyxl for output** — xlsxwriter supports charts, sparklines, data validation, and conditional formatting natively. openpyxl is used only for reading input files.

6. **SUMIFS-based dynamic formulas** — Instead of static snapshots, dashboards contain live formulas that recalculate when filter dropdowns change in Excel.

7. **Two-phase deep analysis** — Statistics are pre-computed in Python to minimize LLM token usage; the LLM only interprets pre-computed numbers.

8. **Fact-dimension auto-join** — Multi-sheet files are automatically unified by detecting fact tables (most rows + most numeric columns) and joining dimension tables via shared text/ID keys.

9. **JSON repair pipeline** — LLM outputs are aggressively repaired (trailing commas, single quotes, Python booleans, truncated brackets) to maximize reliability for the legacy chat mode.

10. **Universal theme baseline** — All 8 color themes currently share the same universal palette, designed as a single cohesive system. The theme registry is extensible for future differentiation.
