"""
Dynamic Excel engine: SUMIFS formula tables, data validation dropdowns, sparklines.

Architecture:
  - Data sheet: raw data (written by base template)
  - Calculations sheet: aggregated SUMIFS tables that reference Dashboard filter cells
  - Dashboard sheet: filter dropdowns, KPI formula cells, chart anchors

The dynamic chain:
  1. User changes dropdown in Dashboard!B3 (primary filter)
  2. Calculations sheet SUMIFS formulas recalculate
  3. Charts referencing Calculations ranges update automatically
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import Any

import pandas as pd
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

from ..models import ChartConfig, KPIConfig, AggFunc


# ── Excel address helpers ──────────────────────────────────────────────────────

def col_letter(col_idx: int) -> str:
    """0-indexed column index → Excel letter (A, B, … ZZ)."""
    return xl_col_to_name(col_idx)


def cell_abs(row: int, col: int) -> str:
    """0-indexed (row, col) → absolute Excel address like $A$1."""
    return f"${col_letter(col)}${row + 1}"


def range_abs(sheet: str, r1: int, c1: int, r2: int, c2: int) -> str:
    """Absolute sheet range string for xlsxwriter chart series."""
    return f"='{sheet}'!${col_letter(c1)}${r1+1}:${col_letter(c2)}${r2+1}"


# ── Aggregation dispatch ───────────────────────────────────────────────────────

_AGG_MAP = {
    AggFunc.SUM: "sum",
    AggFunc.AVG: "mean",
    AggFunc.COUNT: "count",
    AggFunc.MAX: "max",
    AggFunc.MIN: "min",
    AggFunc.MEDIAN: "median",
    AggFunc.DISTINCT: "nunique",
}


def _excel_agg_func(agg: AggFunc) -> str:
    """Map AggFunc to Excel function name."""
    return {
        AggFunc.SUM: "SUMIFS",
        AggFunc.COUNT: "COUNTIFS",
        AggFunc.AVG: "AVERAGEIFS",
        AggFunc.MAX: "MAXIFS",
        AggFunc.MIN: "MINIFS",
    }.get(agg, "SUMIFS")


# ── SUMIFS formula builder ─────────────────────────────────────────────────────

def make_kpi_formula(
    agg: AggFunc,
    val_col_idx: int,       # index of value column in Data sheet
    filter_col_idx: int,    # index of primary filter column in Data sheet
    filter_cell: str,       # e.g. "$B$3" (on Dashboard sheet, absolute)
    n_rows: int = 100001,   # max data rows to scan
) -> str:
    """
    Build a KPI formula string for the Dashboard sheet.

    Examples:
        filter="All"  → SUM(Data!$F$2:$F$100001)
        filter="X"    → SUMIFS(Data!$F$2:$F$100001, Data!$C$2:$C$100001, "X")

    Returns an Excel formula string (without leading =).
    """
    vl = col_letter(val_col_idx)
    fl = col_letter(filter_col_idx)
    val_rng = f"'Data'!${vl}$2:${vl}${n_rows}"
    fil_rng = f"'Data'!${fl}$2:${fl}${n_rows}"

    if agg == AggFunc.COUNT:
        all_formula = f"COUNTA({val_rng})"
        filtered = f"COUNTIFS({fil_rng},{filter_cell})"
    elif agg == AggFunc.AVG:
        all_formula = f"AVERAGE({val_rng})"
        filtered = f"AVERAGEIFS({val_rng},{fil_rng},{filter_cell})"
    elif agg == AggFunc.DISTINCT:
        # Approximate: SUMPRODUCT(1/COUNTIF(...))
        all_formula = f"SUMPRODUCT(1/COUNTIF({val_rng},{val_rng}))"
        filtered = f"SUMPRODUCT(({fil_rng}={filter_cell})/COUNTIF({val_rng},{val_rng}))"
    else:
        xl_func = _excel_agg_func(agg)
        all_formula = f"SUM({val_rng})"
        filtered = f"{xl_func}({val_rng},{fil_rng},{filter_cell})"

    return f'=IF({filter_cell}="All",{all_formula},{filtered})'


def make_calc_formula(
    agg: AggFunc,
    val_col_idx: int,
    x_col_idx: int,        # chart X column (grouping dimension)
    filter_col_idx: int,
    cat_ref: str,           # e.g. "Calculations!$A$5" — the category cell
    filter_cell: str,       # e.g. "Dashboard!$B$3"
    n_rows: int = 100001,
) -> str:
    """
    SUMIFS formula in Calculations sheet that:
    - When filter="All": aggregates across all filter values for that category
    - When filter="X":   aggregates only rows where filter_col = "X"
    """
    vl = col_letter(val_col_idx)
    xl = col_letter(x_col_idx)
    fl = col_letter(filter_col_idx)
    val_rng = f"'Data'!${vl}$2:${vl}${n_rows}"
    x_rng = f"'Data'!${xl}$2:${xl}${n_rows}"
    fil_rng = f"'Data'!${fl}$2:${fl}${n_rows}"

    if agg == AggFunc.COUNT:
        all_f = f"COUNTIFS({x_rng},{cat_ref})"
        filt_f = f"COUNTIFS({x_rng},{cat_ref},{fil_rng},{filter_cell})"
    elif agg == AggFunc.AVG:
        all_f = f"AVERAGEIFS({val_rng},{x_rng},{cat_ref})"
        filt_f = f"AVERAGEIFS({val_rng},{x_rng},{cat_ref},{fil_rng},{filter_cell})"
    else:
        all_f = f"SUMIFS({val_rng},{x_rng},{cat_ref})"
        filt_f = f"SUMIFS({val_rng},{x_rng},{cat_ref},{fil_rng},{filter_cell})"

    return f'=IF({filter_cell}="All",{all_f},{filt_f})'


# ── Calculations table writer ──────────────────────────────────────────────────

@dataclass
class CalcTable:
    """Represents one aggregation table on the Calculations sheet."""
    chart_idx: int
    header: str           # column header label
    cat_start_row: int    # 0-indexed row where categories begin
    val_start_row: int    # same as cat_start_row (parallel)
    n_rows: int           # number of category rows
    cat_col: int = 0      # column index in Calculations sheet
    val_col: int = 1      # column index for values

    @property
    def cat_range(self) -> str:
        """xlsxwriter range string for category column."""
        return range_abs("Calculations", self.cat_start_row, self.cat_col,
                          self.cat_start_row + self.n_rows - 1, self.cat_col)

    @property
    def val_range(self) -> str:
        return range_abs("Calculations", self.val_start_row, self.val_col,
                          self.val_start_row + self.n_rows - 1, self.val_col)

    @property
    def header_cell(self) -> str:
        """Absolute address of header cell (for chart series name)."""
        return f"='Calculations'!${col_letter(self.val_col)}${self.val_start_row}"


@dataclass
class DynamicEngine:
    """
    Manages writing to the Calculations sheet and building dynamic references.
    """
    ws_calc: Any    # xlsxwriter worksheet object
    df: pd.DataFrame
    col_index: dict[str, int]    # column_name → 0-indexed position in Data sheet
    filter_col: str              # primary filter column name
    dashboard_filter_cell: str = "'Dashboard'!$B$3"   # absolute cell ref

    # Internal cursor: next free row in Calculations sheet
    _row_cursor: int = field(default=0, init=False)

    def __post_init__(self):
        self._row_cursor = 2   # leave rows 0-1 for metadata

    def _data_col(self, name: str) -> int:
        return self.col_index.get(name, 0)

    def _n_data_rows(self) -> int:
        return min(len(self.df) + 1, 100001)

    def write_chart_table(self, chart_cfg: ChartConfig) -> CalcTable | None:
        """
        Write aggregation table for one chart to the Calculations sheet.
        Returns a CalcTable describing the ranges written.
        """
        x_col = chart_cfg.x_column
        y_cols = [c for c in chart_cfg.y_columns if c in self.col_index]
        if x_col not in self.col_index or not y_cols:
            return None

        # Get unique categories (top N or limited to 20)
        agg_func = _AGG_MAP.get(chart_cfg.aggregation, "sum")
        try:
            grouped = (self.df.groupby(x_col)[y_cols[0]]
                       .agg(agg_func)
                       .reset_index()
                       .sort_values(y_cols[0], ascending=False))
            top_n = chart_cfg.top_n or 15
            grouped = grouped.head(min(top_n, 20))
        except Exception:
            return None

        categories = grouped[x_col].tolist()
        if not categories:
            return None

        # Write header row
        hdr_row = self._row_cursor
        from xlsxwriter.utility import xl_col_to_name as cln
        self.ws_calc.write(hdr_row, 0, x_col, None)
        self.ws_calc.write(hdr_row, 1, y_cols[0], None)
        self._row_cursor += 1

        # Write category labels + SUMIFS formulas
        data_start = self._row_cursor
        fc = self._data_col(self.filter_col)
        xc = self._data_col(x_col)
        n_rows = self._n_data_rows()

        for idx, cat in enumerate(categories):
            r = self._row_cursor
            # Write category label
            self.ws_calc.write(r, 0, cat)

            # Build SUMIFS formula for each y_column
            for yi, y_col in enumerate(y_cols[:4]):
                vc = self._data_col(y_col)
                cat_ref = f"'Calculations'!${col_letter(0)}${r + 1}"
                formula = make_calc_formula(
                    chart_cfg.aggregation, vc, xc, fc,
                    cat_ref, self.dashboard_filter_cell, n_rows,
                )
                # Pre-compute result so xlsxwriter caches the real value
                # (without this, charts read cached 0 instead of formula result)
                try:
                    result_val = float(grouped.iloc[idx][y_col])
                except Exception:
                    result_val = 0
                self.ws_calc.write_formula(r, 1 + yi, formula, None, result_val)

            self._row_cursor += 1

        # Spacer
        self._row_cursor += 2

        table = CalcTable(
            chart_idx=0,
            header=y_cols[0],
            cat_start_row=data_start,
            val_start_row=data_start,
            n_rows=len(categories),
            cat_col=0,
            val_col=1,
        )
        return table

    def write_filter_options(self, filter_col: str,
                              extra_col_offset: int = 10) -> list[str]:
        """
        Write unique filter values to the Calculations sheet (col offset 10+).
        Returns the list ["All", val1, val2, ...] for data validation.
        """
        if filter_col not in self.df.columns:
            return ["All"]

        unique_vals = sorted(self.df[filter_col].dropna().unique().tolist())[:30]
        options = ["All"] + [str(v) for v in unique_vals]

        # Write to Calculations sheet col 10
        for i, v in enumerate(options):
            self.ws_calc.write(i, extra_col_offset, v)

        return options

    def write_metadata(self) -> None:
        """Write header metadata rows to Calculations sheet."""
        self.ws_calc.write(0, 0, "Excel Master — Calculations Sheet")
        self.ws_calc.write(1, 0, "Auto-generated SUMIFS tables. Do not edit manually.")


# ── Sparkline writer ───────────────────────────────────────────────────────────

def write_sparkline(ws, row: int, col: int,
                    data_range: str, color: str = "#1B3D6E",
                    sparkline_type: str = "line") -> None:
    """
    Write a sparkline to a cell.

    Args:
        ws: xlsxwriter worksheet
        row, col: 0-indexed cell coordinates
        data_range: e.g. "=Data!$F$2:$F$13" (last 12 months)
        color: hex color string
        sparkline_type: "line", "column", or "win_loss"
    """
    try:
        ws.add_sparkline(row, col, {
            "range": data_range.lstrip("="),
            "type": sparkline_type,
            "series_color": color.lstrip("#"),
            "negative_points": True,
            "last_point": True,
            "markers": True,
        })
    except Exception:
        pass   # sparklines may not be available in all xlsxwriter builds


def build_kpi_sparkline_range(df: pd.DataFrame, val_col: str,
                               time_col: str | None = None,
                               n_periods: int = 12) -> str:
    """
    Build a sparkline data range from the last N periods.
    Returns a range string pointing to Data sheet, or "" if not applicable.
    """
    if val_col not in df.columns:
        return ""

    try:
        if time_col and time_col in df.columns:
            # Monthly aggregation
            sub = (df.groupby(time_col)[val_col]
                   .sum()
                   .tail(n_periods)
                   .reset_index())
            val_series = sub[val_col]
        else:
            val_series = df[val_col].dropna().tail(n_periods)

        if len(val_series) < 2:
            return ""

        # We'll write sparkline data directly — return the series for inline writing
        return val_series.tolist()
    except Exception:
        return ""


# ── Data validation helper ─────────────────────────────────────────────────────

def add_filter_dropdown(ws, row: int, col: int,
                         options: list[str],
                         calc_sheet_range: str | None = None) -> None:
    """
    Add data validation dropdown to a cell.

    Prefers a calc_sheet_range reference (for many options), falls back to inline list.
    """
    if calc_sheet_range:
        ws.data_validation(row, col, row, col, {
            "validate": "list",
            "source": calc_sheet_range,
            "input_title": "Filter",
            "input_message": "Select a value to filter the dashboard",
        })
    elif options:
        # Inline list (max ~255 chars total)
        short = options[:20]
        ws.data_validation(row, col, row, col, {
            "validate": "list",
            "source": short,
            "input_title": "Filter",
            "input_message": "Select a value to filter the dashboard",
        })
