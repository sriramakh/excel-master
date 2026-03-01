"""
Base xlsxwriter template for Excel Master dashboards.

This base class provides ONLY utility helpers — no shared layout constants.
Each subclass defines its own spatial layout from scratch.
"""
from __future__ import annotations
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

from ...models import (
    DashboardConfig, DatasetProfile, DeepAnalysis,
    KPIConfig, ChartConfig, AggFunc, NumberFormat,
)
from ..themes import Theme, get_theme
from ..xl_style import StyleFactory, _hex
from ..xl_chart import ChartZone, build_xl_chart
from ..xl_dynamic import (
    DynamicEngine, CalcTable,
    make_kpi_formula, write_sparkline, build_kpi_sparkline_range,
    add_filter_dropdown, col_letter, range_abs,
)

# Shared pixel constants
# Default row ≈ 15pt ≈ 20px.  280px chart spans ~14 rows, 250px spans ~13 rows.
CHART_HALF_W = 640
CHART_HALF_H = 280          # ~14 rows at default height
CHART_FULL_W = 1295
CHART_FULL_H = 250          # ~13 rows at default height
CHART_THIRD_W = 425
CHART_QUARTER_W = 315
N_COLS = 24
COL_W = 7.3


class BaseXLTemplate(ABC):
    """Abstract base for all xlsxwriter dashboard templates.

    Subclasses define their own layout in build(). The base class
    provides helpers for writing data, filter controls, KPI cells,
    chart tables, and summary tables.
    """

    name: str = "base_xl"

    def __init__(self, config: DashboardConfig):
        self.config = config
        self.theme = get_theme(config.theme)
        self._wb = None
        self._sf = None
        self._ws_dash = None
        self._ws_data = None
        self._ws_calc = None
        self._ws_analysis = None
        self._engine: DynamicEngine | None = None
        self._col_index: dict[str, int] = {}
        self._primary_filter_col: str | None = None
        self._filter_cell: str = '"All"'
        self._filter_row: int = 2
        self._filter_col_idx: int = 3   # Column D by default
        self._analysis_df: pd.DataFrame | None = None
        self._analysis_profile: DatasetProfile | None = None

    @abstractmethod
    def build(self, df: pd.DataFrame, output_path: Path) -> Path: ...

    # ── Workbook initialization ────────────────────────────────────────────────

    def _init_workbook(self, output_path: Path, col_width: float = COL_W):
        output_path.parent.mkdir(parents=True, exist_ok=True)
        self._wb = xlsxwriter.Workbook(str(output_path), {
            "constant_memory": False,
            "strings_to_numbers": True,
            "default_date_format": "yyyy-mm-dd",
        })
        self._sf = StyleFactory(self._wb, self.theme)
        self._ws_dash = self._wb.add_worksheet("Dashboard")
        self._ws_data = self._wb.add_worksheet("Data")
        self._ws_calc = self._wb.add_worksheet("Calculations")
        self._ws_calc.hide()
        self._ws_analysis = self._wb.add_worksheet("Deep Analysis")

        # Default column widths
        self._ws_dash.set_column(0, N_COLS - 1, col_width)

    def _write_data_sheet(self, df: pd.DataFrame, n_rows: int = 100000):
        """Write raw data to Data sheet as an Excel Table with AutoFilter."""
        ws = self._ws_data
        sf = self._sf
        hdr_fmt = sf.data_header()
        dat_fmt = sf.data_cell()

        cols = list(df.columns)
        self._col_index = {c: i for i, c in enumerate(cols)}

        for j, c in enumerate(cols):
            ws.write(0, j, c, hdr_fmt)

        data_subset = df.head(n_rows)
        for i, row in enumerate(data_subset.itertuples(index=False), 1):
            for j, val in enumerate(row):
                try:
                    if pd.isna(val) if not isinstance(val, (str, bool)) else False:
                        ws.write_blank(i, j, None, dat_fmt)
                    else:
                        ws.write(i, j, val, dat_fmt)
                except Exception:
                    ws.write(i, j, str(val) if val is not None else "", dat_fmt)

        # Excel Table with AutoFilter (real dropdown filters on every column)
        # Must pass 'columns' so add_table doesn't overwrite manually-written headers
        ws.add_table(0, 0, min(len(data_subset), n_rows), len(cols) - 1, {
            "name": "DataTable",
            "style": "Table Style Medium 9",
            "autofilter": True,
            "columns": [{"header": c, "header_format": hdr_fmt} for c in cols],
        })
        ws.freeze_panes(1, 0)

    # ── Filter panel helpers ────────────────────────────────────────────────────

    def _detect_filter_cols(self, df: pd.DataFrame,
                             prefer_keywords: list[str] | None = None) -> list[str]:
        """Detect up to 3 good filter columns from the dataframe."""
        preferred = prefer_keywords or []
        obj_cols = df.select_dtypes("object").columns.tolist()
        # Exclude obvious ID/name columns
        exclude = ["id", "_id", "name", "description", "url", "email", "address"]
        obj_cols = [c for c in obj_cols
                    if not any(x in c.lower() for x in exclude)]

        # Score by uniqueness (good filters: 2-30 unique values)
        scored = sorted(
            [(c, df[c].nunique()) for c in obj_cols],
            key=lambda x: abs(x[1] - 8)  # prefer ~8 unique values
        )

        # Prefer config dimension
        result = []
        if self.config.primary_dimension and self.config.primary_dimension in df.columns:
            result.append(self.config.primary_dimension)

        # Then filters from config
        for fc in self.config.filters:
            if fc.column in df.columns and fc.column not in result:
                result.append(fc.column)

        # Then preferred keywords
        for kw in preferred:
            for c, _ in scored:
                if kw in c.lower() and c not in result:
                    result.append(c)
                    break

        # Fill up with well-scored columns
        for c, n in scored:
            if len(result) >= 3:
                break
            if c not in result and 2 <= n <= 40:
                result.append(c)

        return result[:3]

    def _write_filter_slicer_panel(self, df: pd.DataFrame,
                                    row: int,
                                    filter_cols: list[str],
                                    panel_bg: str | None = None) -> dict[str, str]:
        """
        Write a prominent FILTER PANEL that looks like Excel slicer buttons.

        Creates a full-width row with:
         - "🔽 FILTERS:" label on the left
         - One dropdown control per filter column
         - Each control is a large merged cell with data validation

        Returns a dict: {column_name: excel_formula_ref_for_filter_cell}
        """
        ws = self._ws_dash
        sf = self._sf
        t = self.theme
        bg = panel_bg or t.bg_dashboard

        # Styles for filter panel
        panel_bg_fmt = sf.bg(bg)
        filter_label_fmt = self._wb.add_format({
            "font_name": t.font_body, "font_size": 8, "bold": True,
            "font_color": _hex(t.text_muted), "bg_color": _hex(bg),
            "align": "right", "valign": "vcenter",
        })
        filter_title_fmt = self._wb.add_format({
            "font_name": t.font_heading, "font_size": 9, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.secondary),
            "align": "left", "valign": "vcenter", "indent": 1,
        })
        filter_cell_fmt = self._wb.add_format({
            "font_name": t.font_body, "font_size": 10, "bold": True,
            "font_color": _hex(t.primary),
            "bg_color": _hex(t.bg_card),
            "align": "left", "valign": "vcenter",
            "indent": 1,
            "border": 2,
            "border_color": _hex(t.primary),
        })
        hint_fmt = self._wb.add_format({
            "font_name": t.font_body, "font_size": 7, "italic": True,
            "font_color": _hex(t.text_muted), "bg_color": _hex(bg),
            "align": "left", "valign": "vcenter", "indent": 1,
        })

        ws.set_row(row, 30)
        ws.set_row(row + 1, 14)  # hint row

        # Fill background cell-by-cell (never merge full row — sub-merges would overlap)
        for c in range(N_COLS):
            ws.write_blank(row, c, None, panel_bg_fmt)
            ws.write_blank(row + 1, c, None, panel_bg_fmt)

        # "FILTERS:" label on left (written individually, not merged)
        ws.write(row, 0, "🔽 FILTERS", filter_title_fmt)
        ws.write(row + 1, 0, "Click to filter", hint_fmt)

        filter_refs = {}
        calc_col_offset = 12  # offset in Calculations sheet for filter option lists

        for fi, fc in enumerate(filter_cols[:3]):
            # Position: each filter control spans 6 cols (no separate arrow cell)
            c_start = 2 + fi * 7
            c_end = c_start + 5  # inclusive end for merge

            if c_end >= N_COLS:
                break

            # Label row (sub-label below each dropdown cell)
            label = fc.replace("_", " ").title()
            ws.write(row + 1, c_start, f"  {label}:", hint_fmt)

            # Dropdown cell — entire width is the clickable dropdown (no separate arrow)
            ws.merge_range(row, c_start, row, c_end, "All", filter_cell_fmt)

            # Write unique options to Calculations sheet
            if fc in df.columns:
                unique_vals = sorted(
                    [str(v) for v in df[fc].dropna().unique().tolist()]
                )[:30]
                options = ["All"] + unique_vals
                for vi, v in enumerate(options):
                    self._ws_calc.write(vi, calc_col_offset + fi, v)
                calc_range = (
                    f"Calculations!${col_letter(calc_col_offset + fi)}$1:"
                    f"${col_letter(calc_col_offset + fi)}${len(options)}"
                )
                add_filter_dropdown(ws, row, c_start, options, calc_range)

            # Store the absolute reference to the filter cell
            filter_refs[fc] = f"'Dashboard'!${col_letter(c_start)}${row + 1}"

        return filter_refs

    # ── Engine setup ─────────────────────────────────────────────────────────

    def _build_engine(self, df: pd.DataFrame,
                       filter_refs: dict[str, str] | None = None,
                       n_data_rows: int = 100000) -> DynamicEngine:
        """Initialize DynamicEngine with the primary filter column reference.

        n_data_rows must match what _write_data_sheet() wrote so that
        pre-computed cached values align with what SUMIFS will return in Excel.
        """
        filter_refs = filter_refs or {}

        # Find the best primary filter column and its formula reference
        if filter_refs:
            fc = list(filter_refs.keys())[0]
            fc_ref = filter_refs[fc]
        elif self.config.primary_dimension and self.config.primary_dimension in df.columns:
            fc = self.config.primary_dimension
            fc_ref = '"All"'
        else:
            obj_cols = df.select_dtypes("object").columns.tolist()
            fc = obj_cols[0] if obj_cols else ""
            fc_ref = '"All"'

        self._primary_filter_col = fc
        self._filter_cell = fc_ref

        # Truncate to the same row count written to the Data sheet so that
        # pre-computed cached values match what SUMIFS will produce in Excel
        df_engine = df.head(n_data_rows) if len(df) > n_data_rows else df

        self._engine = DynamicEngine(
            ws_calc=self._ws_calc,
            df=df_engine,
            col_index=self._col_index,
            filter_col=fc,
            dashboard_filter_cell=fc_ref,
        )
        self._engine.write_metadata()
        return self._engine

    def _add_chart(self, chart_cfg: ChartConfig, df: pd.DataFrame,
                   zone: ChartZone) -> None:
        """Write chart calc table → build xlsxwriter chart → insert into dashboard."""
        if self._engine is None:
            return
        table = self._engine.write_chart_table(chart_cfg)
        if table is None:
            return
        val_ranges = [(table.header_cell, table.val_range)]
        chart = build_xl_chart(
            self._wb, chart_cfg, self.theme,
            table.cat_range, val_ranges, zone,
        )
        self._ws_dash.insert_chart(
            zone.row, zone.col, chart,
            {"x_offset": zone.x_offset, "y_offset": zone.y_offset},
        )

    # ── KPI writing helpers ───────────────────────────────────────────────────

    def _write_kpi_tile(self, row: int, col: int, span_cols: int, span_rows: int,
                         kpi: KPIConfig, df: pd.DataFrame, bg: str,
                         font_color: str | None = None,
                         filter_ref: str | None = None) -> None:
        """Write a single KPI tile at (row, col) spanning span_cols × span_rows."""
        ws = self._ws_dash
        sf = self._sf
        t = self.theme
        fr = filter_ref or self._filter_cell

        bg_fmt = sf.kpi_bg(bg)
        lbl_fmt = sf.kpi_label(bg)
        val_fmt = sf.kpi_value(bg, font_color or (t.text_light if t.dark_mode else None))

        # Top padding row
        ws.merge_range(row, col, row, col + span_cols - 1, "", bg_fmt)

        # Label row (second row of tile)
        lbl_row = row + 1
        ws.set_row(lbl_row, 15)
        label_txt = kpi.icon + "  " + kpi.label if kpi.icon else kpi.label
        ws.merge_range(lbl_row, col, lbl_row, col + span_cols - 1, label_txt, lbl_fmt)

        # Value row (third row of tile)
        val_row = row + 2
        ws.set_row(val_row, 32)
        if kpi.column in self._col_index:
            vc = self._col_index[kpi.column]
            fc_idx = self._col_index.get(self._primary_filter_col, 0) \
                if self._primary_filter_col and self._primary_filter_col in self._col_index \
                else None
            if fc_idx is not None:
                formula = make_kpi_formula(
                    kpi.aggregation, vc, fc_idx, fr,
                    n_rows=min(len(df), 100000) + 1,
                )
                # Pre-compute result so xlsxwriter caches the real value.
                # Use the same 100K-row cap as _write_data_sheet / _build_engine.
                df_cap = df.head(100000) if len(df) > 100000 else df
                result_val = self._compute_kpi_numeric(df_cap, kpi)
                ws.merge_range(val_row, col, val_row, col + span_cols - 1, 0, val_fmt)
                ws.write_formula(val_row, col, formula, val_fmt, result_val)
            else:
                ws.merge_range(val_row, col, val_row, col + span_cols - 1,
                               self._compute_kpi_static(df, kpi), val_fmt)
        else:
            ws.merge_range(val_row, col, val_row, col + span_cols - 1, "N/A", val_fmt)

        # Sparkline row (fourth row of tile)
        if span_rows >= 4:
            spark_row = row + 3
            ws.set_row(spark_row, 18)
            for ci in range(col, col + span_cols):
                ws.write_blank(spark_row, ci, None, bg_fmt)
            self._add_kpi_sparkline(df, kpi, spark_row, col + span_cols // 2, bg)

        # Bottom padding row
        bot_row = row + span_rows - 1
        ws.merge_range(bot_row, col, bot_row, col + span_cols - 1, "", bg_fmt)

    def _add_kpi_sparkline(self, df: pd.DataFrame, kpi: KPIConfig,
                            row: int, col: int, bg: str) -> None:
        """Add sparkline mini-chart to a KPI tile row."""
        t = self.theme
        vals = build_kpi_sparkline_range(df, kpi.column)
        if not vals or not isinstance(vals, list) or len(vals) < 2:
            return
        spark_col_idx = 22 + list(self.config.kpis).index(kpi) if kpi in self.config.kpis else 25
        for si, v in enumerate(vals):
            try:
                self._ws_calc.write(si, spark_col_idx, float(v))
            except Exception:
                pass
        spark_range = (f"Calculations!${col_letter(spark_col_idx)}$1:"
                       f"${col_letter(spark_col_idx)}${len(vals)}")
        try:
            write_sparkline(self._ws_dash, row, col, spark_range,
                            color=_hex(t.accent2), sparkline_type="line")
        except Exception:
            pass

    def _write_section_header(self, row: int, text: str,
                               start_col: int = 0, end_col: int = N_COLS - 1,
                               color: str | None = None, height: int = 22,
                               ws=None) -> None:
        ws = ws or self._ws_dash
        t = self.theme
        bg = color or t.secondary
        fmt = self._wb.add_format({
            "font_name": t.font_heading, "font_size": 10, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(bg),
            "align": "left", "valign": "vcenter", "indent": 1,
        })
        ws.set_row(row, height)
        ws.merge_range(row, start_col, row, end_col, f"  {text}", fmt)

    def _write_summary_table(self, df: pd.DataFrame,
                              start_row: int, start_col: int = 0,
                              max_rows: int = 15, max_cols: int = 8,
                              ws=None) -> int:
        """Write formatted summary table. Returns number of rows consumed."""
        ws = ws or self._ws_dash
        sf = self._sf

        cols = [c for c in self.config.table_columns if c in df.columns][:max_cols]
        if not cols:
            cols = list(df.columns)[:max_cols]

        hdr_fmt = sf.table_header()
        ws.set_row(start_row, 18)
        for j, c in enumerate(cols):
            ws.write(start_row, start_col + j,
                     c.replace("_", " ").title(), hdr_fmt)

        display = df[cols].head(max_rows)
        for i, row in enumerate(display.itertuples(index=False), 1):
            stripe = i % 2 == 0
            ws.set_row(start_row + i, 15)
            for j, (cn, val) in enumerate(zip(cols, row)):
                fmt = sf.table_data_num(stripe) if isinstance(val, (int, float)) \
                    else sf.table_data(stripe)
                try:
                    ws.write(start_row + i, start_col + j, val, fmt)
                except Exception:
                    ws.write(start_row + i, start_col + j, str(val) if val else "", fmt)

        return len(display) + 2

    # ── KPI helpers (fallback static) ─────────────────────────────────────────

    def _compute_kpi_numeric(self, df: pd.DataFrame, kpi: KPIConfig) -> float:
        """Compute raw numeric KPI value for formula caching."""
        if kpi.column not in df.columns:
            return 0
        series = df[kpi.column].dropna()
        if series.empty:
            return 0
        try:
            agg = kpi.aggregation
            if agg == AggFunc.SUM:
                return float(series.sum())
            elif agg == AggFunc.AVG:
                return float(series.mean())
            elif agg == AggFunc.COUNT:
                return float(len(series))
            elif agg == AggFunc.MAX:
                return float(series.max())
            elif agg == AggFunc.MIN:
                return float(series.min())
            elif agg == AggFunc.MEDIAN:
                return float(series.median())
            elif agg == AggFunc.DISTINCT:
                return float(series.nunique())
            else:
                return float(series.sum())
        except Exception:
            return 0

    def _compute_kpi_static(self, df: pd.DataFrame, kpi: KPIConfig) -> str:
        if kpi.column not in df.columns:
            return "N/A"
        series = df[kpi.column].dropna()
        if series.empty:
            return "N/A"
        try:
            agg = kpi.aggregation
            if agg == AggFunc.SUM:
                val = float(series.sum())
            elif agg == AggFunc.AVG:
                val = float(series.mean())
            elif agg == AggFunc.COUNT:
                val = float(len(series))
            elif agg == AggFunc.MAX:
                val = float(series.max())
            elif agg == AggFunc.MIN:
                val = float(series.min())
            elif agg == AggFunc.MEDIAN:
                val = float(series.median())
            elif agg == AggFunc.DISTINCT:
                val = float(series.nunique())
            else:
                val = float(series.sum())
        except Exception:
            return "N/A"
        return self._fmt_val(val, kpi)

    def _fmt_val(self, val: float, kpi: KPIConfig) -> str:
        p, s, fmt = kpi.prefix, kpi.suffix, kpi.format

        # Avoid double symbols: strip currency/percent symbols from prefix/suffix
        # since the format handler adds them. Keeps other prefix text intact.
        if fmt == NumberFormat.CURRENCY:
            for sym in ("$", "£", "€", "¥"):
                p = p.replace(sym, "")
        if fmt == NumberFormat.PERCENTAGE:
            s = s.replace("%", "")

        if fmt == NumberFormat.CURRENCY:
            if abs(val) >= 1_000_000_000:
                return f"{p}${val/1_000_000_000:.1f}B{s}"
            elif abs(val) >= 1_000_000:
                return f"{p}${val/1_000_000:.1f}M{s}"
            elif abs(val) >= 1_000:
                return f"{p}${val/1_000:.1f}K{s}"
            return f"{p}${val:,.2f}{s}"
        elif fmt == NumberFormat.PERCENTAGE:
            return f"{p}{val:.1f}%{s}"
        elif fmt == NumberFormat.NUMBER:
            if abs(val) >= 1_000_000_000:
                return f"{p}{val/1_000_000_000:.1f}B{s}"
            elif abs(val) >= 1_000_000:
                return f"{p}{val/1_000_000:.1f}M{s}"
            elif abs(val) >= 1_000:
                return f"{p}{val/1_000:.1f}K{s}"
            return f"{p}{val:,.0f}{s}"
        elif fmt == NumberFormat.DECIMAL:
            return f"{p}{val:,.2f}{s}"
        elif fmt == NumberFormat.INTEGER:
            return f"{p}{int(val):,}{s}"
        return f"{p}{val:,.1f}{s}"

    # ── Deep Analysis storage ─────────────────────────────────────────────────

    def _store_for_analysis(self, df: pd.DataFrame, profile: DatasetProfile) -> None:
        """Store df and profile so _close() can write the Deep Analysis sheet."""
        self._analysis_df = df
        self._analysis_profile = profile

    # ── Deep Analysis sheet rendering ─────────────────────────────────────────

    def _write_deep_analysis_sheet(self) -> None:
        """Render the full Deep Analysis sheet. Called from _close()."""
        analysis = self.config.deep_analysis
        if analysis is None:
            return

        ws = self._ws_analysis
        sf = self._sf
        t = self.theme

        # Column widths: A=3, B=4, C-N=10 (12 content cols), O=3
        DA_COLS = 15
        ws.set_column(0, 0, 3)
        ws.set_column(1, 1, 4)
        ws.set_column(2, DA_COLS - 2, 10)
        ws.set_column(DA_COLS - 1, DA_COLS - 1, 3)
        ws.hide_gridlines(2)

        row = 0
        row = self._da_write_title_bar(ws, sf, row, analysis, DA_COLS)
        row += 1  # spacer
        row = self._da_write_exec_summary(ws, sf, row, analysis, DA_COLS)
        row += 1
        row = self._da_write_key_findings(ws, sf, row, analysis, DA_COLS)
        row += 1
        row = self._da_write_data_quality(ws, sf, row, analysis, DA_COLS)
        row += 1
        row = self._da_write_distribution(ws, sf, row, analysis, DA_COLS)
        row += 1
        row = self._da_write_correlations(ws, sf, row, analysis, DA_COLS)
        row += 1
        if analysis.trend_insights or analysis.trend_summary:
            row = self._da_write_trends(ws, sf, row, analysis, DA_COLS)
            row += 1
        row = self._da_write_performers(ws, sf, row, analysis, DA_COLS)
        row += 1
        row = self._da_write_outliers(ws, sf, row, analysis, DA_COLS)
        row += 1
        row = self._da_write_outlook(ws, sf, row, analysis, DA_COLS)
        row += 1
        row = self._da_write_industry(ws, sf, row, analysis, DA_COLS)
        row += 1
        self._da_write_footer(ws, sf, row, DA_COLS)

    # ── Section writers ───────────────────────────────────────────────────────

    def _da_write_title_bar(self, ws, sf, row: int, a, nc: int) -> int:
        from datetime import date
        ws.set_row(row, 42)
        ws.merge_range(row, 0, row, nc - 5,
                       f"  Deep Analysis — {self.config.title}",
                       sf.analysis_title())
        ws.merge_range(row, nc - 4, row, nc - 1,
                       f"{date.today().strftime('%B %d, %Y')}  ",
                       sf.analysis_meta_right())
        row += 1
        ws.set_row(row, 18)
        subtitle = self.config.subtitle or f"AI-powered analysis of {self.config.title}"
        ws.merge_range(row, 0, row, nc - 1, f"  {subtitle}", sf.analysis_subtitle())
        return row + 1

    def _da_write_exec_summary(self, ws, sf, row: int, a, nc: int) -> int:
        ws.set_row(row, 24)
        ws.merge_range(row, 0, row, nc - 1,
                       "  §1  EXECUTIVE SUMMARY", sf.analysis_section_header())
        row += 1
        if a.executive_summary:
            ws.set_row(row, 60)
            ws.merge_range(row, 0, row, nc - 1, a.executive_summary, sf.analysis_body())
            row += 1
        return row

    def _da_write_key_findings(self, ws, sf, row: int, a, nc: int) -> int:
        ws.set_row(row, 24)
        ws.merge_range(row, 0, row, nc - 1,
                       "  §2  KEY FINDINGS", sf.analysis_section_header())
        row += 1
        for i, finding in enumerate(a.key_findings[:6], 1):
            ws.set_row(row, 30)
            ws.write(row, 1, str(i), sf.analysis_bullet())
            ws.merge_range(row, 2, row, nc - 1, finding, sf.analysis_body())
            row += 1
        return row

    def _da_write_data_quality(self, ws, sf, row: int, a, nc: int) -> int:
        ws.set_row(row, 24)
        ws.merge_range(row, 0, row, nc - 1,
                       "  §3  DATA QUALITY ASSESSMENT", sf.analysis_section_header())
        row += 1
        # Score badge
        ws.set_row(row, 36)
        ws.write(row, 1, a.data_quality_score, sf.score_badge(a.data_quality_score))
        label = "Excellent" if a.data_quality_score >= 80 else \
                "Good" if a.data_quality_score >= 60 else \
                "Fair" if a.data_quality_score >= 40 else "Poor"
        ws.merge_range(row, 2, row, 5,
                       f"  Quality Score: {a.data_quality_score}/100 — {label}",
                       sf.analysis_subheader())
        ws.merge_range(row, 6, row, nc - 1, "", sf.analysis_body())
        row += 1
        for note in a.data_quality_notes[:4]:
            ws.set_row(row, 22)
            ws.write(row, 1, "•", sf.analysis_bullet())
            ws.merge_range(row, 2, row, nc - 1, note, sf.analysis_body())
            row += 1
        return row

    def _da_write_distribution(self, ws, sf, row: int, a, nc: int) -> int:
        ws.set_row(row, 24)
        ws.merge_range(row, 0, row, nc - 1,
                       "  §4  STATISTICAL DISTRIBUTION INSIGHTS",
                       sf.analysis_section_header())
        row += 1
        for insight in a.distribution_insights[:4]:
            ws.set_row(row, 30)
            ws.write(row, 1, "•", sf.analysis_bullet())
            ws.merge_range(row, 2, row, nc - 1, insight, sf.analysis_body())
            row += 1
        return row

    def _da_write_correlations(self, ws, sf, row: int, a, nc: int) -> int:
        ws.set_row(row, 24)
        ws.merge_range(row, 0, row, nc - 1,
                       "  §5  CORRELATION ANALYSIS", sf.analysis_section_header())
        row += 1
        if not a.correlation_insights:
            ws.set_row(row, 22)
            ws.merge_range(row, 0, row, nc - 1,
                           "  No significant correlations detected (|r| > 0.3).",
                           sf.analysis_body())
            return row + 1

        # Table header
        headers = ["Column A", "Column B", "Coefficient", "Interpretation"]
        col_starts = [1, 4, 7, 9]
        col_ends = [3, 6, 8, nc - 1]
        ws.set_row(row, 20)
        for hi, h in enumerate(headers):
            ws.merge_range(row, col_starts[hi], row, col_ends[hi], h,
                           sf.analysis_table_header())
        row += 1
        for ci in a.correlation_insights[:8]:
            stripe = (row % 2 == 0)
            ws.set_row(row, 20)
            ws.merge_range(row, 1, row, 3, ci.col_a, sf.analysis_table_cell(stripe))
            ws.merge_range(row, 4, row, 6, ci.col_b, sf.analysis_table_cell(stripe))
            ws.merge_range(row, 7, row, 8, ci.coefficient, sf.analysis_table_num(stripe))
            ws.merge_range(row, 9, row, nc - 1, ci.interpretation,
                           sf.analysis_table_cell(stripe))
            row += 1
        return row

    def _da_write_trends(self, ws, sf, row: int, a, nc: int) -> int:
        ws.set_row(row, 24)
        ws.merge_range(row, 0, row, nc - 1,
                       "  §6  TREND ANALYSIS", sf.analysis_section_header())
        row += 1
        if a.trend_summary:
            ws.set_row(row, 45)
            ws.merge_range(row, 0, row, nc - 1, a.trend_summary, sf.analysis_body())
            row += 1
        if a.trend_insights:
            headers = ["Metric", "Direction", "Change %", "Description"]
            col_starts = [1, 5, 7, 9]
            col_ends = [4, 6, 8, nc - 1]
            ws.set_row(row, 20)
            for hi, h in enumerate(headers):
                ws.merge_range(row, col_starts[hi], row, col_ends[hi], h,
                               sf.analysis_table_header())
            row += 1
            for ti in a.trend_insights[:8]:
                stripe = (row % 2 == 0)
                ws.set_row(row, 20)
                ws.merge_range(row, 1, row, 4, ti.column,
                               sf.analysis_table_cell(stripe))
                dir_fmt, symbol = sf.direction_badge(ti.direction)
                ws.merge_range(row, 5, row, 6, f"{symbol} {ti.direction.upper()}",
                               dir_fmt)
                ws.merge_range(row, 7, row, 8,
                               f"{ti.pct_change:+.1f}%",
                               sf.analysis_table_num(stripe))
                ws.merge_range(row, 9, row, nc - 1, ti.description,
                               sf.analysis_table_cell(stripe))
                row += 1
        return row

    def _da_write_performers(self, ws, sf, row: int, a, nc: int) -> int:
        ws.set_row(row, 24)
        ws.merge_range(row, 0, row, nc - 1,
                       "  §7  TOP & BOTTOM PERFORMERS", sf.analysis_section_header())
        row += 1

        if a.dimension_analysis:
            ws.set_row(row, 40)
            ws.merge_range(row, 0, row, nc - 1, a.dimension_analysis, sf.analysis_body())
            row += 1

        # Side-by-side: Top 5 (cols 1-6) | Bottom 5 (cols 8-13)
        # Top header
        ws.set_row(row, 20)
        ws.merge_range(row, 1, row, 3, "Top Performers", sf.analysis_subheader())
        ws.merge_range(row, 4, row, 6, "Value", sf.analysis_subheader())
        ws.merge_range(row, 8, row, 10, "Bottom Performers", sf.analysis_subheader())
        ws.merge_range(row, 11, row, nc - 2, "Value", sf.analysis_subheader())
        row += 1
        max_rows = max(len(a.top_performers), len(a.bottom_performers))
        for i in range(min(max_rows, 5)):
            stripe = (i % 2 == 0)
            ws.set_row(row, 18)
            if i < len(a.top_performers):
                p = a.top_performers[i]
                ws.merge_range(row, 1, row, 3, p.dimension_value,
                               sf.analysis_table_cell(stripe))
                ws.merge_range(row, 4, row, 6, p.metric_value,
                               sf.analysis_table_num(stripe))
            if i < len(a.bottom_performers):
                p = a.bottom_performers[i]
                ws.merge_range(row, 8, row, 10, p.dimension_value,
                               sf.analysis_table_cell(stripe))
                ws.merge_range(row, 11, row, nc - 2, p.metric_value,
                               sf.analysis_table_num(stripe))
            row += 1
        return row

    def _da_write_outliers(self, ws, sf, row: int, a, nc: int) -> int:
        ws.set_row(row, 24)
        ws.merge_range(row, 0, row, nc - 1,
                       "  §8  OUTLIER ANALYSIS", sf.analysis_section_header())
        row += 1
        if not a.outlier_insights:
            ws.set_row(row, 22)
            ws.merge_range(row, 0, row, nc - 1,
                           "  No significant outliers detected.",
                           sf.analysis_body())
            return row + 1

        headers = ["Column", "Count", "% of Data", "Description"]
        col_starts = [1, 4, 6, 8]
        col_ends = [3, 5, 7, nc - 1]
        ws.set_row(row, 20)
        for hi, h in enumerate(headers):
            ws.merge_range(row, col_starts[hi], row, col_ends[hi], h,
                           sf.analysis_table_header())
        row += 1
        for oi in a.outlier_insights[:8]:
            stripe = (row % 2 == 0)
            ws.set_row(row, 20)
            ws.merge_range(row, 1, row, 3, oi.column, sf.analysis_table_cell(stripe))
            ws.merge_range(row, 4, row, 5, oi.count, sf.analysis_table_num(stripe))
            ws.merge_range(row, 6, row, 7, f"{oi.pct:.1f}%",
                           sf.analysis_table_cell(stripe))
            ws.merge_range(row, 8, row, nc - 1, oi.description,
                           sf.analysis_table_cell(stripe))
            row += 1
        return row

    def _da_write_outlook(self, ws, sf, row: int, a, nc: int) -> int:
        ws.set_row(row, 24)
        ws.merge_range(row, 0, row, nc - 1,
                       "  §9  OUTLOOK & RECOMMENDATIONS", sf.analysis_section_header())
        row += 1

        if a.near_term_outlook:
            ws.set_row(row, 20)
            ws.merge_range(row, 0, row, nc - 1,
                           "  Near-Term Outlook", sf.analysis_subheader())
            row += 1
            ws.set_row(row, 40)
            ws.merge_range(row, 0, row, nc - 1, a.near_term_outlook, sf.analysis_body())
            row += 1

        if a.long_term_outlook:
            ws.set_row(row, 20)
            ws.merge_range(row, 0, row, nc - 1,
                           "  Long-Term Outlook", sf.analysis_subheader())
            row += 1
            ws.set_row(row, 40)
            ws.merge_range(row, 0, row, nc - 1, a.long_term_outlook, sf.analysis_body())
            row += 1

        if a.recommendations:
            ws.set_row(row, 20)
            ws.merge_range(row, 0, row, nc - 1,
                           "  Actionable Recommendations", sf.analysis_subheader())
            row += 1
            for i, rec in enumerate(a.recommendations[:6], 1):
                ws.set_row(row, 28)
                ws.write(row, 1, str(i), sf.analysis_bullet())
                ws.merge_range(row, 2, row, nc - 1, rec, sf.analysis_body())
                row += 1
        return row

    def _da_write_industry(self, ws, sf, row: int, a, nc: int) -> int:
        ws.set_row(row, 24)
        ws.merge_range(row, 0, row, nc - 1,
                       "  §10  INDUSTRY CONTEXT", sf.analysis_section_header())
        row += 1
        if a.industry_context:
            ws.set_row(row, 60)
            ws.merge_range(row, 0, row, nc - 1, a.industry_context, sf.analysis_body())
            row += 1
        return row

    def _da_write_footer(self, ws, sf, row: int, nc: int) -> None:
        from datetime import date
        ws.set_row(row, 22)
        ws.merge_range(row, 0, row, nc - 1,
                       f"Generated by Excel Master AI  |  {date.today().strftime('%Y-%m-%d')}  |  "
                       f"Analysis powered by LLM deep-stats pipeline",
                       sf.analysis_footer())

    # ── Close workbook ─────────────────────────────────────────────────────────

    def _close(self, output_path: Path) -> Path:
        # Write Deep Analysis sheet if available (runs before wb.close)
        try:
            self._write_deep_analysis_sheet()
        except Exception:
            pass  # analysis sheet is optional — never break the save
        try:
            self._wb.close()
        except Exception as e:
            raise RuntimeError(f"Failed to save workbook: {e}") from e
        return output_path
