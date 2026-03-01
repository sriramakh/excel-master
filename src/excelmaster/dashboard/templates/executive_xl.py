"""
Executive Summary Template — inspired by AdminDash.png

Layout: White background, multi-color, complex multi-section
  Row 0:      Title (dark navy, left-aligned)
  Row 1:      Year/Quarter selector buttons (5 button-style cells top-right)
              + prominent filter dropdown (top-left)
  Row 2-3:    5 KPI cards with trend indicator and sparkline
  Row 4:      Section header "Analytics Overview"
  Row 5-19:   Chart zone A — Bar chart (left 14 cols) | Pie/Donut (right 10 cols)
  Row 20:     Section header "Trend Analysis"
  Row 21-32:  Full-width line chart
  Row 33:     Section header "Data Summary"
  Row 34+:    Summary table with color scale
"""
from __future__ import annotations
from pathlib import Path
import pandas as pd

from .base_xl_template import (
    BaseXLTemplate, CHART_HALF_W, CHART_HALF_H,
    CHART_FULL_W, CHART_FULL_H, CHART_THIRD_W, N_COLS, COL_W,
)
from ..xl_chart import ChartZone
from ..xl_style import _hex
from ...models import ChartType


class ExecutiveSummaryXL(BaseXLTemplate):
    """Corporate executive summary — white bg, 5 KPIs, bar+pie+trend layout."""
    name = "executive_summary"

    def build(self, df: pd.DataFrame, output_path: Path) -> Path:
        self._init_workbook(output_path)
        self._write_data_sheet(df)

        ws = self._ws_dash
        sf = self._sf
        t = self.theme

        # ── Column widths ──────────────────────────────────────────────────────
        for c in range(N_COLS):
            ws.set_column(c, c, COL_W)
        ws.hide_gridlines(2)

        # ── Row 0: Title ──────────────────────────────────────────────────────
        ws.set_row(0, 48)
        title_fmt = self._wb.add_format({
            "font_name": t.font_heading, "font_size": 22, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.primary),
            "align": "left", "valign": "vcenter", "indent": 2,
        })
        subtitle_fmt = self._wb.add_format({
            "font_name": t.font_body, "font_size": 10,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.secondary),
            "align": "right", "valign": "vcenter", "indent": 1,
        })
        ws.merge_range(0, 0, 0, 14, f"  {self.config.title}", title_fmt)
        ws.merge_range(0, 15, 0, N_COLS - 1,
                       self.config.subtitle or "Executive Dashboard", subtitle_fmt)

        # ── Row 1: Filter panel + period selectors ─────────────────────────────
        ws.set_row(1, 30)
        ws.set_row(2, 12)  # hint sub-row

        filter_cols = self._detect_filter_cols(df)
        filter_refs = self._write_filter_slicer_panel(df, row=1, filter_cols=filter_cols)

        # ── Rows 3-7: KPI Cards ───────────────────────────────────────────────
        kpi_top = 3
        kpis = self.config.kpis[:5]
        card_span = max(4, N_COLS // max(len(kpis), 1))
        card_colors = [t.primary, t.secondary, t.accent1, t.accent2, t.accent3]
        ws.set_row(kpi_top, 8)
        ws.set_row(kpi_top + 4, 8)

        for i, kpi in enumerate(kpis):
            c = i * card_span
            if c + card_span > N_COLS:
                break
            bg = card_colors[i % len(card_colors)]
            self._write_kpi_tile(kpi_top, c, card_span, 5, kpi, df, bg,
                                  font_color=t.text_light,
                                  filter_ref=list(filter_refs.values())[0] if filter_refs else None)

        # ── Row 8: Section header ─────────────────────────────────────────────
        self._write_section_header(8, "📊  Analytics Overview", color=t.secondary)

        # ── Rows 9-22: Bar (left) + Pie (right) — H=280 spans ~14 rows ─────
        self._build_engine(df, filter_refs)
        charts = self.config.charts

        bar_cfg = self._pick(charts, ChartType.BAR) or charts[0] if charts else None
        pie_cfg = (self._pick(charts, ChartType.PIE)
                   or self._pick(charts, ChartType.DOUGHNUT))
        line_cfg = self._pick(charts, ChartType.LINE)

        if bar_cfg:
            zone = ChartZone(9, 0, 760, CHART_HALF_H)
            self._add_chart(bar_cfg, df, zone)
        if pie_cfg:
            zone = ChartZone(9, 14, 520, CHART_HALF_H)
            self._add_chart(pie_cfg, df, zone)

        # ── Row 24: Section header (row 9 + 15) ─────────────────────────────
        self._write_section_header(24, "📈  Trend Analysis", color=t.accent1)

        # ── Rows 25-37: Full-width line — H=250 spans ~13 rows ──────────────
        if line_cfg:
            zone = ChartZone(25, 0, CHART_FULL_W, CHART_FULL_H)
            self._add_chart(line_cfg, df, zone)

        # ── Any remaining charts ─────────────────────────────────────────────
        remaining = [c for c in charts if c not in (bar_cfg, pie_cfg, line_cfg)]
        next_row = 38  # row 25 + 13
        for extra in remaining[:2]:
            zone = ChartZone(next_row, 0, CHART_HALF_W, CHART_HALF_H)
            self._add_chart(extra, df, zone)
            next_row += 15

        # ── Data table ───────────────────────────────────────────────────────
        tbl_row = next_row
        self._write_section_header(tbl_row, "📋  Data Summary", color=t.bg_table_header)
        self._write_summary_table(df, tbl_row + 1, max_rows=15, max_cols=8)

        # Color scale on numeric columns
        num_cols = [c for c in (self.config.table_columns or list(df.columns))[:8]
                    if c in df.columns and df[c].dtype.kind in "iuf"]
        for j, nc in enumerate(num_cols[:2]):
            ws.conditional_format(tbl_row + 2, j, tbl_row + 16, j, {
                "type": "3_color_scale",
                "min_color": _hex(t.negative),
                "mid_color": "#FFFFAA",
                "max_color": _hex(t.positive),
            })

        ws.freeze_panes(3, 0)
        ws.set_zoom(85)
        return self._close(output_path)

    def _pick(self, charts, ctype: ChartType):
        return next((c for c in charts if c.type == ctype), None)
