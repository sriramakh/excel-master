"""
Financial Template — Forest green, P&L reporting style

Layout: Green accent, dense financial data, waterfall-like table
  Row 0:      Green title with period indicator
  Row 1-2:    3-filter panel (Account Type, Period, Entity)
  Rows 3-7:   4 Large KPI tiles (Revenue, Cost, EBITDA, Net Margin)
  Row 8:      Section header
  Rows 9-19:  Full-width area chart (revenue vs expense trend)
  Row 20:     Section header
  Rows 21-31: Bar (left 60%) | Pie (right 40%)
  Row 32:     Section header
  Rows 33+:   Financial table with color-coded +/- rows
"""
from __future__ import annotations
from pathlib import Path
import pandas as pd

from .base_xl_template import (
    BaseXLTemplate, CHART_HALF_W, CHART_HALF_H, CHART_THIRD_W,
    CHART_FULL_W, CHART_FULL_H, N_COLS, COL_W,
)
from ..xl_chart import ChartZone
from ..xl_style import _hex
from ...models import ChartType


class FinancialXL(BaseXLTemplate):
    """Financial P&L dashboard — forest green, area trend + bar/pie breakdown."""
    name = "financial"

    def build(self, df: pd.DataFrame, output_path: Path) -> Path:
        self._init_workbook(output_path)
        self._write_data_sheet(df)

        ws = self._ws_dash
        sf = self._sf
        t = self.theme

        ws.set_column(0, N_COLS - 1, COL_W)
        ws.hide_gridlines(2)

        # Light mint background
        light_bg_fmt = self._wb.add_format({"bg_color": _hex(t.bg_dashboard)})
        for r in range(70):
            for c in range(N_COLS):
                ws.write_blank(r, c, None, light_bg_fmt)

        # ── Row 0: Green title with period on right ────────────────────────────
        ws.set_row(0, 46)
        title_fmt = self._wb.add_format({
            "font_name": t.font_heading, "font_size": 20, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.primary),
            "align": "left", "valign": "vcenter", "indent": 2,
        })
        period_fmt = self._wb.add_format({
            "font_name": t.font_body, "font_size": 10, "bold": True,
            "font_color": _hex(t.accent1), "bg_color": _hex(t.primary),
            "align": "right", "valign": "vcenter", "indent": 2,
        })
        ws.merge_range(0, 0, 0, 15, f"  {self.config.title}", title_fmt)
        ws.merge_range(0, 16, 0, N_COLS - 1,
                       self.config.subtitle or "Financial Analytics", period_fmt)

        # ── Rows 1-2: Filter panel (3 dimensions) ─────────────────────────────
        ws.set_row(1, 30)
        ws.set_row(2, 14)
        filter_cols = self._detect_filter_cols(
            df, prefer_keywords=["account_type", "category", "department",
                                  "entity", "segment", "region"])
        filter_refs = self._write_filter_slicer_panel(
            df, row=1, filter_cols=filter_cols,
            panel_bg=t.bg_section if t.bg_section else t.bg_dashboard)

        # ── Rows 3-7: 4 KPI tiles ─────────────────────────────────────────────
        ws.set_row(3, 8)
        ws.set_row(4, 14)
        ws.set_row(5, 32)
        ws.set_row(6, 16)
        ws.set_row(7, 8)

        kpis = self.config.kpis[:4]
        kpi_colors = [t.primary, t.secondary, t.accent2, t.accent1]
        tile_span = N_COLS // max(len(kpis), 1)

        for i, kpi in enumerate(kpis):
            c = i * tile_span
            bg = kpi_colors[i % len(kpi_colors)]
            self._write_kpi_tile(3, c, tile_span, 5, kpi, df, bg,
                                  font_color=t.text_light,
                                  filter_ref=list(filter_refs.values())[0]
                                  if filter_refs else None)

        # ── Row 8: Section header ─────────────────────────────────────────────
        self._write_section_header(8, "💰  P&L Trend Analysis", color=t.secondary)

        # ── Rows 9-21: Full-width area/line (H=250 → ~13 rows) ──────────────
        self._build_engine(df, filter_refs)
        charts = self.config.charts

        area_cfg = (self._pick(charts, ChartType.AREA)
                    or self._pick(charts, ChartType.LINE))
        if area_cfg:
            zone = ChartZone(9, 0, CHART_FULL_W, CHART_FULL_H)
            self._add_chart(area_cfg, df, zone)

        # ── Row 23: Section header (row 9 + 14) ─────────────────────────────
        self._write_section_header(23, "📊  Category Breakdown", color=t.accent2)

        # ── Rows 24-37: Bar (left) | Pie (right) (H=280 → ~14 rows) ────────
        bar_cfg = self._pick(charts, ChartType.BAR)
        pie_cfg = (self._pick(charts, ChartType.PIE)
                   or self._pick(charts, ChartType.DOUGHNUT))

        next_row = 24
        if bar_cfg:
            zone = ChartZone(next_row, 0, 760, CHART_HALF_H)
            self._add_chart(bar_cfg, df, zone)
        if pie_cfg:
            zone = ChartZone(next_row, 14, 520, CHART_HALF_H)
            self._add_chart(pie_cfg, df, zone)
        next_row += 15

        # Remaining charts in pairs
        placed = [area_cfg, bar_cfg, pie_cfg]
        remaining = [c for c in charts if c not in placed]
        for ri in range(0, len(remaining), 2):
            zone = ChartZone(next_row, 0, CHART_HALF_W, CHART_HALF_H)
            self._add_chart(remaining[ri], df, zone)
            if ri + 1 < len(remaining):
                zone = ChartZone(next_row, 12, CHART_HALF_W, CHART_HALF_H)
                self._add_chart(remaining[ri + 1], df, zone)
            next_row += 15

        # ── Table section ────────────────────────────────────────────────────
        self._write_section_header(next_row, "📋  Financial Data", color=t.bg_table_header)
        self._write_financial_table(df, next_row + 1)

        ws.freeze_panes(3, 0)
        ws.set_zoom(85)
        return self._close(output_path)

    def _write_financial_table(self, df: pd.DataFrame, start_row: int) -> None:
        """Write table with currency formatting and +/- conditional colors."""
        ws = self._ws_dash
        sf = self._sf
        t = self.theme

        cols = [c for c in self.config.table_columns if c in df.columns][:8]
        if not cols:
            cols = list(df.columns)[:8]

        curr_kw = ["usd", "spend", "budget", "revenue", "cost", "value", "amount", "price"]
        pct_kw = ["pct", "rate", "percent", "margin", "roi"]
        num_cols_list = [c for c in cols if df[c].dtype.kind in "iuf"]

        hdr_fmt = sf.table_header()
        ws.set_row(start_row, 18)
        for j, c in enumerate(cols):
            ws.write(start_row, j, c.replace("_", " ").title(), hdr_fmt)

        display = df[cols].head(15)
        for i, row in enumerate(display.itertuples(index=False), 1):
            stripe = i % 2 == 0
            ws.set_row(start_row + i, 15)
            for j, (cn, val) in enumerate(zip(cols, row)):
                cn_l = cn.lower()
                if any(k in cn_l for k in pct_kw):
                    fmt = sf.table_pct(stripe)
                elif any(k in cn_l for k in curr_kw):
                    fmt = sf.table_currency(stripe)
                elif isinstance(val, (int, float)):
                    fmt = sf.table_data_num(stripe)
                else:
                    fmt = sf.table_data(stripe)
                try:
                    ws.write(start_row + i, j, val, fmt)
                except Exception:
                    ws.write(start_row + i, j, str(val) if val else "", fmt)

        # 3-color scale on numeric columns
        for j, col in enumerate(cols):
            if col in num_cols_list:
                ws.conditional_format(start_row + 1, j, start_row + 15, j, {
                    "type": "3_color_scale",
                    "min_color": _hex(t.negative),
                    "mid_color": "#FFFFAA",
                    "max_color": _hex(t.positive),
                })

    def _pick(self, charts, ctype: ChartType):
        return next((c for c in charts if c.type == ctype), None)
