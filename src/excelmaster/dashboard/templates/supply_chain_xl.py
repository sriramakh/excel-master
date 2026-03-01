"""
Supply Chain Template — inspired by Supply-Chain.png

Layout: Warm bg, LEFT SIDEBAR for KPIs, TOP month buttons, central charts
  Row 0:      Title bar (minimal, warm)
  Row 1:      Month/Period selector buttons (12 horizontal cells) + gear icon
  Row 2:      Section divider
  Rows 3-24:  TWO-COLUMN layout:
              LEFT (cols 0-5): Stacked KPI metrics + operational data
              RIGHT (cols 6-23): Large line chart, then 2-col breakdown
  Row 25:     Section header — Shipment Data
  Rows 26+:   Operational table with data bars
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


class SupplyChainXL(BaseXLTemplate):
    """Supply chain dashboard — warm bg, left sidebar KPIs, period selectors."""
    name = "supply_chain"

    def build(self, df: pd.DataFrame, output_path: Path) -> Path:
        self._init_workbook(output_path)
        self._write_data_sheet(df)

        ws = self._ws_dash
        sf = self._sf
        t = self.theme

        ws.set_column(0, N_COLS - 1, COL_W)
        ws.hide_gridlines(2)

        # Warm beige/tan background fill
        warm_bg = _hex(t.bg_dashboard)
        bg_fmt = self._wb.add_format({"bg_color": warm_bg})
        for r in range(70):
            for c in range(N_COLS):
                ws.write_blank(r, c, None, bg_fmt)

        # ── Row 0: Clean title ─────────────────────────────────────────────────
        ws.set_row(0, 40)
        title_fmt = self._wb.add_format({
            "font_name": t.font_heading, "font_size": 16, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.primary),
            "align": "left", "valign": "vcenter", "indent": 2,
        })
        ws.merge_range(0, 0, 0, 5, f"  {self.config.title}", title_fmt)
        # Logo area (right side of title)
        logo_fmt = self._wb.add_format({
            "font_name": t.font_body, "font_size": 9,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.secondary),
            "align": "center", "valign": "vcenter",
        })
        ws.merge_range(0, 6, 0, N_COLS - 1,
                       self.config.subtitle or "Freight & Logistics Analytics", logo_fmt)

        # ── Row 1: Period selector buttons (month/quarter) ────────────────────
        ws.set_row(1, 28)
        filter_cols = self._detect_filter_cols(
            df, prefer_keywords=["carrier", "supplier", "warehouse",
                                  "route", "origin", "mode", "region"])
        filter_refs = self._write_filter_slicer_panel(
            df, row=1, filter_cols=filter_cols, panel_bg=t.bg_dashboard)

        # Also write period selector if time column exists
        time_col = self._find_time_col(df)
        if time_col and time_col in df.columns:
            self._write_period_buttons(df, time_col, row=1)

        # ── Rows 2-3: Divider ─────────────────────────────────────────────────
        ws.set_row(2, 14)
        ws.set_row(3, 4)
        div_fmt = self._wb.add_format({
            "bg_color": _hex(t.secondary),
            "font_color": _hex(t.text_light),
            "font_name": t.font_body, "font_size": 8,
            "align": "left", "valign": "vcenter", "indent": 2,
        })
        ws.merge_range(2, 0, 2, N_COLS - 1, "  Click filter cells above to filter dashboard", div_fmt)

        # ── LEFT SIDEBAR (cols 0-5): Stacked KPI metrics ──────────────────────
        sidebar_col = 0
        sidebar_width = 6

        kpis = self.config.kpis[:6]
        kpi_colors = [t.primary, t.accent1, t.accent2, t.secondary, t.accent3, t.neutral]

        sidebar_header_fmt = self._wb.add_format({
            "font_name": t.font_heading, "font_size": 9, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.primary),
            "align": "left", "valign": "vcenter", "indent": 1,
        })
        ws.set_row(4, 20)
        ws.merge_range(4, sidebar_col, 4, sidebar_col + sidebar_width - 1,
                       "  📊 KEY METRICS", sidebar_header_fmt)

        for i, kpi in enumerate(kpis):
            row_start = 5 + i * 4
            bg = kpi_colors[i % len(kpi_colors)]
            self._write_kpi_tile(row_start, sidebar_col, sidebar_width, 4,
                                  kpi, df, bg, font_color=t.text_light,
                                  filter_ref=list(filter_refs.values())[0]
                                  if filter_refs else None)

        # ── RIGHT CONTENT (cols 6-23): Charts ─────────────────────────────────
        right_start_col = sidebar_width
        right_width = N_COLS - sidebar_width  # 18 cols

        self._build_engine(df, filter_refs)
        charts = self.config.charts

        # Large line chart at top of right panel (H=250 → ~13 rows)
        line_cfg = self._pick(charts, ChartType.LINE)
        if line_cfg:
            zone = ChartZone(4, right_start_col, 1100, CHART_FULL_H)
            self._add_chart(line_cfg, df, zone)

        # Two charts below (row 4 + 14 = row 18)
        bar_cfg = (self._pick(charts, ChartType.BAR)
                   or self._pick(charts, ChartType.BAR_HORIZONTAL))
        pie_cfg = (self._pick(charts, ChartType.PIE)
                   or self._pick(charts, ChartType.DOUGHNUT))
        scatter_cfg = self._pick(charts, ChartType.SCATTER)

        next_row = 18
        if bar_cfg:
            zone = ChartZone(next_row, right_start_col, 540, CHART_HALF_H)
            self._add_chart(bar_cfg, df, zone)
        second = pie_cfg or scatter_cfg
        if second:
            zone = ChartZone(next_row, right_start_col + 9, 540, CHART_HALF_H)
            self._add_chart(second, df, zone)
        next_row += 15

        # Remaining charts in pairs (full-width)
        placed = [line_cfg, bar_cfg, second]
        remaining = [c for c in charts if c not in placed]
        for ri in range(0, len(remaining), 2):
            zone = ChartZone(next_row, 0, CHART_HALF_W, CHART_HALF_H)
            self._add_chart(remaining[ri], df, zone)
            if ri + 1 < len(remaining):
                zone = ChartZone(next_row, 12, CHART_HALF_W, CHART_HALF_H)
                self._add_chart(remaining[ri + 1], df, zone)
            next_row += 15

        # ── Data table ───────────────────────────────────────────────────────
        self._write_section_header(next_row, "📋  Shipment Data", color=t.primary)
        self._write_ops_table(df, next_row + 1)

        ws.freeze_panes(2, 0)
        ws.set_zoom(85)
        return self._close(output_path)

    def _write_period_buttons(self, df: pd.DataFrame, time_col: str, row: int) -> None:
        """Write month/period selector buttons on the right side of filter row."""
        ws = self._ws_dash
        t = self.theme
        # Display last 12 unique periods as button-style cells
        try:
            periods = sorted(df[time_col].dropna().astype(str).unique().tolist())[-12:]
        except Exception:
            return

        btn_col = N_COLS - len(periods) - 1
        for i, p in enumerate(periods):
            col_idx = btn_col + i
            if col_idx >= N_COLS:
                break
            btn_fmt = self._wb.add_format({
                "font_name": t.font_body, "font_size": 8, "bold": True,
                "font_color": _hex(t.text_light),
                "bg_color": _hex(t.accent2 if i == len(periods) - 1 else t.secondary),
                "align": "center", "valign": "vcenter",
                "border": 1, "border_color": _hex(t.bg_card),
            })
            ws.write(row, col_idx, str(p)[:7], btn_fmt)

    def _write_ops_table(self, df: pd.DataFrame, start_row: int) -> None:
        """Operational table with data bars on numeric columns."""
        ws = self._ws_dash
        sf = self._sf
        t = self.theme

        cols = [c for c in self.config.table_columns if c in df.columns][:8]
        if not cols:
            cols = list(df.columns)[:8]

        hdr_fmt = sf.table_header()
        ws.set_row(start_row, 18)
        for j, c in enumerate(cols):
            ws.write(start_row, j, c.replace("_", " ").title(), hdr_fmt)

        display = df[cols].head(15)
        num_cols_list = [c for c in cols if df[c].dtype.kind in "iuf"]

        for i, row in enumerate(display.itertuples(index=False), 1):
            stripe = i % 2 == 0
            ws.set_row(start_row + i, 15)
            for j, (cn, val) in enumerate(zip(cols, row)):
                fmt = sf.table_data_num(stripe) if cn in num_cols_list \
                    else sf.table_data(stripe)
                try:
                    ws.write(start_row + i, j, val, fmt)
                except Exception:
                    ws.write(start_row + i, j, str(val) if val else "", fmt)

        colors = [_hex(c) for c in t.chart_colors]
        for idx, col in enumerate(cols):
            if col in num_cols_list:
                ws.conditional_format(start_row + 1, idx, start_row + 15, idx, {
                    "type": "data_bar",
                    "bar_color": colors[idx % len(colors)],
                    "bar_solid": True,
                })

    def _find_time_col(self, df: pd.DataFrame) -> str | None:
        if self.config.time_column and self.config.time_column in df.columns:
            return self.config.time_column
        for c in df.columns:
            if any(k in c.lower() for k in ["month", "period", "date", "week", "quarter"]):
                return c
        return None

    def _pick(self, charts, ctype: ChartType):
        return next((c for c in charts if c.type == ctype), None)
