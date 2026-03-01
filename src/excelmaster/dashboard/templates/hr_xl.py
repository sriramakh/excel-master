"""
HR Analytics Template — inspired by Employee-Turnover.png

Layout: Dark purple background, neon accents, 2×3 chart grid
  Row 0:      Dark title (tall, neon text)
  Row 1:      Department filter buttons (circular-style cells in a row)
  Row 2:      Sub-filter + large KPI badge (top-right)
  Row 3-6:    KPI strip (6 compact tiles, 2 rows × 3 per row)
  Row 7:      Section header
  Rows 8-20:  TOP chart row: Bar (left) | Donut (center) | Line (right) — 3 equal thirds
  Row 21:     Section header
  Rows 22-33: BOTTOM chart row: Dual-line (left) | Scatter (center) | Bar-H (right)
  Row 34:     Section header
  Rows 35+:   Dark-styled table
"""
from __future__ import annotations
from pathlib import Path
import pandas as pd

from .base_xl_template import (
    BaseXLTemplate, CHART_THIRD_W, CHART_HALF_W, CHART_HALF_H,
    CHART_FULL_W, CHART_FULL_H, N_COLS, COL_W,
)
from ..xl_chart import ChartZone
from ..xl_style import _hex
from ...models import ChartType


class HRAnalyticsXL(BaseXLTemplate):
    """HR analytics — dark purple, neon accents, 2×3 chart grid layout."""
    name = "hr_analytics"

    def build(self, df: pd.DataFrame, output_path: Path) -> Path:
        self._init_workbook(output_path)
        self._write_data_sheet(df)

        ws = self._ws_dash
        sf = self._sf
        t = self.theme

        ws.set_column(0, N_COLS - 1, COL_W)
        ws.hide_gridlines(2)

        # Flood fill dark background
        dark_bg_fmt = self._wb.add_format({"bg_color": _hex(t.bg_dashboard)})
        for r in range(90):
            for c in range(N_COLS):
                ws.write_blank(r, c, None, dark_bg_fmt)

        # ── Row 0: Tall dark title ─────────────────────────────────────────────
        ws.set_row(0, 56)
        title_fmt = self._wb.add_format({
            "font_name": t.font_heading, "font_size": 24, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.primary),
            "align": "left", "valign": "vcenter", "indent": 2,
        })
        badge_fmt = self._wb.add_format({
            "font_name": t.font_heading, "font_size": 11, "bold": True,
            "font_color": _hex(t.accent1), "bg_color": _hex(t.primary),
            "align": "right", "valign": "vcenter", "indent": 1,
        })
        ws.merge_range(0, 0, 0, 15, f"  {self.config.title}", title_fmt)
        ws.merge_range(0, 16, 0, N_COLS - 1,
                       self.config.subtitle or "People Analytics", badge_fmt)

        # ── Row 1-2: Prominent department filter panel ─────────────────────────
        ws.set_row(1, 32)
        ws.set_row(2, 14)
        filter_cols = self._detect_filter_cols(
            df, prefer_keywords=["department", "dept", "division", "level", "location"])
        filter_refs = self._write_filter_slicer_panel(
            df, row=1, filter_cols=filter_cols,
            panel_bg=t.bg_card if not t.dark_mode else t.bg_dashboard)

        # ── Rows 3-8: Two rows of 3 KPI tiles ─────────────────────────────────
        kpis = self.config.kpis[:6]
        tile_span = 8  # each tile = 8 cols (3 tiles per row)
        card_colors = [t.accent1, t.accent2, t.accent3,
                       t.secondary, t.primary, t.neutral]
        # tile_rows: first group rows 3-6, second group rows 7-10 (no shared rows)
        tile_rows = [3, 7]

        for tile_row_idx, row_start in enumerate(tile_rows):
            ws.set_row(row_start, 8)
            ws.set_row(row_start + 1, 14)
            ws.set_row(row_start + 2, 30)
            ws.set_row(row_start + 3, 8)
            for ti in range(3):
                kpi_idx = tile_row_idx * 3 + ti
                if kpi_idx >= len(kpis):
                    break
                kpi = kpis[kpi_idx]
                c = ti * tile_span
                bg = card_colors[kpi_idx % len(card_colors)]
                self._write_kpi_tile(row_start, c, tile_span, 4, kpi, df,
                                      bg, font_color=t.text_light,
                                      filter_ref=list(filter_refs.values())[0]
                                      if filter_refs else None)

        # ── Row 11: Section header ────────────────────────────────────────────
        self._write_section_header(11, "👥  People Analytics — Top Row",
                                   color=t.secondary)

        # ── Rows 12-25: TOP row — 3 charts (H=280 → ~14 rows) ──────────────
        self._build_engine(df, filter_refs)
        charts = self.config.charts

        bar_cfg = self._pick(charts, ChartType.BAR)
        donut_cfg = (self._pick(charts, ChartType.DOUGHNUT)
                     or self._pick(charts, ChartType.PIE))
        line_cfg = self._pick(charts, ChartType.LINE)
        hr_charts_top = [c for c in (bar_cfg, donut_cfg, line_cfg) if c]

        for i, cfg in enumerate(hr_charts_top[:3]):
            zone = ChartZone(12, i * 8, CHART_THIRD_W, CHART_HALF_H)
            self._add_chart(cfg, df, zone)

        # ── Row 27: Section header (row 12 + 15) ────────────────────────────
        self._write_section_header(27, "📊  Distribution & Comparison",
                                   color=t.accent1)

        # ── Rows 28-41: BOTTOM row — remaining charts (H=280 → ~14 rows) ───
        remaining = [c for c in charts if c not in hr_charts_top]
        for i, cfg in enumerate(remaining[:2]):
            zone = ChartZone(28, i * 12, CHART_HALF_W, CHART_HALF_H)
            self._add_chart(cfg, df, zone)

        # ── Row 43: Table section (row 28 + 15) ─────────────────────────────
        self._write_section_header(43, "📋  Employee Data",
                                   color=t.bg_table_header)
        self._write_summary_table(df, 44, max_rows=15, max_cols=8)

        ws.freeze_panes(3, 0)
        ws.set_zoom(80)
        return self._close(output_path)

    def _pick(self, charts, ctype: ChartType):
        return next((c for c in charts if c.type == ctype), None)
