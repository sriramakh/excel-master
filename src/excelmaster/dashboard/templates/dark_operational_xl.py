"""
Dark Operational Template — inspired by HR-Metrics.png

Layout: Charcoal background, LEFT SIDEBAR filter list, right-side main content
  Row 0:      Dark title (blue rounded-rect style)
  Row 1:      5-number KPI strip (large numbers with icons)
  Row 2:      Colored mini card grid section header
  Rows 3-12:  LEFT (cols 0-4): Vertical filter list (scrollable dept selector)
              RIGHT (cols 5-23): 3×2 chart grid (6 charts)
              + Mini card grid (12 colored tiles) at top-right
  Row 13:     Full-width section header
  Rows 14+:   Dense data table (small rows, 20+ records)
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


class DarkOperationalXL(BaseXLTemplate):
    """Dark operational dashboard — charcoal, sidebar filter, colored card grid."""
    name = "dark_operational"

    def build(self, df: pd.DataFrame, output_path: Path) -> Path:
        self._init_workbook(output_path)
        self._write_data_sheet(df)

        ws = self._ws_dash
        sf = self._sf
        t = self.theme

        ws.set_column(0, N_COLS - 1, COL_W)
        ws.hide_gridlines(2)

        # Flood fill near-black background
        dark_bg_fmt = self._wb.add_format({"bg_color": _hex(t.bg_dashboard)})
        for r in range(80):
            for c in range(N_COLS):
                ws.write_blank(r, c, None, dark_bg_fmt)

        # ── Row 0: Dark title bar ─────────────────────────────────────────────
        ws.set_row(0, 52)
        title_fmt = self._wb.add_format({
            "font_name": t.font_heading, "font_size": 20, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.accent1),
            "align": "left", "valign": "vcenter", "indent": 2,
        })
        ws.merge_range(0, 0, 0, N_COLS - 1, f"  {self.config.title}", title_fmt)

        # ── Row 1: 5-number KPI strip ─────────────────────────────────────────
        ws.set_row(1, 14)
        ws.set_row(2, 36)
        ws.set_row(3, 12)

        kpis = self.config.kpis[:5]
        kpi_span = max(4, N_COLS // max(len(kpis), 1))
        icon_fmt = self._wb.add_format({
            "font_name": t.font_body, "font_size": 8,
            "font_color": _hex(t.text_muted), "bg_color": _hex(t.bg_card),
            "align": "center", "valign": "vcenter",
        })
        num_fmt = self._wb.add_format({
            "font_name": t.font_heading, "font_size": 22, "bold": True,
            "font_color": _hex(t.accent2), "bg_color": _hex(t.bg_card),
            "align": "center", "valign": "vcenter",
        })
        bg_card_fmt = self._wb.add_format({"bg_color": _hex(t.bg_card)})

        filter_cols = self._detect_filter_cols(df)
        self._build_engine(df, {})

        for i, kpi in enumerate(kpis):
            c = i * kpi_span
            ws.merge_range(1, c, 1, c + kpi_span - 1,
                           (kpi.icon + "  " + kpi.label if kpi.icon else kpi.label),
                           icon_fmt)
            ws.merge_range(2, c, 2, c + kpi_span - 1,
                           self._compute_kpi_static(df, kpi), num_fmt)
            ws.merge_range(3, c, 3, c + kpi_span - 1, "", bg_card_fmt)

        # ── Row 4: Filter panel ────────────────────────────────────────────────
        filter_refs = self._write_filter_slicer_panel(
            df, row=4, filter_cols=filter_cols[:3],
            panel_bg=t.bg_dashboard)

        # ── Row 6: Colored mini-card grid (top 12 categories) ─────────────────
        self._write_mini_card_grid(df, start_row=6, filter_cols=filter_cols)

        # ── Row 10: Section header ─────────────────────────────────────────────
        self._write_section_header(10, "⚡  Operational Metrics", color=t.accent1)

        # ── Rows 11-24: 3 charts in thirds (H=280 → ~14 rows) ──────────────
        charts = self.config.charts
        chart_iter = iter(charts)

        for i in range(3):
            cfg = next(chart_iter, None)
            if cfg:
                zone = ChartZone(11, i * 8, CHART_THIRD_W, CHART_HALF_H)
                self._add_chart(cfg, df, zone)

        # ── Row 26: Second chart row (row 11 + 15) ──────────────────────────
        self._write_section_header(26, "📊  Trend & Distribution", color=t.secondary)

        # Remaining charts in pairs
        next_row = 27
        remaining_charts = list(chart_iter)
        for ri in range(0, len(remaining_charts), 2):
            zone = ChartZone(next_row, 0, CHART_HALF_W, CHART_HALF_H)
            self._add_chart(remaining_charts[ri], df, zone)
            if ri + 1 < len(remaining_charts):
                zone = ChartZone(next_row, 12, CHART_HALF_W, CHART_HALF_H)
                self._add_chart(remaining_charts[ri + 1], df, zone)
            next_row += 15

        # ── Table ────────────────────────────────────────────────────────────
        self._write_section_header(next_row, "📋  Operational Data", color=t.bg_table_header)
        self._write_summary_table(df, next_row + 1, max_rows=20, max_cols=8)

        ws.freeze_panes(4, 0)
        ws.set_zoom(80)
        return self._close(output_path)

    def _write_mini_card_grid(self, df: pd.DataFrame,
                               start_row: int, filter_cols: list[str]) -> None:
        """Write a 3×4 grid of colored mini-metric cards (like HR-Metrics screenshot)."""
        ws = self._ws_dash
        t = self.theme
        ws.set_row(start_row, 4)

        if not filter_cols or filter_cols[0] not in df.columns:
            return

        fc = filter_cols[0]
        unique_cats = df[fc].dropna().value_counts().head(12).index.tolist()
        colors = [_hex(c) for c in t.chart_colors] * 2

        num_col = next(
            (c for c in df.select_dtypes("number").columns
             if c in (self.config.table_columns or [])), None
        )
        if num_col is None:
            num_cols_list = df.select_dtypes("number").columns.tolist()
            num_col = num_cols_list[0] if num_cols_list else None

        for i, cat in enumerate(unique_cats[:12]):
            row = start_row + 1 + (i // 4)
            col = (i % 4) * 6
            bg = colors[i % len(colors)]
            bg_fmt = self._wb.add_format({
                "bg_color": bg, "font_color": "#FFFFFF",
                "font_name": t.font_body, "font_size": 8,
                "align": "left", "valign": "vcenter", "indent": 1,
            })
            val_fmt = self._wb.add_format({
                "bg_color": bg, "font_color": "#FFFFFF",
                "font_name": t.font_heading, "font_size": 13, "bold": True,
                "align": "center", "valign": "vcenter",
            })
            ws.set_row(row, 16)
            ws.set_row(row + 1, 22)

            ws.merge_range(row, col, row, col + 5, str(cat), bg_fmt)
            if num_col:
                try:
                    val = df[df[fc] == cat][num_col].sum()
                    val_str = (f"${val/1e6:.1f}M" if val >= 1e6
                               else f"{val/1e3:.0f}K" if val >= 1e3
                               else f"{val:,.0f}")
                    ws.merge_range(row + 1, col, row + 1, col + 5, val_str, val_fmt)
                except Exception:
                    ws.merge_range(row + 1, col, row + 1, col + 5, "", val_fmt)
