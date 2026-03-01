"""
Minimal Clean Template — Ultra-sparse, slate/white, survey/research reporting

Layout: White, lots of whitespace, report-like
  Row 0:      Simple clean title (no subtitle bar)
  Row 1:      One filter control (wide, centered)
  Row 2:      Spacer
  Row 3-5:    3 KPI numbers — text-only (no card bg), extra large
  Row 6:      Horizontal rule
  Row 7:      Section label
  Rows 8-20:  One full-width chart
  Row 21:     Section label
  Rows 22+:   Clean table with null% annotations, wider columns
"""
from __future__ import annotations
from pathlib import Path
import pandas as pd

from .base_xl_template import (
    BaseXLTemplate, CHART_HALF_W, CHART_HALF_H,
    CHART_FULL_W, CHART_FULL_H, N_COLS, COL_W,
)
from ..xl_chart import ChartZone
from ..xl_style import _hex
from ...models import ChartType


class MinimalCleanXL(BaseXLTemplate):
    """Minimal clean dashboard — white, sparse layout, report-style."""
    name = "minimal_clean"

    def build(self, df: pd.DataFrame, output_path: Path) -> Path:
        self._init_workbook(output_path, col_width=8.5)
        self._write_data_sheet(df)

        ws = self._ws_dash
        sf = self._sf
        t = self.theme

        # Wider columns for minimal feel
        ws.set_column(0, N_COLS - 1, 8.5)
        ws.hide_gridlines(2)

        # Clean white background
        white_fmt = self._wb.add_format({"bg_color": "#FFFFFF"})
        for r in range(100):
            for c in range(N_COLS):
                ws.write_blank(r, c, None, white_fmt)

        # ── Row 0: Simple title ────────────────────────────────────────────────
        ws.set_row(0, 40)
        title_fmt = self._wb.add_format({
            "font_name": t.font_heading, "font_size": 18, "bold": True,
            "font_color": _hex(t.primary), "bg_color": "#FFFFFF",
            "align": "left", "valign": "vcenter", "indent": 2,
            "bottom": 3, "bottom_color": _hex(t.primary),
        })
        ws.merge_range(0, 0, 0, N_COLS - 1, f"  {self.config.title}", title_fmt)

        # ── Row 1: Single prominent filter ────────────────────────────────────
        ws.set_row(1, 30)
        ws.set_row(2, 12)

        filter_col = self._detect_filter_col(df)
        filter_refs = {}
        if filter_col and filter_col in df.columns:
            options = ["All"] + sorted(
                [str(v) for v in df[filter_col].dropna().unique().tolist()]
            )[:25]
            from ..xl_dynamic import add_filter_dropdown, col_letter
            from ..xl_dynamic import col_letter
            import io

            # Write filter options to Calculations sheet
            for vi, v in enumerate(options):
                self._ws_calc.write(vi, 12, v)
            calc_range = f"Calculations!$M$1:$M${len(options)}"

            filter_fmt = self._wb.add_format({
                "font_name": t.font_body, "font_size": 11, "bold": True,
                "font_color": _hex(t.primary), "bg_color": "#F8FAFC",
                "align": "left", "valign": "vcenter", "indent": 2,
                "border": 2, "border_color": _hex(t.primary),
            })
            filter_label_fmt = self._wb.add_format({
                "font_name": t.font_body, "font_size": 9,
                "font_color": _hex(t.text_muted), "bg_color": "#FFFFFF",
                "align": "left", "valign": "vcenter",
            })
            ws.write(1, 0, f"Filter by {filter_col.replace('_', ' ').title()}:", filter_label_fmt)
            ws.merge_range(1, 1, 1, 6, "All", filter_fmt)
            add_filter_dropdown(ws, 1, 1, options, calc_range)
            ws.write(2, 1, "▲ Click the cell above to select a filter value", filter_label_fmt)

            filter_refs[filter_col] = f"'Dashboard'!$B$2"

        self._build_engine(df, filter_refs)

        # ── Row 3: Spacer ─────────────────────────────────────────────────────
        ws.set_row(3, 16)

        # ── Rows 4-6: 3 large text-only KPI numbers ───────────────────────────
        kpis = self.config.kpis[:3]
        kpi_colors = [t.primary, t.accent1, t.accent2]
        kpi_span = N_COLS // max(len(kpis), 1)

        ws.set_row(4, 14)  # label
        ws.set_row(5, 46)  # BIG number
        ws.set_row(6, 10)  # bottom pad

        for i, kpi in enumerate(kpis):
            c = i * kpi_span
            color = kpi_colors[i % len(kpi_colors)]

            lbl_fmt = self._wb.add_format({
                "font_name": t.font_body, "font_size": 9,
                "font_color": _hex(t.text_muted), "bg_color": "#FFFFFF",
                "align": "center", "valign": "vcenter",
            })
            val_fmt = self._wb.add_format({
                "font_name": t.font_heading, "font_size": 28, "bold": True,
                "font_color": _hex(color), "bg_color": "#FFFFFF",
                "align": "center", "valign": "vcenter",
                "bottom": 2, "bottom_color": _hex(color),
            })
            bot_fmt = self._wb.add_format({"bg_color": "#FFFFFF"})

            label_txt = kpi.label.replace("_", " ").title()
            ws.merge_range(4, c, 4, c + kpi_span - 1, label_txt, lbl_fmt)
            ws.merge_range(6, c, 6, c + kpi_span - 1, "", bot_fmt)

            if kpi.column in self._col_index:
                val_str = self._compute_kpi_static(df, kpi)
                ws.merge_range(5, c, 5, c + kpi_span - 1, val_str, val_fmt)
            else:
                ws.merge_range(5, c, 5, c + kpi_span - 1, "N/A", val_fmt)

        # ── Row 7: Horizontal rule ─────────────────────────────────────────────
        ws.set_row(7, 3)
        rule_fmt = self._wb.add_format({"bg_color": _hex(t.primary)})
        ws.merge_range(7, 0, 7, N_COLS - 1, "", rule_fmt)

        # ── Row 8: Section label ───────────────────────────────────────────────
        self._write_section_header(8, "Analysis", color=t.bg_section if t.bg_section else t.secondary,
                                   height=18)

        # ── Charts: render ALL in pairs (2 per row, 15-row gap each) ────────
        charts = self.config.charts
        chart_row = 9
        for ci in range(0, len(charts), 2):
            left = charts[ci]
            right = charts[ci + 1] if ci + 1 < len(charts) else None
            zone = ChartZone(chart_row, 0, CHART_HALF_W, CHART_HALF_H)
            self._add_chart(left, df, zone)
            if right:
                zone = ChartZone(chart_row, 12, CHART_HALF_W, CHART_HALF_H)
                self._add_chart(right, df, zone)
            chart_row += 15

        # ── Data table: starts right after last chart row ────────────────────
        tbl_start = chart_row
        self._write_section_header(tbl_start, "Data Summary",
                                   color=t.bg_section if t.bg_section else t.secondary, height=18)
        self._write_clean_table(df, tbl_start + 1)

        ws.freeze_panes(4, 0)
        ws.set_zoom(90)
        return self._close(output_path)

    def _write_clean_table(self, df: pd.DataFrame, start_row: int) -> None:
        """Clean table with null% annotations and wider columns."""
        ws = self._ws_dash
        sf = self._sf

        cols = [c for c in self.config.table_columns if c in df.columns][:8]
        if not cols:
            cols = list(df.columns)[:8]

        for j in range(len(cols)):
            ws.set_column(j, j, 12)

        hdr_fmt = sf.table_header()
        ws.set_row(start_row, 18)
        for j, c in enumerate(cols):
            null_pct = df[c].isna().mean() * 100
            label = c.replace("_", " ").title()
            if null_pct > 5:
                label += f" ({null_pct:.0f}% null)"
            ws.write(start_row, j, label, hdr_fmt)

        display = df[cols].head(20)
        for i, row in enumerate(display.itertuples(index=False), 1):
            stripe = i % 2 == 0
            ws.set_row(start_row + i, 16)
            for j, (cn, val) in enumerate(zip(cols, row)):
                fmt = sf.table_data_num(stripe) if isinstance(val, (int, float)) \
                    else sf.table_data(stripe)
                try:
                    ws.write(start_row + i, j, val, fmt)
                except Exception:
                    ws.write(start_row + i, j, str(val) if val else "", fmt)

    def _detect_filter_col(self, df: pd.DataFrame) -> str | None:
        if self.config.primary_dimension and self.config.primary_dimension in df.columns:
            return self.config.primary_dimension
        cats = df.select_dtypes("object").columns.tolist()
        if cats:
            return min(cats, key=lambda c: df[c].nunique())
        return None
