"""
Marketing Template — Vibrant orange, conversion funnel KPIs, campaign analytics

Layout: Orange accent, funnel-style KPIs, channel + trend analysis
  Row 0:      Bold orange title
  Row 1-2:    3 filter dropdowns (Channel, Campaign Type, Quarter)
  Rows 3-6:   FUNNEL KPI strip — Impressions → Clicks → Leads → MQLs → Deals
              (with arrows between, showing conversion funnel)
  Row 7:      Section header "Channel Performance"
  Rows 8-20:  Bar (left 14 cols) | Donut (right 10 cols)
  Row 21:     Section header "Monthly Trend"
  Rows 22-31: Full-width line chart
  Row 32:     Section header "Campaign Data"
  Rows 33+:   Campaign table with pct/currency formatting + ROI color scale
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


class MarketingXL(BaseXLTemplate):
    """Marketing analytics — orange, conversion funnel KPIs, channel charts."""
    name = "marketing"

    def build(self, df: pd.DataFrame, output_path: Path) -> Path:
        self._init_workbook(output_path)
        self._write_data_sheet(df)

        ws = self._ws_dash
        sf = self._sf
        t = self.theme

        ws.set_column(0, N_COLS - 1, COL_W)
        ws.hide_gridlines(2)

        # Warm light background
        warm_bg_fmt = self._wb.add_format({"bg_color": _hex(t.bg_dashboard)})
        for r in range(70):
            for c in range(N_COLS):
                ws.write_blank(r, c, None, warm_bg_fmt)

        # ── Row 0: Bold orange title ───────────────────────────────────────────
        ws.set_row(0, 48)
        title_fmt = self._wb.add_format({
            "font_name": t.font_heading, "font_size": 22, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.primary),
            "align": "left", "valign": "vcenter", "indent": 2,
        })
        tagline_fmt = self._wb.add_format({
            "font_name": t.font_body, "font_size": 10,
            "font_color": _hex(t.primary), "bg_color": _hex(t.bg_section),
            "align": "right", "valign": "vcenter", "indent": 2,
        })
        ws.merge_range(0, 0, 0, 14, f"  {self.config.title}", title_fmt)
        ws.merge_range(0, 15, 0, N_COLS - 1,
                       self.config.subtitle or "Campaign Analytics", tagline_fmt)

        # ── Rows 1-2: Filter panel ─────────────────────────────────────────────
        ws.set_row(1, 30)
        ws.set_row(2, 14)
        filter_cols = self._detect_filter_cols(
            df, prefer_keywords=["channel", "campaign_type", "platform",
                                  "geo_target", "quarter", "status"])
        filter_refs = self._write_filter_slicer_panel(
            df, row=1, filter_cols=filter_cols,
            panel_bg=t.bg_section)

        # ── Rows 3-6: Funnel KPI strip with arrows ─────────────────────────────
        self._write_funnel_kpis(df, start_row=3,
                                 filter_ref=list(filter_refs.values())[0]
                                 if filter_refs else None)

        # ── Row 7: Section header ─────────────────────────────────────────────
        self._write_section_header(7, "📣  Channel Performance", color=t.primary)

        # ── Rows 8-21: Bar (left) | Donut (right) (H=280 → ~14 rows) ───────
        self._build_engine(df, filter_refs)
        charts = self.config.charts

        bar_cfg = self._pick(charts, ChartType.BAR) or charts[0] if charts else None
        donut_cfg = (self._pick(charts, ChartType.DOUGHNUT)
                     or self._pick(charts, ChartType.PIE))
        line_cfg = self._pick(charts, ChartType.LINE)

        if bar_cfg:
            zone = ChartZone(8, 0, 770, CHART_HALF_H)
            self._add_chart(bar_cfg, df, zone)
        if donut_cfg:
            zone = ChartZone(8, 14, 510, CHART_HALF_H)
            self._add_chart(donut_cfg, df, zone)

        # ── Row 23: Section header (row 8 + 15) ─────────────────────────────
        self._write_section_header(23, "📈  Monthly Performance Trend", color=t.secondary)

        # ── Rows 24-36: Full-width line (H=250 → ~13 rows) ──────────────────
        next_row = 24
        if line_cfg:
            zone = ChartZone(next_row, 0, CHART_FULL_W, CHART_FULL_H)
            self._add_chart(line_cfg, df, zone)
            next_row += 14

        # ── Remaining charts in pairs ────────────────────────────────────────
        placed = [bar_cfg, donut_cfg, line_cfg]
        remaining = [c for c in charts if c not in placed]
        for ri in range(0, len(remaining), 2):
            left = remaining[ri]
            zone = ChartZone(next_row, 0, CHART_HALF_W, CHART_HALF_H)
            self._add_chart(left, df, zone)
            if ri + 1 < len(remaining):
                zone = ChartZone(next_row, 12, CHART_HALF_W, CHART_HALF_H)
                self._add_chart(remaining[ri + 1], df, zone)
            next_row += 15

        # ── Table section ────────────────────────────────────────────────────
        self._write_section_header(next_row, "📋  Campaign Data", color=t.bg_table_header)
        self._write_campaign_table(df, next_row + 1)

        ws.freeze_panes(3, 0)
        ws.set_zoom(85)
        return self._close(output_path)

    def _write_funnel_kpis(self, df: pd.DataFrame, start_row: int,
                            filter_ref: str | None = None) -> None:
        """Write 5 KPI tiles showing marketing conversion funnel with arrows."""
        ws = self._ws_dash
        t = self.theme

        ws.set_row(start_row, 8)
        ws.set_row(start_row + 1, 14)
        ws.set_row(start_row + 2, 32)
        ws.set_row(start_row + 3, 8)

        kpis = self.config.kpis[:5]
        funnel_colors = [t.primary, t.secondary, t.accent1, t.accent2, t.accent3]
        n = max(len(kpis), 1)
        tile_w = N_COLS // n
        arrow_col = tile_w - 1  # position of arrow within each tile section

        for i, kpi in enumerate(kpis):
            c = i * tile_w
            bg = funnel_colors[i % len(funnel_colors)]
            self._write_kpi_tile(start_row, c, tile_w - 1, 4, kpi, df,
                                  bg, font_color=t.text_light,
                                  filter_ref=filter_ref)

            # Arrow between tiles (except last)
            if i < len(kpis) - 1:
                arrow_fmt = self._wb.add_format({
                    "font_name": t.font_heading, "font_size": 18, "bold": True,
                    "font_color": _hex(t.text_muted),
                    "bg_color": _hex(t.bg_dashboard),
                    "align": "center", "valign": "vcenter",
                })
                ws.merge_range(start_row, c + tile_w - 1,
                               start_row + 3, c + tile_w - 1, "›", arrow_fmt)

    def _write_campaign_table(self, df: pd.DataFrame, start_row: int) -> None:
        ws = self._ws_dash
        sf = self._sf
        t = self.theme

        cols = [c for c in self.config.table_columns if c in df.columns][:8]
        if not cols:
            cols = list(df.columns)[:8]

        pct_kw = ["pct", "rate", "percent", "roi", "roas", "conversion"]
        curr_kw = ["usd", "spend", "budget", "revenue", "cost", "cpc", "cpl", "cac"]

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

        for j, col in enumerate(cols):
            if any(k in col.lower() for k in ["roi", "roas", "rate", "conversion"]):
                ws.conditional_format(start_row + 1, j, start_row + 15, j, {
                    "type": "3_color_scale",
                    "min_color": _hex(t.negative),
                    "mid_color": "#FFFFAA",
                    "max_color": _hex(t.positive),
                })

    def _pick(self, charts, ctype: ChartType):
        return next((c for c in charts if c.type == ctype), None)
