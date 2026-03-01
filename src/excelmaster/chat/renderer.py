"""FlexibleTemplate — renders a WorkbookState to an xlsx file."""
from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from ..dashboard.templates.base_xl_template import (
    BaseXLTemplate,
    CHART_HALF_W,
    CHART_HALF_H,
    CHART_FULL_W,
    CHART_FULL_H,
    N_COLS,
    COL_W,
)
from ..dashboard.xl_chart import ChartZone
from ..dashboard.xl_style import StyleFactory, _hex
from ..dashboard.themes import get_theme
from ..models import (
    AggFunc,
    ChartConfig,
    ColorTheme,
    DashboardConfig,
    DashboardTemplate,
    KPIConfig,
    NumberFormat,
)
from .models import (
    ObjectType,
    PlacedChart,
    PlacedFilterPanel,
    PlacedKPIRow,
    PlacedObject,
    PlacedPivot,
    PlacedSectionHeader,
    PlacedTable,
    PlacedText,
    PlacedTitle,
    SheetLayout,
    WorkbookState,
)


class FlexibleTemplate(BaseXLTemplate):
    """Renders an arbitrary WorkbookState to xlsx using inherited base helpers."""

    name = "flexible_chat"

    def __init__(self, state: WorkbookState):
        # Build a minimal DashboardConfig from state so the base class is happy
        config = DashboardConfig(
            template=DashboardTemplate.EXECUTIVE_SUMMARY,
            title=state.title,
            theme=ColorTheme(state.theme_key) if state.theme_key else ColorTheme.CORPORATE_BLUE,
            kpis=self._collect_kpis(state),
            charts=self._collect_charts(state),
            table_columns=self._collect_table_cols(state),
            filters=[],
        )
        super().__init__(config)
        self.state = state

    # ── Helpers to extract config fields from state ──────────────────────────

    @staticmethod
    def _collect_kpis(state: WorkbookState) -> list[KPIConfig]:
        kpis: list[KPIConfig] = []
        for sheet in state.sheets:
            for obj in sheet.objects:
                if obj.type == ObjectType.KPI_ROW:
                    payload: PlacedKPIRow = obj.payload  # type: ignore[assignment]
                    kpis.extend(payload.kpis)
        return kpis

    @staticmethod
    def _collect_charts(state: WorkbookState) -> list[ChartConfig]:
        charts: list[ChartConfig] = []
        for sheet in state.sheets:
            for obj in sheet.objects:
                if obj.type == ObjectType.CHART:
                    payload: PlacedChart = obj.payload  # type: ignore[assignment]
                    charts.append(payload.chart)
        return charts

    @staticmethod
    def _collect_table_cols(state: WorkbookState) -> list[str]:
        for sheet in state.sheets:
            for obj in sheet.objects:
                if obj.type == ObjectType.TABLE:
                    payload: PlacedTable = obj.payload  # type: ignore[assignment]
                    if payload.columns:
                        return payload.columns
        return []

    # ── build() is required by ABC but we use build_from_state() ─────────────

    def build(self, df: pd.DataFrame, output_path: Path) -> Path:
        return self.build_from_state(df, output_path)

    def build_from_state(self, df: pd.DataFrame, output_path: Path) -> Path:
        """Full render: write data, then render every sheet's objects."""
        output_path = Path(output_path)
        self._init_workbook(output_path)
        self._write_data_sheet(df)

        ws_dash = self._ws_dash
        t = self.theme

        # ── Discover filter panel and set up engine ──────────────────────────
        filter_refs: dict[str, str] = {}
        dash_sheet = self.state.dashboard_sheet()

        for obj in dash_sheet.sorted_objects():
            if obj.type == ObjectType.FILTER_PANEL:
                fp: PlacedFilterPanel = obj.payload  # type: ignore[assignment]
                valid_cols = [c for c in fp.filter_columns if c in df.columns]
                if valid_cols:
                    filter_refs = self._write_filter_slicer_panel(
                        df, row=obj.anchor_row, filter_cols=valid_cols,
                    )
                break

        self._build_engine(df, filter_refs)

        # ── Render Dashboard sheet ───────────────────────────────────────────
        ws_dash.hide_gridlines(2)
        for c in range(N_COLS):
            ws_dash.set_column(c, c, COL_W)

        for obj in dash_sheet.sorted_objects():
            if obj.type == ObjectType.FILTER_PANEL:
                continue  # already rendered above
            self._dispatch_render(obj, df, ws_dash)

        if dash_sheet.freeze_row:
            ws_dash.freeze_panes(dash_sheet.freeze_row, 0)
        ws_dash.set_zoom(dash_sheet.zoom or 85)

        # ── Render extra sheets ──────────────────────────────────────────────
        for sheet_layout in self.state.sheets:
            if sheet_layout.name == "Dashboard":
                continue
            if sheet_layout.name in ("Data", "Calculations", "Deep Analysis"):
                continue
            ws_extra = self._wb.add_worksheet(sheet_layout.name)
            ws_extra.hide_gridlines(2)
            for c in range(N_COLS):
                ws_extra.set_column(c, c, COL_W)
            for obj in sheet_layout.sorted_objects():
                self._dispatch_render(obj, df, ws_extra)
            if sheet_layout.freeze_row:
                ws_extra.freeze_panes(sheet_layout.freeze_row, 0)
            ws_extra.set_zoom(sheet_layout.zoom or 85)

        return self._close(output_path)

    # ── Dispatch ─────────────────────────────────────────────────────────────

    def _dispatch_render(self, obj: PlacedObject, df: pd.DataFrame, ws) -> None:
        renderers = {
            ObjectType.TITLE: self._render_title,
            ObjectType.KPI_ROW: self._render_kpi_row,
            ObjectType.SECTION_HEADER: self._render_section_header,
            ObjectType.CHART: self._render_chart,
            ObjectType.TABLE: self._render_table,
            ObjectType.PIVOT: self._render_pivot,
            ObjectType.TEXT: self._render_text,
        }
        fn = renderers.get(obj.type)
        if fn:
            fn(obj, df, ws)

    # ── Render methods ───────────────────────────────────────────────────────

    def _render_title(self, obj: PlacedObject, df: pd.DataFrame, ws) -> None:
        p: PlacedTitle = obj.payload  # type: ignore[assignment]
        t = self.theme
        row = obj.anchor_row

        ws.set_row(row, 48)
        title_fmt = self._wb.add_format({
            "font_name": t.font_heading, "font_size": 22, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.primary),
            "align": "left", "valign": "vcenter", "indent": 2,
        })
        ws.merge_range(row, 0, row, N_COLS - 1, f"  {p.text}", title_fmt)

        if p.subtitle:
            row += 1
            ws.set_row(row, 20)
            sub_fmt = self._wb.add_format({
                "font_name": t.font_body, "font_size": 10,
                "font_color": _hex(t.text_light), "bg_color": _hex(t.secondary),
                "align": "left", "valign": "vcenter", "indent": 2,
            })
            ws.merge_range(row, 0, row, N_COLS - 1, f"  {p.subtitle}", sub_fmt)

    def _render_kpi_row(self, obj: PlacedObject, df: pd.DataFrame, ws) -> None:
        p: PlacedKPIRow = obj.payload  # type: ignore[assignment]
        t = self.theme
        kpis = p.kpis
        if not kpis:
            return

        n = len(kpis)
        card_span = max(4, N_COLS // n)
        card_colors = [t.primary, t.secondary, t.accent1, t.accent2, t.accent3]
        row = obj.anchor_row

        for i, kpi in enumerate(kpis):
            c = i * card_span
            if c + card_span > N_COLS:
                break
            bg = card_colors[i % len(card_colors)]
            # Use the Dashboard worksheet for KPI tiles (they need _ws_dash for sparklines)
            saved_ws = self._ws_dash
            self._ws_dash = ws
            self._write_kpi_tile(
                row, c, card_span, 5, kpi, df, bg,
                font_color=t.text_light,
                filter_ref=self._filter_cell,
            )
            self._ws_dash = saved_ws

    def _render_section_header(self, obj: PlacedObject, df: pd.DataFrame, ws) -> None:
        p: PlacedSectionHeader = obj.payload  # type: ignore[assignment]
        color = p.color if p.color else None
        self._write_section_header(obj.anchor_row, p.text, color=color, ws=ws)

    def _render_chart(self, obj: PlacedObject, df: pd.DataFrame, ws) -> None:
        p: PlacedChart = obj.payload  # type: ignore[assignment]
        chart_cfg = p.chart

        if p.width == "full":
            zone = ChartZone(obj.anchor_row, 0, CHART_FULL_W, CHART_FULL_H)
        elif p.side == "right":
            zone = ChartZone(obj.anchor_row, 12, CHART_HALF_W, CHART_HALF_H)
        else:
            zone = ChartZone(obj.anchor_row, 0, CHART_HALF_W, CHART_HALF_H)

        if self._engine is None:
            return

        # Write chart table to Calculations sheet and build chart
        table = self._engine.write_chart_table(chart_cfg)
        if table is None:
            return

        from ..dashboard.xl_chart import build_xl_chart
        val_ranges = [(table.header_cell, table.val_range)]
        chart = build_xl_chart(
            self._wb, chart_cfg, self.theme,
            table.cat_range, val_ranges, zone,
        )
        ws.insert_chart(
            zone.row, zone.col, chart,
            {"x_offset": zone.x_offset, "y_offset": zone.y_offset},
        )

    def _render_table(self, obj: PlacedObject, df: pd.DataFrame, ws) -> None:
        p: PlacedTable = obj.payload  # type: ignore[assignment]
        t = self.theme

        # Override config.table_columns temporarily
        saved_cols = self.config.table_columns
        if p.columns:
            self.config.table_columns = p.columns

        saved_ws = self._ws_dash
        self._ws_dash = ws
        rows_used = self._write_summary_table(
            df, obj.anchor_row, max_rows=p.max_rows, ws=ws,
        )
        self._ws_dash = saved_ws
        self.config.table_columns = saved_cols

        # Conditional formatting on numeric columns
        if p.show_conditional:
            cols = p.columns or list(df.columns)[:8]
            valid_cols = [c for c in cols if c in df.columns]
            num_cols = [c for c in valid_cols if df[c].dtype.kind in "iuf"]
            for j, nc in enumerate(num_cols[:3]):
                col_idx = valid_cols.index(nc) if nc in valid_cols else j
                ws.conditional_format(
                    obj.anchor_row + 1, col_idx,
                    obj.anchor_row + p.max_rows, col_idx,
                    {
                        "type": "3_color_scale",
                        "min_color": _hex(t.negative),
                        "mid_color": "#FFFFAA",
                        "max_color": _hex(t.positive),
                    },
                )

    def _render_pivot(self, obj: PlacedObject, df: pd.DataFrame, ws) -> None:
        p: PlacedPivot = obj.payload  # type: ignore[assignment]
        t = self.theme
        sf = self._sf

        # Validate columns
        if p.index_col not in df.columns or p.value_col not in df.columns:
            return

        agg_map = {
            AggFunc.SUM: "sum", AggFunc.AVG: "mean", AggFunc.COUNT: "count",
            AggFunc.MAX: "max", AggFunc.MIN: "min", AggFunc.MEDIAN: "median",
        }
        agg_fn = agg_map.get(p.agg, "sum")

        try:
            if p.columns_col and p.columns_col in df.columns:
                pivot = pd.pivot_table(
                    df, values=p.value_col, index=p.index_col,
                    columns=p.columns_col, aggfunc=agg_fn, fill_value=0,
                )
            else:
                pivot = pd.pivot_table(
                    df, values=p.value_col, index=p.index_col,
                    aggfunc=agg_fn, fill_value=0,
                )
        except Exception:
            return

        # Write pivot to worksheet
        row = obj.anchor_row
        hdr_fmt = sf.table_header()
        data_fmt = sf.table_data()
        num_fmt = sf.table_data_num()

        # Header row
        ws.set_row(row, 18)
        ws.write(row, 0, p.index_col.replace("_", " ").title(), hdr_fmt)
        if isinstance(pivot.columns, pd.MultiIndex):
            col_labels = [str(c) for c in pivot.columns.get_level_values(-1)]
        else:
            col_labels = [str(c) for c in pivot.columns]
        for j, label in enumerate(col_labels[:15]):
            ws.write(row, 1 + j, label, hdr_fmt)

        # Data rows
        for i, (idx_val, data_row) in enumerate(pivot.head(20).iterrows()):
            r = row + 1 + i
            stripe = i % 2 == 0
            ws.set_row(r, 15)
            ws.write(r, 0, str(idx_val), sf.table_data(stripe))
            for j, val in enumerate(data_row.values[:15]):
                try:
                    ws.write(r, 1 + j, float(val), sf.table_data_num(stripe))
                except (ValueError, TypeError):
                    ws.write(r, 1 + j, str(val), sf.table_data(stripe))

    def _render_text(self, obj: PlacedObject, df: pd.DataFrame, ws) -> None:
        p: PlacedText = obj.payload  # type: ignore[assignment]
        t = self.theme

        style_map = {
            "heading": {"font_size": 14, "bold": True, "font_color": _hex(t.primary)},
            "insight": {"font_size": 10, "italic": True, "font_color": _hex(t.secondary),
                        "bg_color": _hex(t.bg_card)},
            "footnote": {"font_size": 8, "italic": True, "font_color": _hex(t.text_muted)},
            "body": {"font_size": 10, "font_color": _hex(t.text_primary)},
        }
        props = style_map.get(p.style, style_map["body"])
        props["font_name"] = t.font_body
        props["text_wrap"] = True
        props["valign"] = "top"
        fmt = self._wb.add_format(props)

        row = obj.anchor_row
        ws.set_row(row, 40 if p.style == "heading" else 30)
        ws.merge_range(row, 0, row + obj.height_rows - 1, N_COLS - 1, p.content, fmt)
