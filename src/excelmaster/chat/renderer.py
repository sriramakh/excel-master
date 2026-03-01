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
    DatasetProfile,
    KPIConfig,
    NumberFormat,
)
from .models import (
    CellFormatOp,
    CellWrite,
    CommentOp,
    ConditionalFormatOp,
    DataValidationOp,
    HyperlinkOp,
    ImageOp,
    MergeOp,
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


# ── Cell Address Parsing ──────────────────────────────────────────────────────

def _parse_cell_address(cell: str) -> tuple[int, int]:
    """Parse 'B3' → (row=2, col=1). Zero-based."""
    from xlsxwriter.utility import xl_cell_to_rowcol
    return xl_cell_to_rowcol(cell)


def _parse_range_address(rng: str) -> tuple[int, int, int, int]:
    """Parse 'A1:F20' → (r1, c1, r2, c2). Zero-based."""
    from xlsxwriter.utility import xl_range_abs
    from xlsxwriter.utility import xl_cell_to_rowcol
    if ":" in rng:
        parts = rng.split(":")
        r1, c1 = xl_cell_to_rowcol(parts[0])
        r2, c2 = xl_cell_to_rowcol(parts[1])
        return r1, c1, r2, c2
    r, c = xl_cell_to_rowcol(rng)
    return r, c, r, c


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

    def build_from_state(
        self,
        df: pd.DataFrame,
        output_path: Path,
        profile: DatasetProfile | None = None,
    ) -> Path:
        """Full render: write data, then render every sheet's objects.

        Args:
            df: The dataset to render.
            output_path: Where to save the xlsx.
            profile: If provided, enables Deep Analysis sheet generation.
        """
        output_path = Path(output_path)
        self._init_workbook(output_path)
        self._write_data_sheet(df)

        # ── Deep Analysis generation ──────────────────────────────────────
        if profile is not None:
            self._store_for_analysis(df, profile)
            self._generate_deep_analysis(df, profile)

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

        # ── Cell-level operations on Dashboard ─────────────────────────
        self._render_cell_operations(dash_sheet, ws_dash)

        if dash_sheet.freeze_row or dash_sheet.freeze_col:
            ws_dash.freeze_panes(dash_sheet.freeze_row, dash_sheet.freeze_col)
        ws_dash.set_zoom(dash_sheet.zoom or 85)
        if dash_sheet.tab_color:
            ws_dash.set_tab_color(dash_sheet.tab_color)
        if dash_sheet.hidden:
            ws_dash.hide()

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
            self._render_cell_operations(sheet_layout, ws_extra)
            if sheet_layout.freeze_row or sheet_layout.freeze_col:
                ws_extra.freeze_panes(sheet_layout.freeze_row, sheet_layout.freeze_col)
            ws_extra.set_zoom(sheet_layout.zoom or 85)
            if sheet_layout.tab_color:
                ws_extra.set_tab_color(sheet_layout.tab_color)
            if sheet_layout.hidden:
                ws_extra.hide()

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
            saved_ws = self._ws_dash
            self._ws_dash = ws
            self._write_kpi_tile_formatted(
                row, c, card_span, 5, kpi, df, bg, ws,
                font_color=t.text_light,
            )
            self._ws_dash = saved_ws

    def _write_kpi_tile_formatted(
        self,
        row: int, col: int, span_cols: int, span_rows: int,
        kpi: KPIConfig, df: pd.DataFrame, bg: str,
        ws=None,
        font_color: str | None = None,
    ) -> None:
        """Write a KPI tile with pre-formatted display text and sparkline."""
        ws = ws or self._ws_dash
        sf = self._sf
        t = self.theme

        bg_fmt = sf.kpi_bg(bg)
        lbl_fmt = sf.kpi_label(bg)
        val_fmt = sf.kpi_value(bg, font_color or (t.text_light if t.dark_mode else None))

        # Top padding
        ws.merge_range(row, col, row, col + span_cols - 1, "", bg_fmt)

        # Label row
        lbl_row = row + 1
        ws.set_row(lbl_row, 15)
        label_txt = kpi.icon + "  " + kpi.label if kpi.icon else kpi.label
        ws.merge_range(lbl_row, col, lbl_row, col + span_cols - 1, label_txt, lbl_fmt)

        # Value row — formatted static text (e.g. "$3.1B", "45.2%", "12.5K")
        val_row = row + 2
        ws.set_row(val_row, 32)
        display_val = self._compute_kpi_static(df, kpi)
        ws.merge_range(val_row, col, val_row, col + span_cols - 1, display_val, val_fmt)

        # Sparkline row
        if span_rows >= 4:
            spark_row = row + 3
            ws.set_row(spark_row, 18)
            for ci in range(col, col + span_cols):
                ws.write_blank(spark_row, ci, None, bg_fmt)
            self._add_kpi_sparkline(df, kpi, spark_row, col + span_cols // 2, bg)

        # Bottom padding
        bot_row = row + span_rows - 1
        ws.merge_range(bot_row, col, bot_row, col + span_cols - 1, "", bg_fmt)

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

    # ── Deep Analysis Generation ─────────────────────────────────────────────

    def _generate_deep_analysis(
        self,
        df: pd.DataFrame,
        profile: DatasetProfile,
    ) -> None:
        """Generate deep analysis via LLM and attach to config."""
        try:
            from ..dashboard.deep_analysis import (
                compute_deep_stats,
                build_analysis_prompt,
                safe_parse_analysis,
            )
            from ..dashboard.llm_client import LLMClient

            stats = compute_deep_stats(df, profile)
            sys_prompt, user_prompt = build_analysis_prompt(
                stats, profile, self.config,
            )
            llm = LLMClient()
            raw = llm.generate_json(
                sys_prompt, user_prompt, max_tokens_override=6144,
            )
            self.config.deep_analysis = safe_parse_analysis(raw)
        except Exception:
            # Deep analysis is optional — never break the build
            self.config.deep_analysis = None

    # ── Cell-Level Operations ─────────────────────────────────────────────────

    def _render_cell_operations(self, sheet_layout: SheetLayout, ws) -> None:
        """Render all cell-level operations after placed objects."""
        wb = self._wb

        # 1. Row heights
        for row_idx, height in sheet_layout.row_heights.items():
            ws.set_row(row_idx, height)

        # 2. Column widths
        for col_idx, width in sheet_layout.col_widths.items():
            ws.set_column(col_idx, col_idx, width)

        # 3. Hidden rows
        for r in sheet_layout.hidden_rows:
            ws.set_row(r, None, None, {"hidden": True})

        # 4. Hidden columns
        for c in sheet_layout.hidden_cols:
            ws.set_column(c, c, None, None, {"hidden": True})

        # 5. Cell writes (values + formulas + inline format)
        for cw in sheet_layout.cell_writes:
            row, col = _parse_cell_address(cw.cell)
            fmt = wb.add_format(cw.format) if cw.format else None
            if cw.value is not None:
                if isinstance(cw.value, str) and cw.value.startswith("="):
                    ws.write_formula(row, col, cw.value, fmt)
                else:
                    ws.write(row, col, cw.value, fmt)
            elif fmt:
                ws.write_blank(row, col, "", fmt)

        # 6. Cell formats (overlay formatting on a range)
        for cf in sheet_layout.cell_formats:
            r1, c1, r2, c2 = _parse_range_address(cf.range)
            fmt = wb.add_format(cf.format) if cf.format else None
            if fmt:
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        ws.write_blank(r, c, "", fmt)

        # 7. Merges
        for m in sheet_layout.merges:
            r1, c1, r2, c2 = _parse_range_address(m.range)
            fmt = wb.add_format(m.format) if m.format else None
            ws.merge_range(r1, c1, r2, c2, m.value, fmt)

        # 8. Conditional formats
        for cf in sheet_layout.conditional_formats:
            r1, c1, r2, c2 = _parse_range_address(cf.range)
            ws.conditional_format(r1, c1, r2, c2, cf.params)

        # 9. Data validations
        for dv in sheet_layout.data_validations:
            r1, c1, r2, c2 = _parse_range_address(dv.range)
            ws.data_validation(r1, c1, r2, c2, dv.params)

        # 10. Hyperlinks
        for hl in sheet_layout.hyperlinks:
            row, col = _parse_cell_address(hl.cell)
            ws.write_url(row, col, hl.url, string=hl.display_text or hl.url)

        # 11. Comments
        for cm in sheet_layout.comments:
            row, col = _parse_cell_address(cm.cell)
            ws.write_comment(row, col, cm.text, {"author": cm.author})

        # 12. Images
        for img in sheet_layout.images:
            row, col = _parse_cell_address(img.cell)
            ws.insert_image(row, col, img.image_path, {
                "x_scale": img.x_scale,
                "y_scale": img.y_scale,
            })
