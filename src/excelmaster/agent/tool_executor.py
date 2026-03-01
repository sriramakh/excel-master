"""ToolExecutor — maps tool calls to WorkbookState mutations + registry."""
from __future__ import annotations

import json
from typing import Any

import pandas as pd

from ..models import (
    AggFunc,
    ChartConfig,
    ChartType,
    ColorTheme,
    KPIConfig,
    NumberFormat,
)
from ..chat.layout import LayoutEngine
from ..chat.models import (
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
from .registry import ObjectRegistry, OperationType


def _ok(msg: str, object_id: str = "") -> dict:
    return {"success": True, "message": msg, "object_id": object_id}


def _err(msg: str) -> dict:
    return {"success": False, "message": msg, "object_id": ""}


class ToolExecutor:
    """Executes tool calls against WorkbookState and ObjectRegistry."""

    def __init__(
        self,
        state: WorkbookState,
        registry: ObjectRegistry,
        df: pd.DataFrame | None = None,
        turn: int = 0,
    ) -> None:
        self.state = state
        self.registry = registry
        self.df = df
        self.turn = turn

    def execute(self, tool_name: str, args: dict) -> dict:
        """Dispatch a tool call and return result dict."""
        dispatch = {
            "add_chart": self._add_chart,
            "modify_object": self._modify_object,
            "remove_object": self._remove_object,
            "add_kpi_row": self._add_kpi_row,
            "add_table": self._add_table,
            "add_content": self._add_content,
            "write_cells": self._write_cells,
            "format_range": self._format_range,
            "sheet_operation": self._sheet_operation,
            "row_col_operation": self._row_col_operation,
            "add_excel_feature": self._add_excel_feature,
            "change_theme": self._change_theme,
            "query_workbook": self._query_workbook,
        }
        fn = dispatch.get(tool_name)
        if fn is None:
            return _err(f"Unknown tool: {tool_name}")
        try:
            return fn(args)
        except Exception as e:
            return _err(f"Error executing {tool_name}: {e}")

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _get_sheet(self, args: dict) -> SheetLayout:
        name = args.get("sheet", "Dashboard") or "Dashboard"
        sheet = self.state.get_sheet(name)
        if sheet is None:
            sheet = self.state.add_sheet(name)
        return sheet

    def _validate_columns(self, cols: list[str]) -> list[str]:
        if self.df is None:
            return cols
        valid = set(self.df.columns)
        return [c for c in cols if c in valid]

    # ── 1. add_chart ──────────────────────────────────────────────────────────

    def _add_chart(self, args: dict) -> dict:
        sheet = self._get_sheet(args)

        chart_type_str = args.get("type", "bar")
        try:
            chart_type = ChartType(chart_type_str)
        except ValueError:
            chart_type = ChartType.BAR

        x_col = args.get("x_column", "")
        y_cols = args.get("y_columns", [])
        if isinstance(y_cols, str):
            y_cols = [y_cols]

        if self.df is not None:
            if x_col and x_col not in self.df.columns:
                return _err(f"Column '{x_col}' not found in data")
            y_cols = self._validate_columns(y_cols)
            if not y_cols:
                return _err("No valid y columns found in data")

        agg_str = args.get("aggregation", "sum")
        try:
            agg = AggFunc(agg_str)
        except ValueError:
            agg = AggFunc.SUM

        chart_cfg = ChartConfig(
            type=chart_type,
            title=args.get("title", f"{chart_type.value.replace('_', ' ').title()} Chart"),
            x_column=x_col,
            y_columns=y_cols,
            aggregation=agg,
            top_n=args.get("top_n", 0),
            show_data_labels=args.get("show_data_labels", True),
        )

        width = args.get("width", "half")
        side = args.get("side", "left")
        if width not in ("full", "half"):
            width = "half"
        if side not in ("left", "right"):
            side = "left"

        obj_id = LayoutEngine.generate_id(sheet, ObjectType.CHART)
        obj = PlacedObject(
            id=obj_id,
            type=ObjectType.CHART,
            payload=PlacedChart(chart=chart_cfg, width=width, side=side),
        )
        LayoutEngine.insert_object(sheet, obj, args.get("position", "end"))

        self.registry.register(
            op_type=OperationType.CHART,
            sheet=sheet.name,
            location=f"row {obj.anchor_row}",
            description=f"{chart_type.value} chart: {chart_cfg.title}",
            turn=self.turn,
            params=args,
            entry_id=obj_id,
        )
        return _ok(f"Added {chart_type.value} chart '{chart_cfg.title}'", obj_id)

    # ── 2. modify_object ──────────────────────────────────────────────────────

    def _modify_object(self, args: dict) -> dict:
        obj_id = args.get("object_id", "")
        changes = args.get("changes", {})
        sheet = self._get_sheet(args)

        obj = sheet.find_object(obj_id)
        if obj is None:
            # Search all sheets
            for s in self.state.sheets:
                obj = s.find_object(obj_id)
                if obj:
                    sheet = s
                    break
            if obj is None:
                return _err(f"Object '{obj_id}' not found")

        if obj.type == ObjectType.CHART:
            pc: PlacedChart = obj.payload  # type: ignore[assignment]
            cfg = pc.chart
            if "type" in changes:
                try:
                    cfg.type = ChartType(changes["type"])
                except ValueError:
                    pass
            if "title" in changes:
                cfg.title = changes["title"]
            if "x_column" in changes:
                cfg.x_column = changes["x_column"]
            if "y_columns" in changes:
                y = changes["y_columns"]
                cfg.y_columns = y if isinstance(y, list) else [y]
            if "aggregation" in changes:
                try:
                    cfg.aggregation = AggFunc(changes["aggregation"])
                except ValueError:
                    pass
            if "top_n" in changes:
                cfg.top_n = int(changes["top_n"])
            if "width" in changes:
                pc.width = changes["width"]
            if "side" in changes:
                pc.side = changes["side"]

        elif obj.type == ObjectType.TABLE:
            pt: PlacedTable = obj.payload  # type: ignore[assignment]
            if "columns" in changes:
                cols = changes["columns"]
                pt.columns = self._validate_columns(cols if isinstance(cols, list) else [cols])
            if "max_rows" in changes:
                pt.max_rows = int(changes["max_rows"])
            if "show_conditional" in changes:
                pt.show_conditional = bool(changes["show_conditional"])

        elif obj.type == ObjectType.KPI_ROW:
            pk: PlacedKPIRow = obj.payload  # type: ignore[assignment]
            if "kpis" in changes:
                pk.kpis = self._build_kpis(changes["kpis"])

        elif obj.type == ObjectType.TITLE:
            ptitle: PlacedTitle = obj.payload  # type: ignore[assignment]
            if "text" in changes:
                ptitle.text = changes["text"]
                self.state.title = changes["text"]
            if "subtitle" in changes:
                ptitle.subtitle = changes["subtitle"]

        elif obj.type == ObjectType.SECTION_HEADER:
            psh: PlacedSectionHeader = obj.payload  # type: ignore[assignment]
            if "text" in changes:
                psh.text = changes["text"]
            if "color" in changes:
                psh.color = changes["color"]

        elif obj.type == ObjectType.TEXT:
            ptxt: PlacedText = obj.payload  # type: ignore[assignment]
            if "content" in changes:
                ptxt.content = changes["content"]
            if "style" in changes:
                ptxt.style = changes["style"]

        elif obj.type == ObjectType.PIVOT:
            ppiv: PlacedPivot = obj.payload  # type: ignore[assignment]
            if "index_col" in changes:
                ppiv.index_col = changes["index_col"]
            if "value_col" in changes:
                ppiv.value_col = changes["value_col"]
            if "columns_col" in changes:
                ppiv.columns_col = changes["columns_col"]
            if "agg" in changes:
                try:
                    ppiv.agg = AggFunc(changes["agg"])
                except ValueError:
                    pass

        # Update registry entry description
        entry = self.registry.get(obj_id)
        if entry:
            entry.description = f"(modified) {entry.description}"

        return _ok(f"Modified {obj.type.value} '{obj_id}'", obj_id)

    # ── 3. remove_object ──────────────────────────────────────────────────────

    def _remove_object(self, args: dict) -> dict:
        obj_id = args.get("object_id", "")

        for s in self.state.sheets:
            removed = s.remove_object(obj_id)
            if removed:
                LayoutEngine.reflow(s)
                self.registry.remove(obj_id)
                return _ok(f"Removed {removed.type.value} '{obj_id}'", obj_id)

        return _err(f"Object '{obj_id}' not found")

    # ── 4. add_kpi_row ────────────────────────────────────────────────────────

    def _build_kpis(self, kpis_raw: list[dict]) -> list[KPIConfig]:
        kpis = []
        for k in kpis_raw:
            col = k.get("column", "")
            if self.df is not None and col and col not in self.df.columns:
                continue
            try:
                agg = AggFunc(k.get("aggregation", "sum"))
            except ValueError:
                agg = AggFunc.SUM
            try:
                fmt = NumberFormat(k.get("format", "number"))
            except ValueError:
                fmt = NumberFormat.NUMBER
            kpis.append(KPIConfig(
                label=k.get("label", col),
                column=col,
                aggregation=agg,
                format=fmt,
                prefix=k.get("prefix", ""),
                suffix=k.get("suffix", ""),
                icon=k.get("icon", ""),
                trend_column=k.get("trend_column", ""),
            ))
        return kpis

    def _add_kpi_row(self, args: dict) -> dict:
        sheet = self._get_sheet(args)
        kpis = self._build_kpis(args.get("kpis", []))
        if not kpis:
            return _err("No valid KPIs provided")

        payload = PlacedKPIRow(kpis=kpis)
        obj_id = LayoutEngine.generate_id(sheet, ObjectType.KPI_ROW)
        obj = PlacedObject(id=obj_id, type=ObjectType.KPI_ROW, payload=payload)
        LayoutEngine.insert_object(sheet, obj, args.get("position", "end"))

        labels = [k.label for k in kpis]
        self.registry.register(
            op_type=OperationType.KPI_ROW,
            sheet=sheet.name,
            location=f"row {obj.anchor_row}",
            description=f"KPI row: {', '.join(labels)}",
            turn=self.turn,
            params=args,
            entry_id=obj_id,
        )
        return _ok(f"Added KPI row with {len(kpis)} tiles", obj_id)

    # ── 5. add_table ──────────────────────────────────────────────────────────

    def _add_table(self, args: dict) -> dict:
        sheet = self._get_sheet(args)
        table_type = args.get("table_type", "data")

        if table_type == "pivot":
            idx = args.get("index_col", "")
            val = args.get("value_col", "")
            if self.df is not None:
                if idx and idx not in self.df.columns:
                    return _err(f"Column '{idx}' not found for pivot index")
                if val and val not in self.df.columns:
                    return _err(f"Column '{val}' not found for pivot values")

            try:
                agg = AggFunc(args.get("agg", "sum"))
            except ValueError:
                agg = AggFunc.SUM

            payload_pivot = PlacedPivot(
                index_col=idx,
                value_col=val,
                columns_col=args.get("columns_col", ""),
                agg=agg,
            )
            obj_id = LayoutEngine.generate_id(sheet, ObjectType.PIVOT)
            obj = PlacedObject(id=obj_id, type=ObjectType.PIVOT, payload=payload_pivot)
            LayoutEngine.insert_object(sheet, obj, args.get("position", "end"))

            self.registry.register(
                op_type=OperationType.PIVOT,
                sheet=sheet.name,
                location=f"row {obj.anchor_row}",
                description=f"Pivot: {idx} x {val}",
                turn=self.turn,
                params=args,
                entry_id=obj_id,
            )
            return _ok(f"Added pivot table", obj_id)

        # Data table
        cols = args.get("columns", [])
        if isinstance(cols, str):
            cols = [cols]
        cols = self._validate_columns(cols) if cols else []

        payload_table = PlacedTable(
            columns=cols,
            max_rows=args.get("max_rows", 15),
            show_conditional=args.get("show_conditional", True),
        )
        obj_id = LayoutEngine.generate_id(sheet, ObjectType.TABLE)
        obj = PlacedObject(id=obj_id, type=ObjectType.TABLE, payload=payload_table)
        LayoutEngine.insert_object(sheet, obj, args.get("position", "end"))

        self.registry.register(
            op_type=OperationType.TABLE,
            sheet=sheet.name,
            location=f"row {obj.anchor_row}",
            description=f"Data table ({len(cols)} columns, {payload_table.max_rows} rows)",
            turn=self.turn,
            params=args,
            entry_id=obj_id,
        )
        return _ok(f"Added data table", obj_id)

    # ── 6. add_content ────────────────────────────────────────────────────────

    def _add_content(self, args: dict) -> dict:
        sheet = self._get_sheet(args)
        content_type = args.get("content_type", "text")
        text = args.get("text", "")

        if content_type == "title":
            payload: Any = PlacedTitle(
                text=text,
                subtitle=args.get("subtitle", ""),
            )
            obj_type = ObjectType.TITLE
            self.state.title = text

        elif content_type == "section_header":
            payload = PlacedSectionHeader(
                text=text,
                color=args.get("color", ""),
            )
            obj_type = ObjectType.SECTION_HEADER

        else:  # text
            payload = PlacedText(
                content=text,
                style=args.get("style", "body"),
            )
            obj_type = ObjectType.TEXT

        obj_id = LayoutEngine.generate_id(sheet, obj_type)
        obj = PlacedObject(id=obj_id, type=obj_type, payload=payload)
        LayoutEngine.insert_object(sheet, obj, args.get("position", "end"))

        op = OperationType.TITLE if content_type == "title" else (
            OperationType.SECTION_HEADER if content_type == "section_header"
            else OperationType.TEXT
        )
        self.registry.register(
            op_type=op,
            sheet=sheet.name,
            location=f"row {obj.anchor_row}",
            description=f"{content_type}: {text[:50]}",
            turn=self.turn,
            params=args,
            entry_id=obj_id,
        )
        return _ok(f"Added {content_type}", obj_id)

    # ── 7. write_cells ────────────────────────────────────────────────────────

    def _write_cells(self, args: dict) -> dict:
        sheet = self._get_sheet(args)
        writes = args.get("writes", [])
        ids = []

        for w in writes:
            cell = w.get("cell", "A1")
            value = w.get("value")
            fmt_dict: dict[str, Any] = {}
            for key in ("bold", "italic", "font_size", "font_color", "bg_color",
                        "num_format", "align", "border"):
                if key in w:
                    fmt_dict[key] = w[key]

            cw = CellWrite(cell=cell, value=value, format=fmt_dict)
            sheet.cell_writes.append(cw)

            rid = self.registry.register(
                op_type=OperationType.CELL_WRITE,
                sheet=sheet.name,
                location=cell,
                description=f"Write '{value}' to {cell}",
                turn=self.turn,
                params=w,
            )
            ids.append(rid)

        return _ok(f"Wrote {len(writes)} cell(s)", ", ".join(ids))

    # ── 8. format_range ───────────────────────────────────────────────────────

    def _format_range(self, args: dict) -> dict:
        rng = args.get("range", "A1")
        sheet_name = args.get("sheet", "Dashboard") or "Dashboard"
        sheet = self._get_sheet(args)

        fmt_dict: dict[str, Any] = {}
        for key in ("bold", "italic", "font_size", "font_color", "bg_color",
                    "num_format", "align", "valign", "border", "text_wrap"):
            if key in args:
                fmt_dict[key] = args[key]

        cf = CellFormatOp(range=rng, format=fmt_dict)
        sheet.cell_formats.append(cf)

        rid = self.registry.register(
            op_type=OperationType.CELL_FORMAT,
            sheet=sheet.name,
            location=rng,
            description=f"Format range {rng}",
            turn=self.turn,
            params=args,
        )
        return _ok(f"Formatted range {rng}", rid)

    # ── 9. sheet_operation ────────────────────────────────────────────────────

    def _sheet_operation(self, args: dict) -> dict:
        op = args.get("operation", "")
        sheet_name = args.get("sheet", "")

        if op == "create":
            name = sheet_name or "Sheet2"
            self.state.add_sheet(name)
            rid = self.registry.register(
                op_type=OperationType.SHEET,
                sheet=name,
                description=f"Created sheet '{name}'",
                turn=self.turn,
                params=args,
            )
            return _ok(f"Created sheet '{name}'", rid)

        elif op == "rename":
            new_name = args.get("new_name", "")
            if not new_name:
                return _err("new_name required for rename")
            s = self.state.get_sheet(sheet_name)
            if s is None:
                return _err(f"Sheet '{sheet_name}' not found")
            s.name = new_name
            return _ok(f"Renamed '{sheet_name}' to '{new_name}'")

        elif op == "delete":
            if not sheet_name:
                return _err("sheet name required")
            if sheet_name == "Dashboard":
                return _err("Cannot delete the Dashboard sheet")
            self.state.sheets = [s for s in self.state.sheets if s.name != sheet_name]
            return _ok(f"Deleted sheet '{sheet_name}'")

        elif op == "reorder":
            pos = args.get("position", 0)
            s = self.state.get_sheet(sheet_name)
            if s is None:
                return _err(f"Sheet '{sheet_name}' not found")
            self.state.sheets.remove(s)
            self.state.sheets.insert(pos, s)
            return _ok(f"Moved sheet '{sheet_name}' to position {pos}")

        elif op == "set_tab_color":
            color = args.get("tab_color", "")
            s = self.state.get_sheet(sheet_name)
            if s is None:
                return _err(f"Sheet '{sheet_name}' not found")
            s.tab_color = color
            return _ok(f"Set tab color of '{sheet_name}' to {color}")

        elif op == "hide":
            s = self.state.get_sheet(sheet_name)
            if s is None:
                return _err(f"Sheet '{sheet_name}' not found")
            s.hidden = True
            return _ok(f"Hidden sheet '{sheet_name}'")

        elif op == "show":
            s = self.state.get_sheet(sheet_name)
            if s is None:
                return _err(f"Sheet '{sheet_name}' not found")
            s.hidden = False
            return _ok(f"Shown sheet '{sheet_name}'")

        return _err(f"Unknown sheet operation: {op}")

    # ── 10. row_col_operation ─────────────────────────────────────────────────

    def _row_col_operation(self, args: dict) -> dict:
        sheet = self._get_sheet(args)
        target = args.get("target", "row")
        op = args.get("operation", "")
        idx = args.get("index", 0)
        end_idx = args.get("end_index", idx)

        if op == "resize":
            size = args.get("size", 15)
            if target == "row":
                for r in range(idx, end_idx + 1):
                    sheet.row_heights[r] = size
            else:
                for c in range(idx, end_idx + 1):
                    sheet.col_widths[c] = size
            return _ok(f"Resized {target}(s) {idx}-{end_idx} to {size}")

        elif op == "hide":
            if target == "row":
                for r in range(idx, end_idx + 1):
                    if r not in sheet.hidden_rows:
                        sheet.hidden_rows.append(r)
            else:
                for c in range(idx, end_idx + 1):
                    if c not in sheet.hidden_cols:
                        sheet.hidden_cols.append(c)
            return _ok(f"Hidden {target}(s) {idx}-{end_idx}")

        elif op == "show":
            if target == "row":
                sheet.hidden_rows = [r for r in sheet.hidden_rows if r < idx or r > end_idx]
            else:
                sheet.hidden_cols = [c for c in sheet.hidden_cols if c < idx or c > end_idx]
            return _ok(f"Shown {target}(s) {idx}-{end_idx}")

        return _err(f"Unknown row/col operation: {op}")

    # ── 11. add_excel_feature ─────────────────────────────────────────────────

    def _add_excel_feature(self, args: dict) -> dict:
        sheet = self._get_sheet(args)
        feature = args.get("feature", "")

        if feature == "conditional_format":
            rng = args.get("range", "A1:A10")
            rule_type = args.get("rule_type", "3_color_scale")
            params: dict[str, Any] = {"type": rule_type}

            if rule_type in ("3_color_scale", "2_color_scale"):
                if args.get("min_color"):
                    params["min_color"] = args["min_color"]
                if args.get("mid_color"):
                    params["mid_color"] = args["mid_color"]
                if args.get("max_color"):
                    params["max_color"] = args["max_color"]
            elif rule_type == "data_bar":
                if args.get("bar_color"):
                    params["bar_color"] = args["bar_color"]
            elif rule_type == "cell_is":
                params["criteria"] = args.get("criteria", ">")
                params["value"] = args.get("value", 0)

            cf = ConditionalFormatOp(range=rng, rule_type=rule_type, params=params)
            sheet.conditional_formats.append(cf)

            rid = self.registry.register(
                op_type=OperationType.CONDITIONAL_FORMAT,
                sheet=sheet.name,
                location=rng,
                description=f"Conditional format ({rule_type}) on {rng}",
                turn=self.turn,
                params=args,
            )
            return _ok(f"Added {rule_type} conditional format on {rng}", rid)

        elif feature == "data_validation":
            rng = args.get("range", "A1")
            validate = args.get("validate", "list")
            dv_params: dict[str, Any] = {"validate": validate}
            if validate == "list" and "source" in args:
                dv_params["source"] = args["source"]

            dv = DataValidationOp(range=rng, validation_type=validate, params=dv_params)
            sheet.data_validations.append(dv)

            rid = self.registry.register(
                op_type=OperationType.DATA_VALIDATION,
                sheet=sheet.name,
                location=rng,
                description=f"Data validation ({validate}) on {rng}",
                turn=self.turn,
                params=args,
            )
            return _ok(f"Added data validation on {rng}", rid)

        elif feature == "freeze_panes":
            sheet.freeze_row = args.get("freeze_row", sheet.freeze_row)
            sheet.freeze_col = args.get("freeze_col", sheet.freeze_col)
            return _ok(f"Freeze panes set to row={sheet.freeze_row}, col={sheet.freeze_col}")

        elif feature == "zoom":
            level = args.get("zoom_level", 100)
            sheet.zoom = max(10, min(400, level))
            return _ok(f"Zoom set to {sheet.zoom}%")

        elif feature == "merge":
            rng = args.get("range", "A1:B1")
            value = args.get("merge_value", "")
            fmt = args.get("format", {})
            m = MergeOp(range=rng, value=value, format=fmt)
            sheet.merges.append(m)

            rid = self.registry.register(
                op_type=OperationType.MERGE,
                sheet=sheet.name,
                location=rng,
                description=f"Merge {rng}: '{value}'",
                turn=self.turn,
                params=args,
            )
            return _ok(f"Merged {rng}", rid)

        elif feature == "hyperlink":
            cell = args.get("cell", "A1")
            url = args.get("url", "")
            display = args.get("display_text", url)
            hl = HyperlinkOp(cell=cell, url=url, display_text=display)
            sheet.hyperlinks.append(hl)

            rid = self.registry.register(
                op_type=OperationType.HYPERLINK,
                sheet=sheet.name,
                location=cell,
                description=f"Hyperlink at {cell}: {url}",
                turn=self.turn,
                params=args,
            )
            return _ok(f"Added hyperlink at {cell}", rid)

        elif feature == "comment":
            cell = args.get("cell", "A1")
            text = args.get("comment_text", "")
            author = args.get("author", "Excel Master")
            cm = CommentOp(cell=cell, text=text, author=author)
            sheet.comments.append(cm)

            rid = self.registry.register(
                op_type=OperationType.COMMENT,
                sheet=sheet.name,
                location=cell,
                description=f"Comment at {cell}: '{text[:30]}'",
                turn=self.turn,
                params=args,
            )
            return _ok(f"Added comment at {cell}", rid)

        elif feature == "image":
            cell = args.get("cell", "A1")
            image_path = args.get("image_path", "")
            img = ImageOp(
                cell=cell,
                image_path=image_path,
                x_scale=args.get("x_scale", 1.0),
                y_scale=args.get("y_scale", 1.0),
            )
            sheet.images.append(img)

            rid = self.registry.register(
                op_type=OperationType.IMAGE,
                sheet=sheet.name,
                location=cell,
                description=f"Image at {cell}: {image_path}",
                turn=self.turn,
                params=args,
            )
            return _ok(f"Added image at {cell}", rid)

        return _err(f"Unknown feature: {feature}")

    # ── 12. change_theme ──────────────────────────────────────────────────────

    def _change_theme(self, args: dict) -> dict:
        theme = args.get("theme", "corporate_blue")
        try:
            ColorTheme(theme)
        except ValueError:
            return _err(f"Unknown theme: {theme}")

        self.state.theme_key = theme
        self.registry.register(
            op_type=OperationType.THEME,
            description=f"Changed theme to {theme}",
            turn=self.turn,
            params=args,
        )
        return _ok(f"Changed theme to '{theme}'")

    # ── 13. query_workbook ────────────────────────────────────────────────────

    def _query_workbook(self, args: dict) -> dict:
        query = args.get("query", "list_objects")
        sheet_filter = args.get("sheet")

        if query == "list_objects":
            lines = []
            for s in self.state.sheets:
                if sheet_filter and s.name != sheet_filter:
                    continue
                lines.append(f"Sheet: {s.name}")
                if not s.objects:
                    lines.append("  (no objects)")
                for obj in s.sorted_objects():
                    from ..chat.prompts import _describe_object
                    lines.append(f"  [{obj.id}] row {obj.anchor_row}: {_describe_object(obj)}")
            return _ok("\n".join(lines))

        elif query == "object_details":
            obj_id = args.get("object_id", "")
            for s in self.state.sheets:
                obj = s.find_object(obj_id)
                if obj:
                    return _ok(json.dumps(obj.model_dump(), default=str, indent=2))
            return _err(f"Object '{obj_id}' not found")

        elif query == "data_summary":
            if self.df is None:
                return _ok("No data loaded")
            info = {
                "rows": len(self.df),
                "columns": len(self.df.columns),
                "column_names": list(self.df.columns),
                "dtypes": {c: str(self.df[c].dtype) for c in self.df.columns[:20]},
                "sample": self.df.head(3).to_dict(orient="records"),
            }
            return _ok(json.dumps(info, default=str, indent=2))

        elif query == "list_sheets":
            sheets = [{"name": s.name, "objects": len(s.objects), "hidden": s.hidden}
                      for s in self.state.sheets]
            return _ok(json.dumps(sheets, indent=2))

        elif query == "registry_snapshot":
            return _ok(self.registry.to_snapshot())

        return _err(f"Unknown query: {query}")
