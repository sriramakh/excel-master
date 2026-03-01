"""ChatEngine — interactive REPL for building dashboards via natural language."""
from __future__ import annotations

import traceback
from pathlib import Path
from typing import Any

import pandas as pd

from ..config import get_settings
from ..data.data_engine import profile_dataset, discover_and_join
from ..dashboard.llm_client import LLMClient
from ..dashboard.template_selector import TemplateSelector
from ..models import (
    AggFunc,
    ChartConfig,
    ChartType,
    ColorTheme,
    DashboardConfig,
    KPIConfig,
    NumberFormat,
)
from .layout import LayoutEngine
from .models import (
    ActionType,
    ChatAction,
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
from .prompts import build_system_prompt, build_user_message
from .renderer import FlexibleTemplate

MAX_UNDO = 30
MAX_HISTORY = 20


class ChatEngine:
    """Interactive chat engine that builds/modifies dashboards."""

    def __init__(self, data_path: Path, output_path: Path | None = None):
        self.data_path = Path(data_path)
        self.output_path = output_path or self._default_output(self.data_path)
        self.df: pd.DataFrame | None = None
        self.profile = None
        self.state = WorkbookState()
        self.undo_stack: list[WorkbookState] = []
        self.redo_stack: list[WorkbookState] = []
        self.messages: list[dict[str, str]] = []  # conversation history
        self.llm = LLMClient()
        self._system_prompt = ""

    @staticmethod
    def _default_output(data_path: Path) -> Path:
        stem = data_path.stem
        return data_path.parent.parent / "output" / f"{stem}_chat_dashboard.xlsx"

    # ── Data Loading ─────────────────────────────────────────────────────────

    def _load_data(self) -> None:
        path = self.data_path
        if path.suffix.lower() == ".csv":
            self.df = pd.read_csv(path)
        else:
            xf = pd.ExcelFile(path)
            if len(xf.sheet_names) > 1:
                self.df, _, _ = discover_and_join(path, verbose=False)
            else:
                self.df = pd.read_excel(path)

        self.profile = profile_dataset(str(path))
        self._system_prompt = build_system_prompt(self.profile)
        self.messages = [{"role": "system", "content": self._system_prompt}]

    # ── REPL ─────────────────────────────────────────────────────────────────

    def run(self) -> None:
        try:
            from rich.console import Console
            from rich.panel import Panel
            from rich.table import Table
            console = Console()
        except ImportError:
            console = None

        self._load_data()

        # Welcome message
        info = (
            f"File: {self.data_path.name}\n"
            f"Shape: {self.df.shape[0]:,} rows x {self.df.shape[1]} columns\n"
            f"Columns: {', '.join(self.df.columns[:12])}"
            + (f" ... +{len(self.df.columns)-12} more" if len(self.df.columns) > 12 else "")
        )
        if console:
            console.print(Panel(info, title="Excel Master Chat", border_style="blue"))
            console.print(
                "[bold]Commands:[/bold] [cyan]auto[/cyan] (full dashboard) | "
                "[cyan]start[/cyan] (empty canvas) | or type an instruction\n"
                "Special: [cyan]undo[/cyan] | [cyan]redo[/cyan] | "
                "[cyan]show[/cyan] | [cyan]reset[/cyan] | "
                "[cyan]save as <name>[/cyan] | [cyan]quit[/cyan]\n"
            )
        else:
            print(f"\n=== Excel Master Chat ===\n{info}")
            print("Commands: auto | start | undo | redo | show | reset | save as <name> | quit\n")

        while True:
            try:
                user_input = input("You: ").strip()
            except (EOFError, KeyboardInterrupt):
                print("\nGoodbye!")
                break

            if not user_input:
                continue

            low = user_input.lower()

            # Special commands
            if low in ("quit", "exit", "q"):
                print("Goodbye!")
                break
            elif low == "undo":
                self._undo()
                continue
            elif low == "redo":
                self._redo()
                continue
            elif low == "show":
                self._show_state(console)
                continue
            elif low == "reset":
                self._push_undo()
                self.state = WorkbookState()
                self.redo_stack.clear()
                print("Dashboard reset to empty state.")
                continue
            elif low.startswith("save as "):
                name = user_input[8:].strip()
                if name:
                    new_path = self.output_path.parent / f"{name}.xlsx"
                    self.output_path = new_path
                    self._render_and_save(console)
                    print(f"Saved as: {new_path}")
                continue
            elif low == "start":
                self.state = WorkbookState(title="Dashboard")
                print("Starting with empty canvas. Tell me what to add!")
                continue
            elif low == "auto":
                self._push_undo()
                try:
                    self._auto_dashboard()
                    self._render_and_save(console)
                    self._print_summary(console)
                except Exception as e:
                    print(f"Error building auto dashboard: {e}")
                    traceback.print_exc()
                continue

            # Normal LLM-driven turn
            self._push_undo()
            try:
                result = self._llm_turn(user_input)
                if result:
                    self._render_and_save(console)
                    self._print_summary(console)
            except Exception as e:
                print(f"Error: {e}")
                traceback.print_exc()
                # Undo the failed turn
                if self.undo_stack:
                    self.state = self.undo_stack.pop()

    # ── LLM Turn ─────────────────────────────────────────────────────────────

    def _llm_turn(self, user_input: str) -> bool:
        """Send user input to LLM, parse response, execute actions. Returns True if state changed."""
        user_msg = build_user_message(self.state, user_input)
        self.messages.append({"role": "user", "content": user_msg})

        # Truncate history to keep last N turns (keep system prompt)
        self._trim_history()

        try:
            result = self.llm.generate_chat_json(self.messages, max_tokens_override=4096)
        except Exception as e:
            print(f"LLM error: {e}")
            self.messages.pop()  # remove the failed user message
            return False

        msg = result.get("message", "")
        actions_raw = result.get("actions", [])

        # Store assistant message for history
        self.messages.append({"role": "assistant", "content": msg})

        if msg:
            print(f"Assistant: {msg}")

        if not actions_raw:
            return False

        # Parse and execute actions
        changed = False
        for action_dict in actions_raw:
            try:
                action = self._parse_action(action_dict)
                summary = self._execute_action(action)
                if summary:
                    print(f"  -> {summary}")
                    changed = True
            except Exception as e:
                print(f"  -> Skipped action: {e}")

        return changed

    def _trim_history(self) -> None:
        """Keep system prompt + last MAX_HISTORY user/assistant turns."""
        if len(self.messages) <= 1 + MAX_HISTORY * 2:
            return
        system = self.messages[0]
        recent = self.messages[-(MAX_HISTORY * 2):]
        self.messages = [system] + recent

    # ── Action Parsing ───────────────────────────────────────────────────────

    def _parse_action(self, raw: dict) -> ChatAction:
        action_str = raw.get("action", "")
        try:
            action_type = ActionType(action_str)
        except ValueError:
            raise ValueError(f"Unknown action type: {action_str}")
        return ChatAction(
            action=action_type,
            target_sheet=raw.get("target_sheet", "Dashboard"),
            target_id=raw.get("target_id", ""),
            params=raw.get("params", {}),
        )

    # ── Action Execution ─────────────────────────────────────────────────────

    def _execute_action(self, action: ChatAction) -> str:
        dispatch = {
            ActionType.ADD_CHART: self._exec_add_chart,
            ActionType.MODIFY_CHART: self._exec_modify_chart,
            ActionType.ADD_TABLE: self._exec_add_table,
            ActionType.MODIFY_TABLE: self._exec_modify_table,
            ActionType.ADD_KPI_ROW: self._exec_add_kpi_row,
            ActionType.MODIFY_KPI: self._exec_modify_kpi,
            ActionType.ADD_PIVOT: self._exec_add_pivot,
            ActionType.ADD_SECTION_HEADER: self._exec_add_section_header,
            ActionType.ADD_TEXT: self._exec_add_text,
            ActionType.REMOVE: self._exec_remove,
            ActionType.MOVE: self._exec_move,
            ActionType.ADD_SHEET: self._exec_add_sheet,
            ActionType.CHANGE_THEME: self._exec_change_theme,
            ActionType.CHANGE_TITLE: self._exec_change_title,
            ActionType.AUTO_DASHBOARD: lambda a: self._auto_dashboard() or "Auto dashboard built",
        }
        fn = dispatch.get(action.action)
        if fn is None:
            return f"Unknown action: {action.action}"
        return fn(action)

    def _get_sheet(self, action: ChatAction) -> SheetLayout:
        name = action.target_sheet or "Dashboard"
        sheet = self.state.get_sheet(name)
        if sheet is None:
            sheet = self.state.add_sheet(name)
        return sheet

    def _validate_columns(self, cols: list[str]) -> list[str]:
        """Filter to only valid column names in the dataframe."""
        if self.df is None:
            return cols
        valid = set(self.df.columns)
        return [c for c in cols if c in valid]

    def _exec_add_chart(self, action: ChatAction) -> str:
        p = action.params
        sheet = self._get_sheet(action)

        chart_type_str = p.get("type", "bar")
        try:
            chart_type = ChartType(chart_type_str)
        except ValueError:
            chart_type = ChartType.BAR

        x_col = p.get("x_column", "")
        y_cols = p.get("y_columns", [])
        if isinstance(y_cols, str):
            y_cols = [y_cols]

        # Validate columns
        if self.df is not None:
            if x_col not in self.df.columns:
                return f"Column '{x_col}' not found, skipping chart"
            y_cols = self._validate_columns(y_cols)
            if not y_cols:
                return "No valid y columns, skipping chart"

        agg_str = p.get("aggregation", "sum")
        try:
            agg = AggFunc(agg_str)
        except ValueError:
            agg = AggFunc.SUM

        chart_cfg = ChartConfig(
            type=chart_type,
            title=p.get("title", f"{chart_type.value.title()} Chart"),
            x_column=x_col,
            y_columns=y_cols,
            aggregation=agg,
            top_n=p.get("top_n", 0),
            show_data_labels=p.get("show_data_labels", True),
        )

        width = p.get("width", "half")
        side = p.get("side", "left")
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
        LayoutEngine.insert_object(sheet, obj, p.get("position", "end"))
        return f"Added {chart_type.value} chart '{chart_cfg.title}' [{obj_id}]"

    def _exec_modify_chart(self, action: ChatAction) -> str:
        sheet = self._get_sheet(action)
        obj = sheet.find_object(action.target_id)
        if obj is None or obj.type != ObjectType.CHART:
            return f"Chart '{action.target_id}' not found"

        p = action.params
        pc: PlacedChart = obj.payload  # type: ignore[assignment]
        cfg = pc.chart

        if "type" in p:
            try:
                cfg.type = ChartType(p["type"])
            except ValueError:
                pass
        if "title" in p:
            cfg.title = p["title"]
        if "x_column" in p:
            cfg.x_column = p["x_column"]
        if "y_columns" in p:
            y = p["y_columns"]
            cfg.y_columns = y if isinstance(y, list) else [y]
        if "aggregation" in p:
            try:
                cfg.aggregation = AggFunc(p["aggregation"])
            except ValueError:
                pass
        if "top_n" in p:
            cfg.top_n = int(p["top_n"])
        if "width" in p:
            pc.width = p["width"]
        if "side" in p:
            pc.side = p["side"]

        return f"Modified chart '{action.target_id}'"

    def _exec_add_table(self, action: ChatAction) -> str:
        p = action.params
        sheet = self._get_sheet(action)

        cols = p.get("columns", [])
        if isinstance(cols, str):
            cols = [cols]
        cols = self._validate_columns(cols) if cols else []

        payload = PlacedTable(
            columns=cols,
            max_rows=p.get("max_rows", 15),
            show_conditional=p.get("show_conditional", True),
        )
        obj_id = LayoutEngine.generate_id(sheet, ObjectType.TABLE)
        obj = PlacedObject(id=obj_id, type=ObjectType.TABLE, payload=payload)
        LayoutEngine.insert_object(sheet, obj)
        return f"Added table [{obj_id}]"

    def _exec_modify_table(self, action: ChatAction) -> str:
        sheet = self._get_sheet(action)
        obj = sheet.find_object(action.target_id)
        if obj is None or obj.type != ObjectType.TABLE:
            return f"Table '{action.target_id}' not found"

        p = action.params
        pt: PlacedTable = obj.payload  # type: ignore[assignment]

        if "columns" in p:
            cols = p["columns"]
            pt.columns = self._validate_columns(cols if isinstance(cols, list) else [cols])
        if "max_rows" in p:
            pt.max_rows = int(p["max_rows"])
        if "show_conditional" in p:
            pt.show_conditional = bool(p["show_conditional"])

        return f"Modified table '{action.target_id}'"

    def _exec_add_kpi_row(self, action: ChatAction) -> str:
        p = action.params
        sheet = self._get_sheet(action)

        kpis_raw = p.get("kpis", [])
        kpis = []
        for k in kpis_raw:
            col = k.get("column", "")
            if self.df is not None and col not in self.df.columns:
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

        if not kpis:
            return "No valid KPIs, skipping"

        payload = PlacedKPIRow(kpis=kpis)
        obj_id = LayoutEngine.generate_id(sheet, ObjectType.KPI_ROW)
        obj = PlacedObject(id=obj_id, type=ObjectType.KPI_ROW, payload=payload)
        LayoutEngine.insert_object(sheet, obj)
        return f"Added KPI row with {len(kpis)} tiles [{obj_id}]"

    def _exec_modify_kpi(self, action: ChatAction) -> str:
        sheet = self._get_sheet(action)
        obj = sheet.find_object(action.target_id)
        if obj is None or obj.type != ObjectType.KPI_ROW:
            return f"KPI row '{action.target_id}' not found"

        p = action.params
        kpis_raw = p.get("kpis", [])
        if kpis_raw:
            kpis = []
            for k in kpis_raw:
                col = k.get("column", "")
                if self.df is not None and col not in self.df.columns:
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
            pk: PlacedKPIRow = obj.payload  # type: ignore[assignment]
            pk.kpis = kpis

        return f"Modified KPI row '{action.target_id}'"

    def _exec_add_pivot(self, action: ChatAction) -> str:
        p = action.params
        sheet = self._get_sheet(action)

        idx = p.get("index_col", "")
        val = p.get("value_col", "")
        if self.df is not None:
            if idx not in self.df.columns:
                return f"Column '{idx}' not found for pivot index"
            if val not in self.df.columns:
                return f"Column '{val}' not found for pivot values"

        try:
            agg = AggFunc(p.get("agg", "sum"))
        except ValueError:
            agg = AggFunc.SUM

        payload = PlacedPivot(
            index_col=idx,
            value_col=val,
            columns_col=p.get("columns_col", ""),
            agg=agg,
        )
        obj_id = LayoutEngine.generate_id(sheet, ObjectType.PIVOT)
        obj = PlacedObject(id=obj_id, type=ObjectType.PIVOT, payload=payload)
        LayoutEngine.insert_object(sheet, obj)
        return f"Added pivot table [{obj_id}]"

    def _exec_add_section_header(self, action: ChatAction) -> str:
        p = action.params
        sheet = self._get_sheet(action)

        payload = PlacedSectionHeader(
            text=p.get("text", "Section"),
            color=p.get("color", ""),
        )
        obj_id = LayoutEngine.generate_id(sheet, ObjectType.SECTION_HEADER)
        obj = PlacedObject(id=obj_id, type=ObjectType.SECTION_HEADER, payload=payload)
        LayoutEngine.insert_object(sheet, obj)
        return f"Added section header '{payload.text}' [{obj_id}]"

    def _exec_add_text(self, action: ChatAction) -> str:
        p = action.params
        sheet = self._get_sheet(action)

        payload = PlacedText(
            content=p.get("content", ""),
            style=p.get("style", "body"),
        )
        obj_id = LayoutEngine.generate_id(sheet, ObjectType.TEXT)
        obj = PlacedObject(id=obj_id, type=ObjectType.TEXT, payload=payload)
        LayoutEngine.insert_object(sheet, obj)
        return f"Added text block [{obj_id}]"

    def _exec_remove(self, action: ChatAction) -> str:
        sheet = self._get_sheet(action)
        removed = sheet.remove_object(action.target_id)
        if removed is None:
            # Try all sheets
            for s in self.state.sheets:
                removed = s.remove_object(action.target_id)
                if removed:
                    LayoutEngine.reflow(s)
                    return f"Removed {removed.type.value} '{action.target_id}'"
            return f"Object '{action.target_id}' not found"
        LayoutEngine.reflow(sheet)
        return f"Removed {removed.type.value} '{action.target_id}'"

    def _exec_move(self, action: ChatAction) -> str:
        sheet = self._get_sheet(action)
        obj = sheet.find_object(action.target_id)
        if obj is None:
            return f"Object '{action.target_id}' not found"

        # Remove and re-insert at new position
        sheet.remove_object(action.target_id)
        position = action.params.get("position", "end")
        LayoutEngine.insert_object(sheet, obj, position)
        LayoutEngine.reflow(sheet)
        return f"Moved '{action.target_id}' to {position}"

    def _exec_add_sheet(self, action: ChatAction) -> str:
        name = action.params.get("name", "Sheet2")
        self.state.add_sheet(name)
        return f"Added sheet '{name}'"

    def _exec_change_theme(self, action: ChatAction) -> str:
        theme = action.params.get("theme", "corporate_blue")
        try:
            ColorTheme(theme)
        except ValueError:
            return f"Unknown theme: {theme}"
        self.state.theme_key = theme
        return f"Changed theme to '{theme}'"

    def _exec_change_title(self, action: ChatAction) -> str:
        title = action.params.get("title", "Dashboard")
        self.state.title = title

        # Update the title object if it exists
        dash = self.state.dashboard_sheet()
        for obj in dash.objects:
            if obj.type == ObjectType.TITLE:
                pt: PlacedTitle = obj.payload  # type: ignore[assignment]
                pt.text = title
                if "subtitle" in action.params:
                    pt.subtitle = action.params["subtitle"]
                break
        return f"Changed title to '{title}'"

    # ── Auto Dashboard ───────────────────────────────────────────────────────

    def _auto_dashboard(self) -> None:
        """Build a full dashboard using the existing LLM template selector."""
        selector = TemplateSelector()
        config: DashboardConfig = selector.select(self.profile)
        self.state = self._config_to_state(config)

    def _config_to_state(self, config: DashboardConfig) -> WorkbookState:
        """Convert a DashboardConfig into a WorkbookState."""
        state = WorkbookState(
            title=config.title,
            theme_key=config.theme.value,
        )
        sheet = state.dashboard_sheet()
        le = LayoutEngine

        # Title
        title_obj = PlacedObject(
            id=le.generate_id(sheet, ObjectType.TITLE),
            type=ObjectType.TITLE,
            payload=PlacedTitle(text=config.title, subtitle=config.subtitle),
        )
        le.insert_object(sheet, title_obj)

        # Filter panel
        filter_cols = [f.column for f in config.filters if f.column]
        if not filter_cols and config.primary_dimension:
            filter_cols = [config.primary_dimension]
        if filter_cols:
            fp_obj = PlacedObject(
                id=le.generate_id(sheet, ObjectType.FILTER_PANEL),
                type=ObjectType.FILTER_PANEL,
                payload=PlacedFilterPanel(filter_columns=filter_cols[:3]),
            )
            le.insert_object(sheet, fp_obj)

        # KPI row
        if config.kpis:
            kpi_obj = PlacedObject(
                id=le.generate_id(sheet, ObjectType.KPI_ROW),
                type=ObjectType.KPI_ROW,
                payload=PlacedKPIRow(kpis=config.kpis[:6]),
            )
            le.insert_object(sheet, kpi_obj)

        # Section header
        sh_obj = PlacedObject(
            id=le.generate_id(sheet, ObjectType.SECTION_HEADER),
            type=ObjectType.SECTION_HEADER,
            payload=PlacedSectionHeader(text="Analytics Overview"),
        )
        le.insert_object(sheet, sh_obj)

        # Charts — pair them left/right
        charts = config.charts[:6]
        for i, chart_cfg in enumerate(charts):
            side = "left" if i % 2 == 0 else "right"
            width = "half"
            # Make first chart full if it's a line/area
            if i == 0 and chart_cfg.type in (ChartType.LINE, ChartType.AREA):
                width = "full"
                side = "left"

            chart_obj = PlacedObject(
                id=le.generate_id(sheet, ObjectType.CHART),
                type=ObjectType.CHART,
                payload=PlacedChart(chart=chart_cfg, width=width, side=side),
            )
            le.insert_object(sheet, chart_obj)

        # Section header for table
        sh2_obj = PlacedObject(
            id=le.generate_id(sheet, ObjectType.SECTION_HEADER),
            type=ObjectType.SECTION_HEADER,
            payload=PlacedSectionHeader(text="Data Summary"),
        )
        le.insert_object(sheet, sh2_obj)

        # Table
        table_obj = PlacedObject(
            id=le.generate_id(sheet, ObjectType.TABLE),
            type=ObjectType.TABLE,
            payload=PlacedTable(
                columns=config.table_columns[:8],
                max_rows=15,
                show_conditional=True,
            ),
        )
        le.insert_object(sheet, table_obj)

        return state

    # ── Render ───────────────────────────────────────────────────────────────

    def _render_and_save(self, console=None) -> None:
        """Render the current state to xlsx."""
        if self.df is None:
            return

        self.state.version += 1
        self.output_path.parent.mkdir(parents=True, exist_ok=True)

        try:
            if console:
                from rich.status import Status
                with console.status("[bold cyan]Rendering dashboard...", spinner="dots"):
                    template = FlexibleTemplate(self.state)
                    template.build_from_state(self.df, self.output_path)
            else:
                print("Rendering dashboard...")
                template = FlexibleTemplate(self.state)
                template.build_from_state(self.df, self.output_path)
        except Exception as e:
            print(f"Render error: {e}")
            traceback.print_exc()
            return

        print(f"Saved: {self.output_path}")

    # ── Undo / Redo ──────────────────────────────────────────────────────────

    def _push_undo(self) -> None:
        self.undo_stack.append(self.state.snapshot())
        if len(self.undo_stack) > MAX_UNDO:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def _undo(self) -> None:
        if not self.undo_stack:
            print("Nothing to undo.")
            return
        self.redo_stack.append(self.state.snapshot())
        self.state = self.undo_stack.pop()
        print("Undone. Re-rendering...")
        self._render_and_save()

    def _redo(self) -> None:
        if not self.redo_stack:
            print("Nothing to redo.")
            return
        self.undo_stack.append(self.state.snapshot())
        self.state = self.redo_stack.pop()
        print("Redone. Re-rendering...")
        self._render_and_save()

    # ── Display ──────────────────────────────────────────────────────────────

    def _show_state(self, console=None) -> None:
        if console:
            try:
                from rich.table import Table
                for sheet in self.state.sheets:
                    tbl = Table(title=f"Sheet: {sheet.name}")
                    tbl.add_column("ID", style="cyan")
                    tbl.add_column("Type", style="green")
                    tbl.add_column("Row", style="yellow")
                    tbl.add_column("Description")
                    for obj in sheet.sorted_objects():
                        from .prompts import _describe_object
                        tbl.add_row(
                            obj.id,
                            obj.type.value,
                            str(obj.anchor_row),
                            _describe_object(obj),
                        )
                    console.print(tbl)
                return
            except ImportError:
                pass

        # Fallback plain text
        for sheet in self.state.sheets:
            print(f"\n--- Sheet: {sheet.name} ---")
            from .prompts import _describe_object
            for obj in sheet.sorted_objects():
                print(f"  [{obj.id}] row {obj.anchor_row}: {_describe_object(obj)}")

    def _print_summary(self, console=None) -> None:
        n_charts = sum(
            1 for s in self.state.sheets
            for o in s.objects if o.type == ObjectType.CHART
        )
        n_kpis = sum(
            len(o.payload.kpis) for s in self.state.sheets
            for o in s.objects if o.type == ObjectType.KPI_ROW
        )
        n_tables = sum(
            1 for s in self.state.sheets
            for o in s.objects if o.type in (ObjectType.TABLE, ObjectType.PIVOT)
        )
        total = sum(len(s.objects) for s in self.state.sheets)
        print(f"Dashboard v{self.state.version}: {n_charts} charts, {n_kpis} KPIs, {n_tables} tables ({total} total objects)")
