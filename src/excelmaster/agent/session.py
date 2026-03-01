"""AgentSession — programmatic API for the agentic Excel Master."""
from __future__ import annotations

import traceback
from pathlib import Path
from typing import Any

import pandas as pd

from ..config import get_settings
from ..data.data_engine import profile_dataset, discover_and_join, profile_to_prompt_text
from ..dashboard.template_selector import TemplateSelector
from ..models import DatasetProfile
from ..chat.layout import LayoutEngine
from ..chat.models import (
    ObjectType,
    PlacedChart,
    PlacedFilterPanel,
    PlacedKPIRow,
    PlacedObject,
    PlacedSectionHeader,
    PlacedTable,
    PlacedText,
    PlacedTitle,
    SheetLayout,
    WorkbookState,
)
from ..chat.renderer import FlexibleTemplate
from .llm_bridge import ToolCallingBridge, SYSTEM_PREAMBLE
from .registry import ObjectRegistry
from .tool_executor import ToolExecutor
from .tools import get_tool_schemas

MAX_UNDO = 30
MAX_TOOL_ROUNDS = 5
MAX_HISTORY = 20


class AgentSession:
    """Programmatic API — usable from CLI, Python, or outer orchestration agents."""

    def __init__(
        self,
        data_path: str | Path,
        output_path: str | Path | None = None,
    ) -> None:
        self.data_path = Path(data_path)
        self.output_path: Path = (
            Path(output_path) if output_path
            else self._default_output(self.data_path)
        )
        self.df: pd.DataFrame | None = None
        self.profile: DatasetProfile | None = None
        self.state = WorkbookState()
        self.registry = ObjectRegistry()
        self.bridge = ToolCallingBridge()

        # History
        self.messages: list[dict] = []
        self._turn: int = 0
        self._undo_stack: list[tuple[WorkbookState, list[dict]]] = []
        self._redo_stack: list[tuple[WorkbookState, list[dict]]] = []

    @staticmethod
    def _default_output(data_path: Path) -> Path:
        stem = data_path.stem
        return data_path.parent.parent / "output" / f"{stem}_agent_dashboard.xlsx"

    # ── Load ──────────────────────────────────────────────────────────────────

    def load(self) -> dict:
        """Load data, profile it, and initialize the conversation."""
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
        schema_text = profile_to_prompt_text(self.profile)

        system_msg = f"""{SYSTEM_PREAMBLE}

## DATA SCHEMA
{schema_text}
"""
        self.messages = [{"role": "system", "content": system_msg}]

        return {
            "rows": len(self.df),
            "columns": len(self.df.columns),
            "column_names": list(self.df.columns),
        }

    # ── Execute Natural Language Instruction ──────────────────────────────────

    def execute_instruction(self, instruction: str) -> dict:
        """NL instruction → LLM tool calling → execute → result.

        Returns dict with keys: text, actions, object_ids.
        """
        self._push_undo()
        self._turn += 1

        # Build user message with registry snapshot
        snapshot = self.registry.to_snapshot()
        user_content = f"""Current state:
{self._state_snapshot()}

{snapshot}

User instruction: {instruction}"""

        self.messages.append({"role": "user", "content": user_content})
        self._trim_history()

        all_actions: list[dict] = []
        all_ids: list[str] = []
        final_text = ""

        try:
            text, tool_calls = self.bridge.call_with_tools(self.messages)
            final_text = text

            rounds = 0
            while tool_calls and rounds < MAX_TOOL_ROUNDS:
                rounds += 1

                # Build assistant message with tool calls for history
                asst_msg = self.bridge.build_assistant_tool_call_message(text, tool_calls)
                self.messages.append(asst_msg)

                # Execute each tool call
                tool_results = []
                for tc in tool_calls:
                    executor = ToolExecutor(
                        self.state, self.registry, self.df, self._turn,
                    )
                    result = executor.execute(tc["name"], tc["arguments"])
                    all_actions.append({
                        "tool": tc["name"],
                        "args": tc["arguments"],
                        "result": result,
                    })
                    if result.get("object_id"):
                        all_ids.append(result["object_id"])

                    tool_results.append({
                        "tool_call_id": tc["id"],
                        "content": str(result),
                    })

                # Send results back to LLM for follow-up
                text, tool_calls = self.bridge.send_tool_results(
                    self.messages, tool_results,
                )
                # Add tool result messages to history
                for tr in tool_results:
                    self.messages.append({
                        "role": "tool",
                        "tool_call_id": tr["tool_call_id"],
                        "content": tr["content"],
                    })

                if text:
                    final_text = text

            # Store final assistant text
            if final_text:
                self.messages.append({"role": "assistant", "content": final_text})

        except Exception as e:
            final_text = f"Error: {e}"
            # Roll back on failure
            if self._undo_stack:
                self.state, reg_snap = self._undo_stack.pop()
                self.registry.restore(reg_snap)

        return {
            "text": final_text,
            "actions": all_actions,
            "object_ids": all_ids,
        }

    # ── Execute Tool Directly (no LLM) ───────────────────────────────────────

    def execute_tool(self, tool_name: str, args: dict) -> dict:
        """Direct tool call — bypasses LLM entirely. Fast & free."""
        self._push_undo()
        self._turn += 1
        executor = ToolExecutor(self.state, self.registry, self.df, self._turn)
        return executor.execute(tool_name, args)

    # ── Save ──────────────────────────────────────────────────────────────────

    def save(self, output_path: str | Path | None = None) -> Path:
        """Render and save the current state to xlsx."""
        if self.df is None:
            raise RuntimeError("No data loaded. Call load() first.")

        out = Path(output_path) if output_path else self.output_path
        out.parent.mkdir(parents=True, exist_ok=True)

        self.state.version += 1
        template = FlexibleTemplate(self.state)
        template.build_from_state(self.df, out, profile=self.profile)
        self.output_path = out
        return out

    # ── Undo / Redo ───────────────────────────────────────────────────────────

    def _push_undo(self) -> None:
        snapshot = (self.state.snapshot(), self.registry.snapshot_dict())
        self._undo_stack.append(snapshot)
        if len(self._undo_stack) > MAX_UNDO:
            self._undo_stack.pop(0)
        self._redo_stack.clear()

    def undo(self) -> bool:
        if not self._undo_stack:
            return False
        self._redo_stack.append(
            (self.state.snapshot(), self.registry.snapshot_dict())
        )
        self.state, reg_snap = self._undo_stack.pop()
        self.registry.restore(reg_snap)
        return True

    def redo(self) -> bool:
        if not self._redo_stack:
            return False
        self._undo_stack.append(
            (self.state.snapshot(), self.registry.snapshot_dict())
        )
        self.state, reg_snap = self._redo_stack.pop()
        self.registry.restore(reg_snap)
        return True

    # ── State Inspection ──────────────────────────────────────────────────────

    def get_state(self) -> dict:
        """Full state snapshot for external agents."""
        return {
            "title": self.state.title,
            "theme": self.state.theme_key,
            "version": self.state.version,
            "sheets": [
                {
                    "name": s.name,
                    "objects": len(s.objects),
                    "hidden": s.hidden,
                }
                for s in self.state.sheets
            ],
            "registry": self.registry.snapshot_dict(),
            "turn": self._turn,
        }

    def get_tools(self) -> list[dict]:
        """Return tool schemas for discovery by outer agents."""
        return get_tool_schemas()

    # ── Auto Dashboard ────────────────────────────────────────────────────────

    def auto_dashboard(self) -> dict:
        """Build a full dashboard using LLM template selection."""
        if self.profile is None:
            raise RuntimeError("No data loaded. Call load() first.")

        self._push_undo()
        selector = TemplateSelector()
        config = selector.select(self.profile)

        # Reuse the ChatEngine._config_to_state pattern
        from ..chat.engine import ChatEngine
        dummy = ChatEngine.__new__(ChatEngine)
        self.state = dummy._config_to_state(config)

        # Register all objects in the new state
        self.registry.clear()
        self._turn += 1
        for sheet in self.state.sheets:
            for obj in sheet.objects:
                from ..chat.prompts import _describe_object
                self.registry.register(
                    op_type=obj.type.value,
                    sheet=sheet.name,
                    location=f"row {obj.anchor_row}",
                    description=_describe_object(obj),
                    turn=self._turn,
                    entry_id=obj.id,
                )

        return {
            "text": f"Auto-built dashboard: {self.state.title}",
            "actions": [{"tool": "auto_dashboard", "args": {}, "result": {"success": True}}],
            "object_ids": [o.id for s in self.state.sheets for o in s.objects],
        }

    # ── Interactive REPL ──────────────────────────────────────────────────────

    def run_repl(self) -> None:
        """Interactive REPL for CLI usage."""
        try:
            from rich.console import Console
            from rich.panel import Panel
            from rich.table import Table
            console = Console()
        except ImportError:
            console = None

        info = self.load()
        welcome = (
            f"File: {self.data_path.name}\n"
            f"Shape: {info['rows']:,} rows x {info['columns']} columns\n"
            f"Columns: {', '.join(info['column_names'][:12])}"
            + (f" ... +{len(info['column_names'])-12} more"
               if len(info['column_names']) > 12 else "")
        )

        if console:
            console.print(Panel(welcome, title="Excel Master Agent", border_style="blue"))
            console.print(
                "[bold]Commands:[/bold] [cyan]auto[/cyan] (full dashboard) | "
                "type an instruction\n"
                "Special: [cyan]undo[/cyan] | [cyan]redo[/cyan] | "
                "[cyan]show[/cyan] | [cyan]save[/cyan] | [cyan]save as <name>[/cyan] | "
                "[cyan]quit[/cyan]\n"
            )
        else:
            print(f"\n=== Excel Master Agent ===\n{welcome}")
            print("Commands: auto | undo | redo | show | save | save as <name> | quit\n")

        while True:
            try:
                user_input = input("You: ").strip()
            except (EOFError, KeyboardInterrupt):
                print("\nGoodbye!")
                break

            if not user_input:
                continue

            low = user_input.lower()

            if low in ("quit", "exit", "q"):
                print("Goodbye!")
                break
            elif low == "undo":
                if self.undo():
                    self._save_and_print(console)
                else:
                    print("Nothing to undo.")
                continue
            elif low == "redo":
                if self.redo():
                    self._save_and_print(console)
                else:
                    print("Nothing to redo.")
                continue
            elif low == "show":
                self._show_state(console)
                continue
            elif low == "save":
                self._save_and_print(console)
                continue
            elif low.startswith("save as "):
                name = user_input[8:].strip()
                if name:
                    new_path = self.output_path.parent / f"{name}.xlsx"
                    self.save(new_path)
                    print(f"Saved as: {new_path}")
                continue
            elif low == "auto":
                try:
                    result = self.auto_dashboard()
                    print(f"Assistant: {result['text']}")
                    self._save_and_print(console)
                except Exception as e:
                    print(f"Error: {e}")
                    traceback.print_exc()
                continue

            # Normal instruction → LLM tool calling
            try:
                if console:
                    from rich.status import Status
                    with console.status("[bold cyan]Thinking...", spinner="dots"):
                        result = self.execute_instruction(user_input)
                else:
                    result = self.execute_instruction(user_input)

                if result["text"]:
                    print(f"Assistant: {result['text']}")

                for a in result["actions"]:
                    r = a["result"]
                    status = "ok" if r.get("success") else "FAIL"
                    print(f"  [{status}] {a['tool']}: {r.get('message', '')}")

                if result["actions"]:
                    self._save_and_print(console)

            except Exception as e:
                print(f"Error: {e}")
                traceback.print_exc()

    # ── Internal Helpers ──────────────────────────────────────────────────────

    def _save_and_print(self, console=None) -> None:
        try:
            if console:
                from rich.status import Status
                with console.status("[bold cyan]Rendering...", spinner="dots"):
                    path = self.save()
            else:
                path = self.save()
            print(f"Saved: {path}")
            self._print_summary()
        except Exception as e:
            print(f"Render error: {e}")
            traceback.print_exc()

    def _print_summary(self) -> None:
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
        n_cell_ops = sum(
            len(s.cell_writes) + len(s.merges) + len(s.hyperlinks) +
            len(s.conditional_formats) + len(s.comments)
            for s in self.state.sheets
        )
        print(
            f"Dashboard v{self.state.version}: "
            f"{n_charts} charts, {n_kpis} KPIs, {n_tables} tables, "
            f"{n_cell_ops} cell ops ({total} total objects)"
        )

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
                        from ..chat.prompts import _describe_object
                        tbl.add_row(
                            obj.id, obj.type.value,
                            str(obj.anchor_row), _describe_object(obj),
                        )
                    console.print(tbl)

                # Show registry summary
                entries = self.registry.list_all()
                if entries:
                    console.print(f"\n[dim]Registry: {len(entries)} entries[/dim]")
                return
            except ImportError:
                pass

        for sheet in self.state.sheets:
            print(f"\n--- Sheet: {sheet.name} ---")
            from ..chat.prompts import _describe_object
            for obj in sheet.sorted_objects():
                print(f"  [{obj.id}] row {obj.anchor_row}: {_describe_object(obj)}")

    def _state_snapshot(self) -> str:
        from ..chat.prompts import state_to_snapshot
        return state_to_snapshot(self.state)

    def _trim_history(self) -> None:
        """Keep system prompt + last MAX_HISTORY user/assistant turns."""
        if len(self.messages) <= 1 + MAX_HISTORY * 2:
            return
        system = self.messages[0]
        recent = self.messages[-(MAX_HISTORY * 2):]
        self.messages = [system] + recent
