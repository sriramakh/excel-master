"""ObjectRegistry — tracks every artifact created in an agent session."""
from __future__ import annotations

from datetime import datetime, timezone
from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


class OperationType(str, Enum):
    # High-level dashboard objects
    CHART = "chart"
    TABLE = "table"
    KPI_ROW = "kpi_row"
    PIVOT = "pivot"
    TITLE = "title"
    FILTER_PANEL = "filter_panel"
    SECTION_HEADER = "section_header"
    TEXT = "text"
    CONTENT = "content"
    # Cell-level operations
    CELL_WRITE = "cell_write"
    CELL_FORMAT = "cell_format"
    CONDITIONAL_FORMAT = "conditional_format"
    DATA_VALIDATION = "data_validation"
    MERGE = "merge"
    HYPERLINK = "hyperlink"
    COMMENT = "comment"
    IMAGE = "image"
    # Structural
    SHEET = "sheet"
    ROW_COL = "row_col"
    THEME = "theme"


class RegistryEntry(BaseModel):
    id: str                       # "chart_0", "cell_write_3"
    op_type: str                  # OperationType value
    sheet: str = "Dashboard"
    location: str = ""            # "B3", "A1:F20", "row 5"
    description: str = ""         # "Bar chart: Revenue by Region"
    created_at: str = ""          # ISO timestamp
    turn: int = 0                 # conversation turn number
    params: dict = Field(default_factory=dict)


class ObjectRegistry:
    """Dual-layer tracker: semantic objects + cell-level operations."""

    def __init__(self) -> None:
        self._entries: dict[str, RegistryEntry] = {}
        self._counters: dict[str, int] = {}

    def register(
        self,
        op_type: str | OperationType,
        sheet: str = "Dashboard",
        location: str = "",
        description: str = "",
        turn: int = 0,
        params: dict | None = None,
        entry_id: str | None = None,
    ) -> str:
        """Register a new artifact. Returns the assigned ID."""
        # Normalize enum to its string value
        op_type_str = op_type.value if isinstance(op_type, OperationType) else str(op_type)

        if entry_id is None:
            prefix = op_type_str
            n = self._counters.get(prefix, 0)
            entry_id = f"{prefix}_{n}"
            self._counters[prefix] = n + 1

        entry = RegistryEntry(
            id=entry_id,
            op_type=op_type_str,
            sheet=sheet,
            location=location,
            description=description,
            created_at=datetime.now(timezone.utc).isoformat(),
            turn=turn,
            params=params or {},
        )
        self._entries[entry_id] = entry
        return entry_id

    def get(self, entry_id: str) -> RegistryEntry | None:
        return self._entries.get(entry_id)

    def list_all(
        self,
        sheet: str | None = None,
        op_type: str | None = None,
    ) -> list[RegistryEntry]:
        entries = list(self._entries.values())
        if sheet:
            entries = [e for e in entries if e.sheet == sheet]
        if op_type:
            entries = [e for e in entries if e.op_type == op_type]
        return entries

    def remove(self, entry_id: str) -> RegistryEntry | None:
        return self._entries.pop(entry_id, None)

    def to_snapshot(self) -> str:
        """Compact text representation for LLM context injection."""
        if not self._entries:
            return "Registry: (empty)"
        lines = ["Registry:"]
        for e in self._entries.values():
            loc = f" @ {e.location}" if e.location else ""
            lines.append(f"  [{e.id}] {e.op_type} on {e.sheet}{loc}: {e.description}")
        return "\n".join(lines)

    def snapshot_dict(self) -> list[dict]:
        """Serializable list for undo snapshots."""
        return [e.model_dump() for e in self._entries.values()]

    def restore(self, data: list[dict]) -> None:
        """Restore registry from a snapshot_dict result."""
        self._entries.clear()
        self._counters.clear()
        for d in data:
            entry = RegistryEntry(**d)
            self._entries[entry.id] = entry
            # Rebuild counters
            parts = entry.id.rsplit("_", 1)
            if len(parts) == 2 and parts[1].isdigit():
                prefix = parts[0]
                n = int(parts[1]) + 1
                if n > self._counters.get(prefix, 0):
                    self._counters[prefix] = n

    def clear(self) -> None:
        self._entries.clear()
        self._counters.clear()
