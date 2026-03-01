"""Layout engine for positioning objects on a sheet."""
from __future__ import annotations

from .models import (
    ObjectType,
    PlacedChart,
    PlacedObject,
    SheetLayout,
)

# ─── Row Height Constants ────────────────────────────────────────────────────

HEIGHT = {
    ObjectType.TITLE: 2,
    ObjectType.FILTER_PANEL: 2,
    ObjectType.KPI_ROW: 5,
    ObjectType.SECTION_HEADER: 1,
    ObjectType.CHART: 15,         # default for half; full uses 14
    ObjectType.TABLE: 17,         # max_rows+2 default
    ObjectType.PIVOT: 17,
    ObjectType.TEXT: 3,
}

CHART_HEIGHT_HALF = 15
CHART_HEIGHT_FULL = 14


def height_for(obj: PlacedObject) -> int:
    """Compute row height for an object."""
    if obj.type == ObjectType.CHART:
        p: PlacedChart = obj.payload  # type: ignore[assignment]
        return CHART_HEIGHT_FULL if p.width == "full" else CHART_HEIGHT_HALF
    if obj.type in (ObjectType.TABLE, ObjectType.PIVOT):
        max_rows = getattr(obj.payload, "max_rows", 15)
        return max_rows + 2
    return HEIGHT.get(obj.type, 3)


# ─── Layout Engine ───────────────────────────────────────────────────────────

class LayoutEngine:

    @staticmethod
    def generate_id(sheet: SheetLayout, obj_type: ObjectType) -> str:
        prefix = obj_type.value
        existing = [o.id for o in sheet.objects if o.id.startswith(prefix + "_")]
        if not existing:
            return f"{prefix}_0"
        nums = []
        for eid in existing:
            suffix = eid[len(prefix) + 1:]
            if suffix.isdigit():
                nums.append(int(suffix))
        nxt = max(nums, default=-1) + 1
        return f"{prefix}_{nxt}"

    @staticmethod
    def find_half_pair_row(sheet: SheetLayout, side: str) -> int | None:
        """Find a half-chart that can be paired (opposite side at same row)."""
        opposite = "right" if side == "left" else "left"
        for obj in sheet.objects:
            if obj.type != ObjectType.CHART:
                continue
            p: PlacedChart = obj.payload  # type: ignore[assignment]
            if p.width != "half" or p.side != opposite:
                continue
            # Check no existing chart already occupies that side at this row
            already = any(
                o.type == ObjectType.CHART
                and o.anchor_row == obj.anchor_row
                and getattr(o.payload, "side", "") == side
                for o in sheet.objects
                if o.id != obj.id
            )
            if not already:
                return obj.anchor_row
        return None

    @classmethod
    def insert_object(
        cls,
        sheet: SheetLayout,
        obj: PlacedObject,
        position: str = "end",
    ) -> None:
        """Insert object into sheet at the given position.

        position: "end", "after:<id>", "row:<N>"
        """
        obj.height_rows = height_for(obj)

        if position.startswith("row:"):
            target_row = int(position.split(":")[1])
            obj.anchor_row = target_row
        elif position.startswith("after:"):
            ref_id = position.split(":", 1)[1]
            ref = sheet.find_object(ref_id)
            if ref:
                obj.anchor_row = ref.end_row + 1
            else:
                obj.anchor_row = sheet.next_free_row()
        else:
            # "end" — default
            # For half charts, try to pair with an existing opposite half
            if obj.type == ObjectType.CHART:
                p: PlacedChart = obj.payload  # type: ignore[assignment]
                if p.width == "half":
                    pair_row = cls.find_half_pair_row(sheet, p.side)
                    if pair_row is not None:
                        obj.anchor_row = pair_row
                        sheet.objects.append(obj)
                        return
            obj.anchor_row = sheet.next_free_row()

        sheet.objects.append(obj)

    @staticmethod
    def reflow(sheet: SheetLayout) -> None:
        """Recompute anchor_rows from scratch, closing gaps and preserving half-chart pairs."""
        if not sheet.objects:
            return
        sorted_objs = sorted(
            sheet.objects,
            key=lambda o: (o.anchor_row, 0 if getattr(getattr(o, "payload", None), "side", "left") == "left" else 1),
        )

        cursor = 0
        i = 0
        while i < len(sorted_objs):
            obj = sorted_objs[i]
            h = height_for(obj)

            # Check if next object is a paired half-chart at the same original row
            if (
                obj.type == ObjectType.CHART
                and getattr(obj.payload, "width", "") == "half"
                and i + 1 < len(sorted_objs)
                and sorted_objs[i + 1].type == ObjectType.CHART
                and getattr(sorted_objs[i + 1].payload, "width", "") == "half"
                and sorted_objs[i + 1].anchor_row == obj.anchor_row
            ):
                pair = sorted_objs[i + 1]
                obj.anchor_row = cursor
                obj.height_rows = h
                pair.anchor_row = cursor
                pair.height_rows = height_for(pair)
                cursor += max(h, pair.height_rows)
                i += 2
            else:
                obj.anchor_row = cursor
                obj.height_rows = h
                cursor += h
                i += 1

        # Sync back
        for obj in sorted_objs:
            for real in sheet.objects:
                if real.id == obj.id:
                    real.anchor_row = obj.anchor_row
                    real.height_rows = obj.height_rows
                    break
