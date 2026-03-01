"""xlsxwriter format factory for Excel Master dashboards."""
from __future__ import annotations
from .themes import Theme


def _hex(color: str) -> str:
    """Ensure color has # prefix."""
    c = color.lstrip("#")
    return f"#{c}"


class StyleFactory:
    """Creates and caches xlsxwriter format objects for a given workbook+theme."""

    def __init__(self, wb, theme: Theme):
        self.wb = wb
        self.t = theme
        self._cache: dict[str, object] = {}

    def _f(self, key: str, props: dict):
        if key not in self._cache:
            self._cache[key] = self.wb.add_format(props)
        return self._cache[key]

    # ── Title / Header ─────────────────────────────────────────────────────────

    def title(self):
        t = self.t
        return self._f("title", {
            "font_name": t.font_heading, "font_size": 22, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.primary),
            "align": "left", "valign": "vcenter", "indent": 2,
        })

    def subtitle(self):
        t = self.t
        return self._f("subtitle", {
            "font_name": t.font_body, "font_size": 10,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.secondary),
            "align": "left", "valign": "vcenter", "indent": 3,
        })

    def bg(self, color: str | None = None):
        t = self.t
        c = _hex(color or t.bg_dashboard)
        return self._f(f"bg_{c}", {"bg_color": c})

    # ── Filter bar ─────────────────────────────────────────────────────────────

    def filter_label(self):
        t = self.t
        return self._f("filter_label", {
            "font_name": t.font_body, "font_size": 9, "bold": True,
            "font_color": _hex(t.text_muted), "bg_color": _hex(t.bg_dashboard),
            "align": "right", "valign": "vcenter",
        })

    def filter_value(self):
        t = self.t
        return self._f("filter_value", {
            "font_name": t.font_body, "font_size": 10, "bold": True,
            "font_color": _hex(t.primary), "bg_color": _hex(t.bg_card),
            "align": "center", "valign": "vcenter",
            "border": 1, "border_color": _hex(t.primary),
        })

    # ── Section headers ────────────────────────────────────────────────────────

    def section_header(self):
        t = self.t
        return self._f("section_header", {
            "font_name": t.font_heading, "font_size": 11, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.secondary),
            "align": "left", "valign": "vcenter", "indent": 1,
        })

    def section_header_alt(self):
        """Alternate section header using accent1."""
        t = self.t
        return self._f("section_header_alt", {
            "font_name": t.font_heading, "font_size": 11, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.accent1),
            "align": "left", "valign": "vcenter", "indent": 1,
        })

    # ── KPI Card cells ─────────────────────────────────────────────────────────

    def kpi_bg(self, bg_color: str | None = None):
        t = self.t
        bg = _hex(bg_color or t.bg_card)
        return self._f(f"kpi_bg_{bg}", {"bg_color": bg})

    def kpi_label(self, bg_color: str | None = None):
        t = self.t
        bg = _hex(bg_color or t.bg_card)
        return self._f(f"kpi_label_{bg}", {
            "font_name": t.font_body, "font_size": 8,
            "font_color": _hex(t.text_muted), "bg_color": bg,
            "align": "center", "valign": "vcenter",
        })

    def kpi_value(self, bg_color: str | None = None, font_color: str | None = None):
        t = self.t
        bg = _hex(bg_color or t.bg_card)
        fc = _hex(font_color or t.primary)
        return self._f(f"kpi_val_{bg}_{fc}", {
            "font_name": t.font_heading, "font_size": 20, "bold": True,
            "font_color": fc, "bg_color": bg,
            "align": "center", "valign": "vcenter",
        })

    def kpi_delta(self, positive: bool = True, bg_color: str | None = None):
        t = self.t
        bg = _hex(bg_color or t.bg_card)
        fc = _hex(t.positive if positive else t.negative)
        key = f"kpi_delta_{'p' if positive else 'n'}_{bg}"
        return self._f(key, {
            "font_name": t.font_body, "font_size": 8, "bold": True,
            "font_color": fc, "bg_color": bg,
            "align": "center", "valign": "vcenter",
        })

    def kpi_empty(self, bg_color: str | None = None):
        t = self.t
        bg = _hex(bg_color or t.bg_card)
        return self._f(f"kpi_empty_{bg}", {"bg_color": bg})

    # ── Table ──────────────────────────────────────────────────────────────────

    def table_header(self):
        t = self.t
        return self._f("table_header", {
            "font_name": t.font_body, "font_size": 9, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.bg_table_header),
            "align": "center", "valign": "vcenter",
            "border": 1, "border_color": "#CCCCCC",
        })

    def table_data(self, stripe: bool = False, align: str = "left"):
        t = self.t
        bg = "#F7F8FA" if stripe else "#FFFFFF"
        if t.dark_mode:
            bg = _hex(t.secondary) if stripe else _hex(t.bg_card)
        key = f"table_data_{'s' if stripe else 'w'}_{align}"
        return self._f(key, {
            "font_name": t.font_body, "font_size": 9,
            "font_color": _hex(t.text_primary), "bg_color": bg,
            "align": align, "valign": "vcenter",
            "bottom": 1, "bottom_color": "#E5E7EB",
        })

    def table_data_num(self, stripe: bool = False, fmt: str = "#,##0.##"):
        t = self.t
        bg = "#F7F8FA" if stripe else "#FFFFFF"
        if t.dark_mode:
            bg = _hex(t.secondary) if stripe else _hex(t.bg_card)
        key = f"table_data_num_{'s' if stripe else 'w'}_{fmt[:6]}"
        return self._f(key, {
            "font_name": t.font_body, "font_size": 9,
            "font_color": _hex(t.text_primary), "bg_color": bg,
            "align": "right", "valign": "vcenter",
            "num_format": fmt,
            "bottom": 1, "bottom_color": "#E5E7EB",
        })

    def table_currency(self, stripe: bool = False):
        return self.table_data_num(stripe, '"$"#,##0')

    def table_pct(self, stripe: bool = False):
        return self.table_data_num(stripe, '0.0"%"')

    # ── Insight box ────────────────────────────────────────────────────────────

    def insight_header(self):
        t = self.t
        return self._f("insight_header", {
            "font_name": t.font_heading, "font_size": 10, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.accent1),
            "align": "left", "valign": "vcenter", "indent": 1,
        })

    def insight_text(self):
        t = self.t
        return self._f("insight_text", {
            "font_name": t.font_body, "font_size": 9,
            "font_color": _hex(t.text_primary), "bg_color": _hex(t.bg_card),
            "align": "left", "valign": "vcenter",
            "indent": 2, "text_wrap": True,
        })

    # ── Badge / Status ─────────────────────────────────────────────────────────

    def badge_positive(self):
        t = self.t
        return self._f("badge_pos", {
            "font_name": t.font_body, "font_size": 8, "bold": True,
            "font_color": "#FFFFFF", "bg_color": _hex(t.positive),
            "align": "center", "valign": "vcenter",
        })

    def badge_negative(self):
        t = self.t
        return self._f("badge_neg", {
            "font_name": t.font_body, "font_size": 8, "bold": True,
            "font_color": "#FFFFFF", "bg_color": _hex(t.negative),
            "align": "center", "valign": "vcenter",
        })

    def badge_neutral(self):
        t = self.t
        return self._f("badge_neu", {
            "font_name": t.font_body, "font_size": 8, "bold": True,
            "font_color": "#FFFFFF", "bg_color": _hex(t.neutral),
            "align": "center", "valign": "vcenter",
        })

    # ── Data sheet ─────────────────────────────────────────────────────────────

    def data_header(self):
        return self._f("data_header", {
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "bg_color": "#2C3E50", "font_color": "#FFFFFF",
            "align": "center", "valign": "vcenter",
            "border": 1, "border_color": "#555555",
        })

    def data_cell(self):
        return self._f("data_cell", {
            "font_name": "Calibri", "font_size": 9,
            "valign": "vcenter",
            "bottom": 1, "bottom_color": "#EEEEEE",
        })

    # ── Deep Analysis sheet ─────────────────────────────────────────────────

    def analysis_title(self):
        t = self.t
        return self._f("analysis_title", {
            "font_name": t.font_heading, "font_size": 20, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.primary),
            "align": "left", "valign": "vcenter", "indent": 2,
        })

    def analysis_subtitle(self):
        t = self.t
        return self._f("analysis_subtitle", {
            "font_name": t.font_body, "font_size": 9, "italic": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.primary),
            "align": "left", "valign": "vcenter", "indent": 2,
        })

    def analysis_meta_right(self):
        t = self.t
        return self._f("analysis_meta_right", {
            "font_name": t.font_body, "font_size": 9,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.primary),
            "align": "right", "valign": "vcenter",
        })

    def analysis_section_header(self):
        t = self.t
        return self._f("analysis_section_header", {
            "font_name": t.font_heading, "font_size": 12, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.secondary),
            "align": "left", "valign": "vcenter", "indent": 1,
        })

    def analysis_subheader(self):
        t = self.t
        return self._f("analysis_subheader", {
            "font_name": t.font_heading, "font_size": 10, "bold": True,
            "font_color": _hex(t.primary), "bg_color": "#F5F5F5",
            "align": "left", "valign": "vcenter", "indent": 1,
            "bottom": 1, "bottom_color": _hex(t.secondary),
        })

    def analysis_body(self):
        t = self.t
        return self._f("analysis_body", {
            "font_name": t.font_body, "font_size": 10,
            "font_color": _hex(t.text_primary), "bg_color": "#FFFFFF",
            "align": "left", "valign": "top",
            "text_wrap": True, "indent": 1,
        })

    def analysis_bullet(self):
        t = self.t
        return self._f("analysis_bullet", {
            "font_name": t.font_body, "font_size": 9, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.accent1),
            "align": "center", "valign": "vcenter",
        })

    def score_badge(self, score: int):
        if score >= 80:
            bg = self.t.positive
        elif score >= 60:
            bg = "F4A261"     # amber
        else:
            bg = self.t.negative
        key = f"score_badge_{score}"
        return self._f(key, {
            "font_name": self.t.font_heading, "font_size": 16, "bold": True,
            "font_color": "#FFFFFF", "bg_color": _hex(bg),
            "align": "center", "valign": "vcenter",
        })

    def direction_badge(self, direction: str):
        d = direction.lower()
        if d == "up":
            bg = self.t.positive
            symbol = "▲"
        elif d == "down":
            bg = self.t.negative
            symbol = "▼"
        else:
            bg = self.t.neutral
            symbol = "→"
        key = f"dir_badge_{d}"
        fmt = self._f(key, {
            "font_name": self.t.font_body, "font_size": 10, "bold": True,
            "font_color": "#FFFFFF", "bg_color": _hex(bg),
            "align": "center", "valign": "vcenter",
        })
        return fmt, symbol

    def analysis_footer(self):
        t = self.t
        return self._f("analysis_footer", {
            "font_name": t.font_body, "font_size": 8, "italic": True,
            "font_color": _hex(t.text_muted), "bg_color": _hex(t.primary),
            "align": "center", "valign": "vcenter",
        })

    def analysis_table_header(self):
        t = self.t
        return self._f("analysis_tbl_hdr", {
            "font_name": t.font_body, "font_size": 9, "bold": True,
            "font_color": _hex(t.text_light), "bg_color": _hex(t.secondary),
            "align": "center", "valign": "vcenter",
            "border": 1, "border_color": "#CCCCCC",
        })

    def analysis_table_cell(self, stripe: bool = False):
        bg = "#F7F8FA" if stripe else "#FFFFFF"
        key = f"analysis_tbl_cell_{'s' if stripe else 'w'}"
        t = self.t
        return self._f(key, {
            "font_name": t.font_body, "font_size": 9,
            "font_color": _hex(t.text_primary), "bg_color": bg,
            "align": "left", "valign": "vcenter", "indent": 1,
            "bottom": 1, "bottom_color": "#E5E7EB", "text_wrap": True,
        })

    def analysis_table_num(self, stripe: bool = False):
        bg = "#F7F8FA" if stripe else "#FFFFFF"
        key = f"analysis_tbl_num_{'s' if stripe else 'w'}"
        t = self.t
        return self._f(key, {
            "font_name": t.font_body, "font_size": 9,
            "font_color": _hex(t.text_primary), "bg_color": bg,
            "align": "right", "valign": "vcenter",
            "num_format": "#,##0.##",
            "bottom": 1, "bottom_color": "#E5E7EB",
        })
