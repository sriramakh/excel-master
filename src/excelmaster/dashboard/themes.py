"""Color themes for Excel dashboards, inspired by reference screenshots."""
from __future__ import annotations
from dataclasses import dataclass, field
from ..models import ColorTheme


@dataclass
class Theme:
    key: str
    name: str
    # Primary palette
    primary: str          # Main header/accent color (hex without #)
    secondary: str        # Secondary accent
    accent1: str          # Highlight 1
    accent2: str          # Highlight 2
    accent3: str          # Highlight 3
    # Background colors
    bg_dashboard: str     # Dashboard sheet background
    bg_card: str          # KPI card background
    bg_table_header: str  # Table header fill
    bg_section: str       # Section divider
    # Text colors
    text_primary: str     # Main text on light bg
    text_light: str       # Text on dark bg
    text_muted: str       # Secondary/muted text
    # Special
    positive: str         # Up/good indicator
    negative: str         # Down/bad indicator
    neutral: str          # Neutral/info
    # Fonts
    font_heading: str = "Calibri"
    font_body: str = "Calibri"
    # Chart colors (for multi-series)
    chart_colors: list[str] = field(default_factory=list)
    # Style flags
    dark_mode: bool = False
    show_gridlines: bool = False


_UNIVERSAL = Theme(
    key="universal",
    name="Universal",
    primary="1B263B",          # Headers, sidebars, primary data series
    secondary="415A77",        # Secondary data, sub-headers
    accent1="415A77",          # Secondary tone for variety
    accent2="52B788",          # Success green – also used as accent
    accent3="E63946",          # Alert red – also used as accent
    bg_dashboard="E0E1DD",     # Backgrounds, card containers
    bg_card="FFFFFF",
    bg_table_header="1B263B",
    bg_section="415A77",
    text_primary="1B263B",
    text_light="FFFFFF",
    text_muted="778DA9",       # Muted complement to primary
    positive="52B788",         # Success
    negative="E63946",         # Alert
    neutral="415A77",
    chart_colors=[
        "1B263B", "415A77", "52B788", "E63946",
        "778DA9", "0D1B2A", "A3B18A", "D62828",
    ],
)

THEMES: dict[str, Theme] = {
    ColorTheme.CORPORATE_BLUE: _UNIVERSAL,
    ColorTheme.HR_PURPLE: _UNIVERSAL,
    ColorTheme.DARK_MODE: _UNIVERSAL,
    ColorTheme.SUPPLY_GREEN: _UNIVERSAL,
    ColorTheme.FINANCE_GREEN: _UNIVERSAL,
    ColorTheme.MARKETING_ORANGE: _UNIVERSAL,
    ColorTheme.SLATE_MINIMAL: _UNIVERSAL,
    ColorTheme.EXECUTIVE_NAVY: _UNIVERSAL,
}

# Map templates to their default themes
TEMPLATE_DEFAULT_THEME = {
    "executive_summary": ColorTheme.CORPORATE_BLUE,
    "hr_analytics": ColorTheme.HR_PURPLE,
    "dark_operational": ColorTheme.DARK_MODE,
    "financial": ColorTheme.FINANCE_GREEN,
    "supply_chain": ColorTheme.SUPPLY_GREEN,
    "marketing": ColorTheme.MARKETING_ORANGE,
    "minimal_clean": ColorTheme.SLATE_MINIMAL,
}


def get_theme(key: str | ColorTheme) -> Theme:
    if isinstance(key, str) and key in ColorTheme.__members__:
        key = ColorTheme(key)
    elif isinstance(key, str):
        # Try fuzzy match
        for k in THEMES:
            if key.lower() in k.lower() or k.lower() in key.lower():
                return THEMES[k]
    return THEMES.get(key, THEMES[ColorTheme.CORPORATE_BLUE])
