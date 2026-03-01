"""xlsxwriter chart builders for Excel Master dashboards."""
from __future__ import annotations
from dataclasses import dataclass, field

from ..models import ChartConfig, ChartType
from .themes import Theme


def _hex(color: str) -> str:
    return f"#{color.lstrip('#')}"


@dataclass
class ChartZone:
    """Where a chart sits on the worksheet (xlsxwriter-style)."""
    row: int        # 0-indexed anchor row
    col: int        # 0-indexed anchor col
    width: int      # pixels
    height: int     # pixels
    x_offset: int = 5
    y_offset: int = 5


# ── Chart builder functions ────────────────────────────────────────────────────

def _base_style(chart, config: ChartConfig, theme: Theme, zone: ChartZone) -> None:
    """Apply common chart styling."""
    chart.set_title({
        "name": config.title,
        "name_font": {
            "name": theme.font_heading, "size": 11, "bold": True,
            "color": _hex(theme.text_primary),
        },
    })
    chart.set_chartarea({
        "border": {"none": True},
        "fill": {"color": _hex(theme.bg_card)},
    })
    chart.set_plotarea({
        "border": {"none": True},
        "fill": {"color": _hex(theme.bg_card)},
    })
    chart.set_legend({"position": "bottom", "font": {"size": 9}})
    chart.set_size({"width": zone.width, "height": zone.height})


def build_bar_chart(wb, config: ChartConfig, theme: Theme,
                    cat_range: str, val_ranges: list[tuple[str, str]],
                    zone: ChartZone, horizontal: bool = False):
    """Vertical column or horizontal bar chart."""
    ctype = "bar" if horizontal else "column"
    chart = wb.add_chart({"type": ctype})
    colors = [_hex(c) for c in theme.chart_colors]

    for i, (name_r, val_r) in enumerate(val_ranges):
        chart.add_series({
            "name": name_r,
            "categories": cat_range,
            "values": val_r,
            "fill": {"color": colors[i % len(colors)]},
            "gap": 50,
            "data_labels": {"value": False},
        })

    chart.set_x_axis({
        "name_font": {"size": 9, "italic": True},
        "num_font": {"size": 8},
        "line": {"color": "#CCCCCC"},
        "major_gridlines": {"visible": False},
    })
    chart.set_y_axis({
        "num_format": "#,##0",
        "num_font": {"size": 8},
        "major_gridlines": {
            "visible": True,
            "line": {"color": "#E5E7EB", "dash_type": "dash"},
        },
        "line": {"none": True},
    })
    _base_style(chart, config, theme, zone)
    return chart


def build_line_chart(wb, config: ChartConfig, theme: Theme,
                     cat_range: str, val_ranges: list[tuple[str, str]],
                     zone: ChartZone, smooth: bool = True):
    """Smooth or straight line chart."""
    chart = wb.add_chart({"type": "line"})
    colors = [_hex(c) for c in theme.chart_colors]

    for i, (name_r, val_r) in enumerate(val_ranges):
        c = colors[i % len(colors)]
        chart.add_series({
            "name": name_r,
            "categories": cat_range,
            "values": val_r,
            "line": {"color": c, "width": 2.5, "smooth": smooth},
            "marker": {
                "type": "circle", "size": 5,
                "fill": {"color": c},
                "border": {"none": True},
            },
        })

    chart.set_x_axis({
        "num_font": {"size": 8},
        "line": {"color": "#CCCCCC"},
        "major_gridlines": {"visible": False},
    })
    chart.set_y_axis({
        "num_format": "#,##0",
        "num_font": {"size": 8},
        "major_gridlines": {
            "visible": True,
            "line": {"color": "#E5E7EB", "dash_type": "dash"},
        },
        "line": {"none": True},
    })
    _base_style(chart, config, theme, zone)
    return chart


def build_area_chart(wb, config: ChartConfig, theme: Theme,
                     cat_range: str, val_ranges: list[tuple[str, str]],
                     zone: ChartZone):
    """Stacked area chart."""
    chart = wb.add_chart({"type": "area", "subtype": "stacked"})
    colors = [_hex(c) for c in theme.chart_colors]

    for i, (name_r, val_r) in enumerate(val_ranges):
        c = colors[i % len(colors)]
        chart.add_series({
            "name": name_r,
            "categories": cat_range,
            "values": val_r,
            "fill": {"color": c, "transparency": 25},
            "line": {"color": c, "width": 1.5},
        })

    chart.set_x_axis({"num_font": {"size": 8}, "line": {"color": "#CCCCCC"}})
    chart.set_y_axis({
        "num_format": "#,##0",
        "num_font": {"size": 8},
        "major_gridlines": {"visible": True, "line": {"color": "#E5E7EB"}},
        "line": {"none": True},
    })
    _base_style(chart, config, theme, zone)
    return chart


def build_pie_chart(wb, config: ChartConfig, theme: Theme,
                    cat_range: str, val_range: str, series_name: str,
                    zone: ChartZone):
    """Pie chart with percentage labels."""
    chart = wb.add_chart({"type": "pie"})
    colors = [_hex(c) for c in theme.chart_colors]

    chart.add_series({
        "name": series_name,
        "categories": cat_range,
        "values": val_range,
        "data_labels": {"percentage": True, "font": {"size": 8}},
        "points": [{"fill": {"color": colors[i % len(colors)]}}
                   for i in range(20)],
    })

    chart.set_legend({"position": "right", "font": {"size": 9}})
    _base_style(chart, config, theme, zone)
    return chart


def build_doughnut_chart(wb, config: ChartConfig, theme: Theme,
                          cat_range: str, val_range: str, series_name: str,
                          zone: ChartZone):
    """Doughnut chart."""
    chart = wb.add_chart({"type": "doughnut"})
    colors = [_hex(c) for c in theme.chart_colors]

    chart.add_series({
        "name": series_name,
        "categories": cat_range,
        "values": val_range,
        "data_labels": {"percentage": True, "font": {"size": 8}},
        "points": [{"fill": {"color": colors[i % len(colors)]}}
                   for i in range(20)],
    })

    chart.set_legend({"position": "right", "font": {"size": 9}})
    _base_style(chart, config, theme, zone)
    return chart


def build_scatter_chart(wb, config: ChartConfig, theme: Theme,
                         x_range: str, y_range: str, series_name: str,
                         zone: ChartZone):
    """Scatter / bubble chart."""
    chart = wb.add_chart({"type": "scatter", "subtype": "straight_with_markers"})
    colors = [_hex(c) for c in theme.chart_colors]

    chart.add_series({
        "name": series_name,
        "categories": x_range,
        "values": y_range,
        "marker": {
            "type": "circle", "size": 6,
            "fill": {"color": colors[0]},
            "border": {"none": True},
        },
        "line": {"none": True},
    })

    chart.set_x_axis({"num_font": {"size": 8}})
    chart.set_y_axis({
        "num_format": "#,##0",
        "num_font": {"size": 8},
        "major_gridlines": {"visible": True, "line": {"color": "#E5E7EB"}},
    })
    _base_style(chart, config, theme, zone)
    return chart


# ── Dispatch ───────────────────────────────────────────────────────────────────

def build_xl_chart(wb, config: ChartConfig, theme: Theme,
                   cat_range: str, val_ranges: list[tuple[str, str]],
                   zone: ChartZone):
    """Dispatch to the right chart builder based on ChartConfig.type."""
    ctype = config.type

    # For pie/doughnut/scatter: use first val_range
    val_r = val_ranges[0][1] if val_ranges else "=Calculations!$B$2:$B$21"
    name_r = val_ranges[0][0] if val_ranges else f'"{config.title}"'

    if ctype == ChartType.BAR:
        return build_bar_chart(wb, config, theme, cat_range, val_ranges, zone)
    elif ctype == ChartType.BAR_HORIZONTAL:
        return build_bar_chart(wb, config, theme, cat_range, val_ranges, zone, horizontal=True)
    elif ctype == ChartType.LINE:
        return build_line_chart(wb, config, theme, cat_range, val_ranges, zone)
    elif ctype == ChartType.AREA:
        return build_area_chart(wb, config, theme, cat_range, val_ranges, zone)
    elif ctype == ChartType.PIE:
        return build_pie_chart(wb, config, theme, cat_range, val_r, name_r, zone)
    elif ctype == ChartType.DOUGHNUT:
        return build_doughnut_chart(wb, config, theme, cat_range, val_r, name_r, zone)
    elif ctype == ChartType.SCATTER:
        return build_scatter_chart(wb, config, theme, cat_range, val_r, name_r, zone)
    else:
        return build_bar_chart(wb, config, theme, cat_range, val_ranges, zone)
