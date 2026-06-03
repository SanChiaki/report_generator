from __future__ import annotations

from typing import Any

from pptx.dml.color import RGBColor

from report_generator.models import ComponentMapping


def apply_shape(shape: Any, component: ComponentMapping, value: Any) -> None:
    if value is not None and hasattr(shape, "text"):
        shape.text = str(value)

    styles = component.config.get("state_styles", {})
    state_style = styles.get(str(value), {}) if isinstance(styles, dict) else {}
    fill = state_style.get("fill") or component.config.get("fill")
    if fill and hasattr(shape, "fill"):
        shape.fill.solid()
        shape.fill.fore_color.rgb = _rgb(fill)

    line = state_style.get("line") or component.config.get("line")
    if line and hasattr(shape, "line"):
        shape.line.color.rgb = _rgb(line)


def _rgb(value: str) -> RGBColor:
    normalized = value.lstrip("#")
    return RGBColor(
        int(normalized[0:2], 16),
        int(normalized[2:4], 16),
        int(normalized[4:6], 16),
    )
