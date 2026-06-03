from __future__ import annotations

from typing import Any

from pptx.dml.color import RGBColor

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping


def apply_shape(shape: Any, component: ComponentMapping, value: Any) -> None:
    if value is not None and hasattr(shape, "text"):
        shape.text = str(value)

    try:
        styles = component.config.get("state_styles", {})
        state_style = styles.get(str(value), {}) if isinstance(styles, dict) else {}
        fill = state_style.get("fill") or component.config.get("fill")
        if fill and hasattr(shape, "fill"):
            shape.fill.solid()
            shape.fill.fore_color.rgb = _rgb(fill)

        line = state_style.get("line") or component.config.get("line")
        if line and hasattr(shape, "line"):
            shape.line.color.rgb = _rgb(line)
    except Exception as exc:
        raise ReportGenerationError(
            ErrorCode.DATA_SOURCE_INVALID,
            f"组件 {component.location} 的形状样式配置无效: {exc}",
            component,
        ) from exc


def _rgb(value: str) -> RGBColor:
    if not isinstance(value, str):
        raise ValueError(f"invalid RGB color {value!r}")
    normalized = value.lstrip("#")
    if len(normalized) != 6 or any(char not in "0123456789abcdefABCDEF" for char in normalized):
        raise ValueError(f"invalid RGB color {value!r}")
    return RGBColor(
        int(normalized[0:2], 16),
        int(normalized[2:4], 16),
        int(normalized[4:6], 16),
    )
