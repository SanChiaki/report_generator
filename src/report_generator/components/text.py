from __future__ import annotations

import math
from typing import Any

from pptx.util import Pt

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping

EMU_PER_INCH = 914400


def apply_text(shape: Any, component: ComponentMapping, value: Any) -> None:
    if not getattr(shape, "has_text_frame", False):
        raise ReportGenerationError(
            ErrorCode.TYPE_MISMATCH,
            f"组件 {component.location} 不是文本组件",
            component,
        )

    text = "" if value is None else str(value)
    min_font_size = int(component.config.get("min_font_size", 10))
    start_font_size = _existing_font_size(shape) or int(component.config.get("font_size", 18))
    fitted_font_size = _fit_font_size(shape, text, start_font_size, min_font_size)
    if fitted_font_size is None:
        raise ReportGenerationError(
            ErrorCode.TEXT_OVERFLOW,
            f"组件 {component.location} 的文本在最小字号 {min_font_size} 下仍无法放入模板区域",
            component,
        )

    text_frame = shape.text_frame
    text_frame.clear()
    lines = text.splitlines() or [""]
    for index, line in enumerate(lines):
        paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
        run = paragraph.add_run()
        run.text = line
        run.font.size = Pt(fitted_font_size)


def _existing_font_size(shape: Any) -> int | None:
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if run.font.size is not None:
                return int(round(run.font.size.pt))
    return None


def _fit_font_size(shape: Any, text: str, start: int, minimum: int) -> int | None:
    for font_size in range(start, minimum - 1, -1):
        if _fits(shape, text, font_size):
            return font_size
    return None


def _fits(shape: Any, text: str, font_size: int) -> bool:
    width_in = max(shape.width / EMU_PER_INCH, 0.1)
    height_in = max(shape.height / EMU_PER_INCH, 0.1)
    chars_per_line = max(1, int((width_in * 72) / (font_size * 0.55)))
    source_lines = text.splitlines() or [""]
    needed_lines = 0
    for line in source_lines:
        needed_lines += max(1, math.ceil(len(line) / chars_per_line))
    line_height_in = (font_size * 1.25) / 72
    capacity = max(1, math.floor(height_in / line_height_in))
    return needed_lines <= capacity
