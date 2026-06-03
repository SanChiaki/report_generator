from __future__ import annotations

import base64
from io import BytesIO
from typing import Any

import httpx

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping
from report_generator.pptx.document import PptxDocument


def apply_image(
    doc: PptxDocument,
    shape: Any,
    component: ComponentMapping,
    value: Any,
) -> Any:
    image_bytes = _load_image_bytes(component, value)
    x, y, cx, cy = shape.left, shape.top, shape.width, shape.height
    slide = shape.part.slide
    doc.remove_shape(shape)
    new_shape = slide.shapes.add_picture(BytesIO(image_bytes), x, y, width=cx, height=cy)
    new_shape.name = component.location
    return new_shape


def _load_image_bytes(component: ComponentMapping, value: Any) -> bytes:
    src = value.get("src") if isinstance(value, dict) else value
    if not isinstance(src, str) or not src:
        raise ReportGenerationError(
            ErrorCode.IMAGE_LOAD_FAILED,
            f"组件 {component.location} 的图片地址为空",
            component,
        )

    try:
        if src.startswith("data:"):
            return base64.b64decode(src.split(",", 1)[1])
        if src.startswith("http://") or src.startswith("https://"):
            response = httpx.get(src, timeout=10)
            response.raise_for_status()
            return response.content
        with open(src, "rb") as handle:
            return handle.read()
    except Exception as exc:
        raise ReportGenerationError(
            ErrorCode.IMAGE_LOAD_FAILED,
            f"组件 {component.location} 的图片加载失败: {exc}",
            component,
        ) from exc
