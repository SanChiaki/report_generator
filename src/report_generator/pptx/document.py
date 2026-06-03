from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from typing import Any

from pptx import Presentation

from report_generator.errors import ErrorCode, ReportGenerationError


@dataclass(frozen=True)
class ShapeRef:
    name: str
    slide_index: int
    shape: Any
    kind: str


class PptxDocument:
    def __init__(self, presentation: Presentation) -> None:
        self.presentation = presentation

    @classmethod
    def open(cls, template_bytes: bytes) -> "PptxDocument":
        try:
            return cls(Presentation(BytesIO(template_bytes)))
        except Exception as exc:
            raise ReportGenerationError(
                ErrorCode.PPTX_PARSE_FAILED,
                f"PPTX 模板解析失败: {exc}",
            ) from exc

    def shape_index(self, required_names: set[str] | None = None) -> dict[str, ShapeRef]:
        index: dict[str, ShapeRef] = {}
        duplicates: list[str] = []
        for slide_index, slide in enumerate(self.presentation.slides):
            for shape in slide.shapes:
                name = getattr(shape, "name", "")
                if not name:
                    continue
                if required_names is not None and name not in required_names:
                    continue
                if name in index:
                    duplicates.append(name)
                    continue
                index[name] = ShapeRef(
                    name=name,
                    slide_index=slide_index,
                    shape=shape,
                    kind=infer_shape_kind(shape),
                )
        if duplicates:
            raise ReportGenerationError(
                ErrorCode.DUPLICATE_COMPONENT_NAME,
                "PPT 模板中存在重复组件名称",
                details={"duplicates": sorted(set(duplicates))},
            )
        return index

    def remove_shape(self, shape: Any) -> None:
        element = shape._element
        element.getparent().remove(element)

    def to_bytes(self) -> bytes:
        output = BytesIO()
        self.presentation.save(output)
        return output.getvalue()


def infer_shape_kind(shape: Any) -> str:
    if getattr(shape, "has_table", False):
        return "Table"
    if getattr(shape, "has_chart", False):
        return "Chart"
    if getattr(shape, "has_text_frame", False):
        return "Text"
    if getattr(shape, "shape_type", None) is not None:
        return "Shape"
    return "Unknown"
