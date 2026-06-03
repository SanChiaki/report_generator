from __future__ import annotations

from typing import Any

from report_generator.components.chart import apply_chart
from report_generator.components.image import apply_image
from report_generator.components.shape import apply_shape
from report_generator.components.table import apply_table
from report_generator.components.text import apply_text
from report_generator.datasource import resolve_component_value
from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping, ReportMapping
from report_generator.post_processing import PostProcessingRegistry
from report_generator.pptx.document import PptxDocument, ShapeRef


def generate_report(
    template_bytes: bytes,
    mapping_payload: dict[str, Any],
    business_payload: dict[str, Any],
    registry: PostProcessingRegistry | None = None,
) -> bytes:
    registry = registry or PostProcessingRegistry()
    mapping = ReportMapping.model_validate(mapping_payload)
    doc = PptxDocument.open(template_bytes)
    index = doc.shape_index()

    for component in mapping.component_list:
        ref = _find_component(component, index)
        if component.visible is False:
            doc.remove_shape(ref.shape)
            continue
        value = resolve_component_value(component, business_payload, registry)
        _apply_component(doc, ref, component, value)
        index = doc.shape_index()

    return doc.to_bytes()


def _find_component(component: ComponentMapping, index: dict[str, ShapeRef]) -> ShapeRef:
    ref = index.get(component.location)
    if ref is None:
        raise ReportGenerationError(
            ErrorCode.COMPONENT_NOT_FOUND,
            f"模板中未找到组件 {component.location}",
            component,
        )
    return ref


def _apply_component(
    doc: PptxDocument,
    ref: ShapeRef,
    component: ComponentMapping,
    value: Any,
) -> None:
    if component.type == "Text":
        apply_text(ref.shape, component, value)
        return
    if component.type == "Image":
        apply_image(doc, ref.shape, component, value)
        return
    if component.type == "Table":
        apply_table(doc, ref.shape, component, value)
        return
    if component.type == "Chart":
        apply_chart(ref.shape, component, value)
        return
    if component.type == "Shape":
        apply_shape(ref.shape, component, value)
        return
    raise ReportGenerationError(
        ErrorCode.TYPE_MISMATCH,
        f"组件 {component.location} 的类型 {component.type} 尚未支持",
        component,
    )
