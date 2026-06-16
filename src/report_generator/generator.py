from __future__ import annotations

import os
from concurrent.futures import ThreadPoolExecutor
from typing import Any

from report_generator.components.chart import apply_chart
from report_generator.components.image import apply_image
from report_generator.components.milestone import apply_milestone
from report_generator.components.shape import apply_shape
from report_generator.components.table import apply_table
from report_generator.components.text import apply_text
from report_generator.components.top_issues import apply_top_issues
from report_generator.builtin_processors import default_registry
from report_generator.datasource import resolve_component_value
from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.llm import ComponentDataProcessor, OpenAICompletionProcessor
from report_generator.models import ComponentMapping, ReportMapping
from report_generator.post_processing import PostProcessingRegistry
from report_generator.pptx.document import PptxDocument, ShapeRef


SKIPPED_COMPONENT_VALUE = object()


def generate_report(
    template_bytes: bytes,
    mapping_payload: dict[str, Any],
    business_payload: dict[str, Any],
    registry: PostProcessingRegistry | None = None,
    processor: ComponentDataProcessor | None = None,
    llm_concurrency: int | None = None,
) -> bytes:
    registry = registry or default_registry()
    processor = processor or OpenAICompletionProcessor()
    llm_concurrency = _resolve_llm_concurrency(llm_concurrency)
    mapping = ReportMapping.model_validate(mapping_payload)
    doc = PptxDocument.open(template_bytes)
    required_names = {component.location for component in mapping.component_list}
    index = doc.shape_index(required_names=required_names)
    for component in mapping.component_list:
        _find_component(component, index)
    values = _resolve_component_values(mapping.component_list, business_payload, registry, processor, llm_concurrency)

    for component, value in zip(mapping.component_list, values, strict=True):
        ref = _find_component(component, index)
        if component.visible is False:
            doc.remove_shape(ref.shape)
            index = doc.shape_index(required_names=required_names)
            continue
        if value is SKIPPED_COMPONENT_VALUE:
            continue
        _apply_component(doc, ref, component, value)
        index = doc.shape_index(required_names=required_names)

    return doc.to_bytes()


def _resolve_component_values(
    components: list[ComponentMapping],
    business_payload: dict[str, Any],
    registry: PostProcessingRegistry,
    processor: ComponentDataProcessor,
    llm_concurrency: int,
) -> list[Any]:
    values: list[Any] = [None] * len(components)
    post_processing_indexes = [
        index
        for index, component in enumerate(components)
        if component.visible is not False
        and component.data_source is not None
        and component.data_source.needs_post_processing
    ]

    for index, component in enumerate(components):
        if index in post_processing_indexes or component.visible is False:
            continue
        values[index] = _resolve_component_value_or_skip(component, business_payload, registry, processor)

    if not post_processing_indexes:
        return values

    with ThreadPoolExecutor(max_workers=llm_concurrency) as executor:
        futures = {
            executor.submit(
                resolve_component_value,
                components[index],
                business_payload,
                registry,
                processor,
            ): index
            for index in post_processing_indexes
        }
        for future, index in futures.items():
            values[index] = _future_result_or_skip(future)

    return values


def _resolve_component_value_or_skip(
    component: ComponentMapping,
    business_payload: dict[str, Any],
    registry: PostProcessingRegistry,
    processor: ComponentDataProcessor,
) -> Any:
    try:
        return resolve_component_value(component, business_payload, registry, processor)
    except ReportGenerationError as exc:
        if _should_skip_component_for_value_error(exc):
            return SKIPPED_COMPONENT_VALUE
        raise


def _future_result_or_skip(future: Any) -> Any:
    try:
        return future.result()
    except ReportGenerationError as exc:
        if _should_skip_component_for_value_error(exc):
            return SKIPPED_COMPONENT_VALUE
        raise


def _should_skip_component_for_value_error(exc: ReportGenerationError) -> bool:
    return exc.error_code in {
        ErrorCode.DATA_SOURCE_NOT_FOUND,
        ErrorCode.DATA_SOURCE_INVALID,
        ErrorCode.POST_PROCESSING_FAILED,
    }


def _resolve_llm_concurrency(value: int | None) -> int:
    if value is None:
        raw = os.getenv("LLM_CONCURRENCY", "4")
        try:
            value = int(raw)
        except ValueError:
            value = 4
    return max(1, value)


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
    if component.type == "TopIssues":
        apply_top_issues(doc, ref.shape, component, value)
        return
    if component.type == "Milestone":
        apply_milestone(doc, ref.shape, component, value)
        return
    raise ReportGenerationError(
        ErrorCode.TYPE_MISMATCH,
        f"组件 {component.location} 的类型 {component.type} 尚未支持",
        component,
    )
