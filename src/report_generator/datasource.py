from __future__ import annotations

from typing import Any

from jinja2 import StrictUndefined
from jinja2.exceptions import TemplateError, TemplateSyntaxError, UndefinedError
from jinja2.sandbox import SandboxedEnvironment
from jsonpath_ng.exceptions import JSONPathError
from jsonpath_ng import parse

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.llm import ComponentDataProcessor
from report_generator.models import ComponentMapping, DataSource
from report_generator.post_processing import PostProcessingRegistry


def resolve_component_value(
    component: ComponentMapping,
    payload: dict[str, Any],
    registry: PostProcessingRegistry,
    processor: ComponentDataProcessor | None = None,
) -> Any:
    source = component.data_source
    if source is None:
        return None

    if source.name and registry.has(source.name):
        value = _resolve_registered_processor(component, source, payload, registry)
        if source.needs_post_processing:
            return _process_value(component, value, processor)
        return value

    if source.needs_post_processing:
        return _resolve_post_processed(component, source, payload, registry, processor)

    if source.template:
        base = _source_base(component, source, payload)
        return _render_template(component, source.template, base)

    if source.index:
        base = _source_base(component, source, payload)
        return _resolve_jsonpath(component, source.index, base)

    if source.name and source.name in payload:
        return payload[source.name]

    raise ReportGenerationError(
        ErrorCode.DATA_SOURCE_INVALID,
        f"组件 {component.location} 的 data_source 缺少 index、template 或可用的 name",
        component,
    )


def _source_base(component: ComponentMapping, source: DataSource, payload: dict[str, Any]) -> Any:
    if source.name is None:
        return payload
    if source.name not in payload:
        raise ReportGenerationError(
            ErrorCode.DATA_SOURCE_NOT_FOUND,
            f"组件 {component.location} 引用了不存在的数据源 {source.name}",
            component,
        )
    return payload[source.name]


def _resolve_jsonpath(component: ComponentMapping, expression: str, base: Any) -> Any:
    try:
        matches = [match.value for match in parse(expression).find(base)]
    except JSONPathError as exc:
        raise ReportGenerationError(
            ErrorCode.DATA_SOURCE_INVALID,
            f"组件 {component.location} 的 JSONPath 表达式无效: {expression}",
            component,
        ) from exc
    if not matches:
        raise ReportGenerationError(
            ErrorCode.DATA_SOURCE_NOT_FOUND,
            f"组件 {component.location} 的 JSONPath 没有匹配到数据: {expression}",
            component,
        )
    if len(matches) == 1:
        return matches[0]
    return matches


def _render_template(component: ComponentMapping, template: str, base: Any) -> str:
    env = SandboxedEnvironment(undefined=StrictUndefined, autoescape=False)
    try:
        compiled = env.from_string(template)
    except TemplateSyntaxError as exc:
        raise ReportGenerationError(
            ErrorCode.DATA_SOURCE_INVALID,
            f"组件 {component.location} 的模板语法无效: {exc}",
            component,
        ) from exc
    try:
        return compiled.render(base)
    except UndefinedError as exc:
        raise ReportGenerationError(
            ErrorCode.DATA_SOURCE_NOT_FOUND,
            f"组件 {component.location} 的模板变量未找到: {exc}",
            component,
        ) from exc
    except TemplateError as exc:
        raise ReportGenerationError(
            ErrorCode.DATA_SOURCE_INVALID,
            f"组件 {component.location} 的模板渲染失败: {exc}",
            component,
        ) from exc


def _resolve_post_processed(
    component: ComponentMapping,
    source: DataSource,
    payload: dict[str, Any],
    registry: PostProcessingRegistry,
    processor: ComponentDataProcessor | None,
) -> Any:
    return _process_value(component, _resolve_pre_processed_value(component, source, payload), processor)


def _process_value(
    component: ComponentMapping,
    value: Any,
    processor: ComponentDataProcessor | None,
) -> Any:
    if processor is None:
        raise ReportGenerationError(
            ErrorCode.POST_PROCESSING_FAILED,
            f"组件 {component.location} 需要大模型后处理，但未配置组件数据处理器",
            component,
        )
    return processor.process(component, value)


def _resolve_registered_processor(
    component: ComponentMapping,
    source: DataSource,
    payload: dict[str, Any],
    registry: PostProcessingRegistry,
) -> Any:
    params: dict[str, Any] = {}
    for param_name, payload_key in source.params.items():
        params[param_name] = _resolve_processor_param(component, param_name, payload_key, payload)
    try:
        return registry.call(str(source.name), **params)
    except Exception as exc:
        raise ReportGenerationError(
            ErrorCode.POST_PROCESSING_FAILED,
            f"组件 {component.location} 的后处理失败: {exc}",
            component,
        ) from exc


def _resolve_processor_param(
    component: ComponentMapping,
    param_name: str,
    payload_key: Any,
    payload: dict[str, Any],
) -> Any:
    if payload_key == "$":
        return payload
    if isinstance(payload_key, str) and payload_key in payload:
        return payload[payload_key]
    if param_name == "sr_api_data" and payload_key == "sr_api_data":
        return payload
    raise ReportGenerationError(
        ErrorCode.DATA_SOURCE_NOT_FOUND,
        f"组件 {component.location} 的参数 {param_name} 引用了不存在的数据 {payload_key}",
        component,
    )


def _resolve_pre_processed_value(component: ComponentMapping, source: DataSource, payload: dict[str, Any]) -> Any:
    if source.template:
        base = _source_base(component, source, payload)
        return _render_template(component, source.template, base)
    if source.index:
        base = _source_base(component, source, payload)
        return _resolve_jsonpath(component, source.index, base)
    if source.name:
        if source.name not in payload:
            raise ReportGenerationError(
                ErrorCode.DATA_SOURCE_NOT_FOUND,
                f"组件 {component.location} 引用了不存在的数据源 {source.name}",
                component,
            )
        return payload[source.name]
    return payload
