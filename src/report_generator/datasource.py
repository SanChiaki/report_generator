from __future__ import annotations

from typing import Any

from jinja2 import StrictUndefined
from jinja2.exceptions import TemplateError, TemplateSyntaxError, UndefinedError
from jinja2.sandbox import SandboxedEnvironment
from jsonpath_ng.exceptions import JSONPathError
from jsonpath_ng import parse

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping, DataSource
from report_generator.post_processing import PostProcessingRegistry


def resolve_component_value(
    component: ComponentMapping,
    payload: dict[str, Any],
    registry: PostProcessingRegistry,
) -> Any:
    source = component.data_source
    if source is None:
        return None

    if source.needs_post_processing:
        return _resolve_post_processed(component, source, payload, registry)

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
) -> Any:
    if not source.name or not registry.has(source.name):
        raise ReportGenerationError(
            ErrorCode.POST_PROCESSING_FAILED,
            f"组件 {component.location} 引用了未注册的后处理函数 {source.name}",
            component,
        )
    params: dict[str, Any] = {}
    for param_name, payload_key in source.params.items():
        if payload_key not in payload:
            raise ReportGenerationError(
                ErrorCode.DATA_SOURCE_NOT_FOUND,
                f"组件 {component.location} 的参数 {param_name} 引用了不存在的数据 {payload_key}",
                component,
            )
        params[param_name] = payload[payload_key]
    try:
        return registry.call(source.name, **params)
    except Exception as exc:
        raise ReportGenerationError(
            ErrorCode.POST_PROCESSING_FAILED,
            f"组件 {component.location} 的后处理失败: {exc}",
            component,
        ) from exc
