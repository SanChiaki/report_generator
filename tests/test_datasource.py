import pytest

from report_generator.datasource import resolve_component_value
from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping
from report_generator.post_processing import PostProcessingRegistry


def component(data_source):
    return ComponentMapping.model_validate(
        {
            "location": "text.report_title",
            "semantic_description": "报告标题",
            "prompt": "提炼成一句汇报标题",
            "data_example": "智慧园区项目周报",
            "type": "Text",
            "data_source": data_source,
        }
    )


class FakeComponentProcessor:
    def __init__(self, result):
        self.result = result
        self.calls = []

    def process(self, component, value):
        self.calls.append((component, value))
        return self.result


def test_resolves_jsonpath_against_named_source():
    payload = {
        "api_data": {
            "全量数据": {
                "项目概览": {
                    "项目名称": "智慧园区",
                }
            }
        }
    }
    result = resolve_component_value(
        component(
            {
                "name": "api_data",
                "index": "$['全量数据']['项目概览']['项目名称']",
            }
        ),
        payload,
        PostProcessingRegistry(),
    )

    assert result == "智慧园区"


def test_renders_template_against_named_source():
    payload = {
        "api_data": {
            "全量数据": {
                "项目概览": {
                    "项目名称": "智慧园区",
                }
            }
        }
    }
    result = resolve_component_value(
        component(
            {
                "name": "api_data",
                "template": "{{ 全量数据.项目概览.项目名称 }}-报告",
            }
        ),
        payload,
        PostProcessingRegistry(),
    )

    assert result == "智慧园区-报告"


def test_calls_named_post_processing_function_with_params():
    registry = PostProcessingRegistry()
    registry.register("join_names", lambda users: ", ".join(item["name"] for item in users))
    payload = {"users": [{"name": "张三"}, {"name": "李四"}]}

    result = resolve_component_value(
        component(
            {
                "name": "join_names",
                "params": {"users": "users"},
                "needs_post_processing": True,
            }
        ),
        payload,
        registry,
    )

    assert result == "张三, 李四"


def test_needs_post_processing_uses_component_processor_with_resolved_data():
    processor = FakeComponentProcessor("智慧园区项目周报")
    payload = {"api_data": {"raw": {"项目名称": "智慧园区", "状态": "正常"}}}
    target = component(
        {
            "name": "api_data",
            "index": "$.raw",
            "needs_post_processing": True,
        }
    )

    result = resolve_component_value(
        target,
        payload,
        PostProcessingRegistry(),
        processor,
    )

    assert result == "智慧园区项目周报"
    called_component, called_value = processor.calls[0]
    assert called_component.semantic_description == "报告标题"
    assert called_component.prompt == "提炼成一句汇报标题"
    assert called_component.data_example == "智慧园区项目周报"
    assert called_value == {"项目名称": "智慧园区", "状态": "正常"}


def test_missing_source_raises_structured_error():
    with pytest.raises(ReportGenerationError) as exc:
        resolve_component_value(
            component({"name": "api_data", "index": "$.missing"}),
            {},
            PostProcessingRegistry(),
        )

    assert exc.value.error_code == ErrorCode.DATA_SOURCE_NOT_FOUND


def test_invalid_jsonpath_expression_raises_structured_error():
    with pytest.raises(ReportGenerationError) as exc:
        resolve_component_value(
            component({"name": "api_data", "index": "$["}),
            {"api_data": {"title": "智慧园区"}},
            PostProcessingRegistry(),
        )

    assert exc.value.error_code == ErrorCode.DATA_SOURCE_INVALID


def test_missing_template_variable_raises_structured_error():
    with pytest.raises(ReportGenerationError) as exc:
        resolve_component_value(
            component({"name": "api_data", "template": "{{ missing_title }}"}),
            {"api_data": {"title": "智慧园区"}},
            PostProcessingRegistry(),
        )

    assert exc.value.error_code == ErrorCode.DATA_SOURCE_NOT_FOUND


def test_malformed_template_syntax_raises_structured_error():
    with pytest.raises(ReportGenerationError) as exc:
        resolve_component_value(
            component({"name": "api_data", "template": "{{ title "}),
            {"api_data": {"title": "智慧园区"}},
            PostProcessingRegistry(),
        )

    assert exc.value.error_code == ErrorCode.DATA_SOURCE_INVALID
