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
            "type": "Text",
            "data_source": data_source,
        }
    )


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
