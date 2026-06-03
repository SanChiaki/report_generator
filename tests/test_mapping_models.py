import pytest
from pydantic import ValidationError

from report_generator.models import ComponentMapping, DataSource, ReportMapping


def test_report_mapping_accepts_excel_compatible_shape():
    mapping = ReportMapping.model_validate(
        {
            "template_id": "project-monthly-report-ppt-v1",
            "component_list": [
                {
                    "location": "text.report_title",
                    "semantic_description": "报告标题",
                    "type": "Text",
                    "data_source": {
                        "name": "api_data",
                        "template": "{{ 全量数据.项目概览.项目名称 }}-报告",
                    },
                    "config": {"min_font_size": 12},
                }
            ],
        }
    )

    assert mapping.template_id == "project-monthly-report-ppt-v1"
    assert mapping.component_list[0].location == "text.report_title"
    assert mapping.component_list[0].type == "Text"
    assert mapping.component_list[0].data_source.template.startswith("{{")


def test_component_requires_known_type():
    with pytest.raises(ValidationError):
        ComponentMapping.model_validate(
            {
                "location": "unknown.component",
                "semantic_description": "未知组件",
                "type": "Unknown",
            }
        )


def test_data_source_keeps_params_and_post_processing_flag():
    source = DataSource.model_validate(
        {
            "name": "general_top_risks_and_issues",
            "params": {"sr_api_data": "sr_api_data"},
            "needs_post_processing": True,
        }
    )

    assert source.name == "general_top_risks_and_issues"
    assert source.params == {"sr_api_data": "sr_api_data"}
    assert source.needs_post_processing is True
