import pytest

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.generator import generate_report
from report_generator.post_processing import PostProcessingRegistry
from report_generator.pptx.document import PptxDocument


def test_generate_report_populates_text_and_table(simple_template_bytes):
    mapping = {
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
            },
            {
                "location": "table.top_risks",
                "semantic_description": "TOP问题与风险",
                "type": "Table",
                "config": {"order": ["风险类型", "风险描述"]},
                "data_source": {
                    "name": "risks",
                },
            },
        ],
    }
    payload = {
        "api_data": {"全量数据": {"项目概览": {"项目名称": "智慧园区"}}},
        "risks": [{"风险类型": "延期风险", "风险描述": "设备到货存在延期风险"}],
    }

    output = generate_report(simple_template_bytes, mapping, payload, PostProcessingRegistry())
    doc = PptxDocument.open(output)
    index = doc.shape_index()

    assert index["text.report_title"].shape.text_frame.text == "智慧园区-报告"
    assert index["table.top_risks"].shape.table.cell(1, 1).text == "设备到货存在延期风险"


def test_generate_report_rejects_missing_component(simple_template_bytes):
    mapping = {
        "template_id": "project-monthly-report-ppt-v1",
        "component_list": [
            {
                "location": "text.missing",
                "semantic_description": "缺失组件",
                "type": "Text",
                "data_source": {"name": "api_data"},
            }
        ],
    }

    with pytest.raises(ReportGenerationError) as exc:
        generate_report(simple_template_bytes, mapping, {"api_data": "value"}, PostProcessingRegistry())

    assert exc.value.error_code == ErrorCode.COMPONENT_NOT_FOUND
