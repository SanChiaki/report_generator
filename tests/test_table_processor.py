import pytest

from report_generator.components.table import apply_table
from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping
from report_generator.pptx.document import PptxDocument


def table_component(**config):
    return ComponentMapping.model_validate(
        {
            "location": "table.top_risks",
            "semantic_description": "TOP问题与风险",
            "type": "Table",
            "config": config,
        }
    )


def test_apply_table_rebuilds_rows_and_columns(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["table.top_risks"].shape
    data = [
        {"风险类型": "延期风险", "风险描述": "设备到货存在延期风险"},
        {"风险类型": "质量风险", "风险描述": "测试通过率偏低"},
    ]

    new_shape = apply_table(
        doc,
        shape,
        table_component(order=["风险类型", "风险描述"], min_font_size=8),
        data,
    )

    table = new_shape.table
    assert len(table.rows) == 3
    assert len(table.columns) == 2
    assert table.cell(0, 0).text == "风险类型"
    assert table.cell(2, 1).text == "测试通过率偏低"


def test_apply_table_rejects_too_many_rows(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["table.top_risks"].shape
    data = [{"风险类型": str(index), "风险描述": "x"} for index in range(3)]

    with pytest.raises(ReportGenerationError) as exc:
        apply_table(
            doc,
            shape,
            table_component(order=["风险类型", "风险描述"], max_rows=2),
            data,
        )

    assert exc.value.error_code == ErrorCode.TABLE_OVERFLOW


def test_apply_table_accepts_columns_rows_object(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["table.top_risks"].shape
    value = {
        "columns": [
            {"key": "type", "label": "风险类型"},
            {"key": "desc", "label": "风险描述"},
        ],
        "rows": [{"type": "延期风险", "desc": "设备到货存在延期风险"}],
    }

    new_shape = apply_table(doc, shape, table_component(), value)

    assert new_shape.table.cell(0, 0).text == "风险类型"
    assert new_shape.table.cell(1, 1).text == "设备到货存在延期风险"


def test_apply_table_rejects_non_object_rows(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["table.top_risks"].shape

    with pytest.raises(ReportGenerationError) as exc:
        apply_table(doc, shape, table_component(order=["风险类型"]), ["not-a-row"])

    assert exc.value.error_code == ErrorCode.DATA_SOURCE_INVALID


def test_apply_table_rejects_non_object_columns(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["table.top_risks"].shape
    value = {"columns": ["bad"], "rows": [{"bad": "value"}]}

    with pytest.raises(ReportGenerationError) as exc:
        apply_table(doc, shape, table_component(), value)

    assert exc.value.error_code == ErrorCode.DATA_SOURCE_INVALID
