from report_generator.components.chart import apply_chart
from report_generator.components.shape import apply_shape
from report_generator.models import ComponentMapping
from report_generator.pptx.document import PptxDocument


def chart_component(**config):
    return ComponentMapping.model_validate(
        {
            "location": "chart.revenue_trend",
            "semantic_description": "收入趋势",
            "type": "Chart",
            "config": config,
        }
    )


def shape_component(**config):
    return ComponentMapping.model_validate(
        {
            "location": "shape.status_badge",
            "semantic_description": "项目状态",
            "type": "Shape",
            "config": config,
        }
    )


def test_apply_chart_replaces_native_chart_data(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["chart.revenue_trend"].shape
    value = {
        "categories": ["Q1", "Q2", "Q3"],
        "series": [{"name": "收入", "values": [10, 20, 30]}],
    }

    apply_chart(shape, chart_component(max_categories=4, max_series=2), value)

    reopened = PptxDocument.open(doc.to_bytes())
    chart = reopened.shape_index()["chart.revenue_trend"].shape.chart
    assert len(chart.series[0].points) == 3


def test_apply_shape_updates_text(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["shape.status_badge"].shape

    apply_shape(shape, shape_component(), "风险")

    assert shape.text == "风险"
