import pytest

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.pptx.document import PptxDocument


def test_scans_named_shapes(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    index = doc.shape_index()

    assert "text.report_title" in index
    assert "table.top_risks" in index
    assert "chart.revenue_trend" in index


def test_duplicate_shape_names_raise_error(duplicate_name_template_bytes):
    doc = PptxDocument.open(duplicate_name_template_bytes)

    with pytest.raises(ReportGenerationError) as exc:
        doc.shape_index()

    assert exc.value.error_code == ErrorCode.DUPLICATE_COMPONENT_NAME
