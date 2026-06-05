import pytest
from pptx.util import Pt

from report_generator.components.text import apply_text
from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping
from report_generator.pptx.document import PptxDocument


def text_component(**config):
    return ComponentMapping.model_validate(
        {
            "location": "text.summary",
            "semantic_description": "整体状态",
            "type": "Text",
            "config": config,
        }
    )


def test_apply_text_replaces_content(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["text.summary"].shape

    apply_text(shape, text_component(min_font_size=10), "新的摘要")

    assert shape.text_frame.text == "新的摘要"


def test_apply_text_preserves_existing_run_style(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["text.summary"].shape
    original_run = shape.text_frame.paragraphs[0].runs[0]
    original_size = original_run.font.size

    apply_text(shape, text_component(preserve_style=True), "保留样式的新摘要")

    updated_run = shape.text_frame.paragraphs[0].runs[0]
    assert shape.text_frame.text == "保留样式的新摘要"
    assert updated_run.font.size == original_size


def test_apply_text_preserve_style_rejects_when_original_font_cannot_fit(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["text.summary"].shape

    with pytest.raises(ReportGenerationError) as exc:
        apply_text(shape, text_component(preserve_style=True), "超长内容" * 80)

    assert exc.value.error_code == ErrorCode.TEXT_OVERFLOW


def test_apply_text_shrinks_long_content(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["text.summary"].shape

    apply_text(shape, text_component(min_font_size=8), "这是一段较长的摘要内容，用来触发字号缩小逻辑。" * 3)

    assert shape.text_frame.paragraphs[0].runs[0].font.size <= Pt(18)
    assert shape.text_frame.paragraphs[0].runs[0].font.size >= Pt(8)


def test_apply_text_raises_when_content_cannot_fit(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["text.summary"].shape

    with pytest.raises(ReportGenerationError) as exc:
        apply_text(shape, text_component(min_font_size=16), "超长内容" * 400)

    assert exc.value.error_code == ErrorCode.TEXT_OVERFLOW
