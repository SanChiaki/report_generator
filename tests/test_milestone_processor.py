from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

from report_generator.generator import generate_report
from report_generator.post_processing import PostProcessingRegistry
from report_generator.pptx.document import PptxDocument


def _template_bytes() -> bytes:
    from io import BytesIO

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    anchor = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.7),
        Inches(1.0),
        Inches(8.0),
        Inches(1.2),
    )
    anchor.name = "milestone.delivery"
    anchor.text = "milestone anchor"
    anchor.fill.solid()
    anchor.fill.fore_color.rgb = RGBColor(0xFF, 0xF2, 0xCC)
    output = BytesIO()
    prs.save(output)
    return output.getvalue()


def test_milestone_draws_dynamic_nodes_from_items():
    mapping = {
        "template_id": "milestone-test",
        "component_list": [
            {
                "location": "milestone.delivery",
                "semantic_description": "项目里程碑",
                "type": "Milestone",
                "data_source": {"name": "milestones"},
            }
        ],
    }
    payload = {
        "milestones": {
            "items": [
                {"label": "准备", "date": "04-20", "status": "done"},
                {"label": "安装", "date": "05-18", "status": "done"},
                {"label": "调测", "date": "06-20", "status": "active"},
                {"label": "验收", "date": "07-05", "status": "pending"},
                {"label": "上线", "date": "07-20", "status": "pending"},
            ]
        }
    }

    output = generate_report(_template_bytes(), mapping, payload, PostProcessingRegistry())
    doc = PptxDocument.open(output)
    index = doc.shape_index()

    assert index["milestone.delivery"].shape.text == ""
    assert index["milestone.delivery"].shape.fill.fore_color.rgb == RGBColor(0xFF, 0xF2, 0xCC)
    nodes = [ref for name, ref in index.items() if name.startswith("milestone.delivery.item_") and name.endswith(".node")]
    dates = [ref for name, ref in index.items() if name.startswith("milestone.delivery.item_") and name.endswith(".date")]
    labels = [ref for name, ref in index.items() if name.startswith("milestone.delivery.item_") and name.endswith(".label")]
    assert len(nodes) == 5
    assert [ref.shape.text for ref in dates] == ["04-20", "05-18", "06-20", "07-05", "07-20"]
    assert [ref.shape.text for ref in labels] == ["准备", "安装", "调测", "验收", "上线"]
    assert nodes[1].shape.left > nodes[0].shape.left
    assert nodes[4].shape.left > nodes[3].shape.left


def test_milestone_uses_readable_default_text_sizes():
    mapping = {
        "template_id": "milestone-test",
        "component_list": [
            {
                "location": "milestone.delivery",
                "semantic_description": "项目里程碑",
                "type": "Milestone",
                "data_source": {"name": "milestones"},
            }
        ],
    }
    payload = {"milestones": [{"label": "准备", "date": "04-20", "status": "done"}]}

    output = generate_report(_template_bytes(), mapping, payload, PostProcessingRegistry())
    doc = PptxDocument.open(output)
    index = doc.shape_index()
    date_run = index["milestone.delivery.item_1.date"].shape.text_frame.paragraphs[0].runs[0]
    label_run = index["milestone.delivery.item_1.label"].shape.text_frame.paragraphs[0].runs[0]

    assert date_run.font.size.pt == 12
    assert label_run.font.size.pt == 14
    assert label_run.font.bold is True


def test_milestone_accepts_plain_list_and_centers_single_node():
    mapping = {
        "template_id": "milestone-test",
        "component_list": [
            {
                "location": "milestone.delivery",
                "semantic_description": "项目里程碑",
                "type": "Milestone",
                "data_source": {"name": "milestones"},
            }
        ],
    }
    payload = {"milestones": [{"label": "上线", "date": "07-20", "status": "done"}]}

    output = generate_report(_template_bytes(), mapping, payload, PostProcessingRegistry())
    doc = PptxDocument.open(output)
    index = doc.shape_index()

    node = index["milestone.delivery.item_1.node"].shape
    label = index["milestone.delivery.item_1.label"].shape
    assert label.text == "上线"
    assert abs((node.left + node.width / 2) - Inches(4.7)) < Inches(0.05)


def test_milestone_treats_numeric_dimensions_as_inches():
    mapping = {
        "template_id": "milestone-test",
        "component_list": [
            {
                "location": "milestone.delivery",
                "semantic_description": "项目里程碑",
                "type": "Milestone",
                "config": {"node_size": 0.25, "date_height": 0.3, "label_height": 0.3},
                "data_source": {"name": "milestones"},
            }
        ],
    }
    payload = {"milestones": [{"label": "上线", "date": "07-20", "status": "done"}]}

    output = generate_report(_template_bytes(), mapping, payload, PostProcessingRegistry())
    doc = PptxDocument.open(output)
    index = doc.shape_index()

    assert index["milestone.delivery.item_1.node"].shape.width == Inches(0.25)
    assert index["milestone.delivery.item_1.date"].shape.height == Inches(0.3)
    assert index["milestone.delivery.item_1.label"].shape.height == Inches(0.3)


def test_milestone_supports_horizontal_padding_date_color_and_centered_text():
    mapping = {
        "template_id": "milestone-test",
        "component_list": [
            {
                "location": "milestone.delivery",
                "semantic_description": "项目里程碑",
                "type": "Milestone",
                "config": {"horizontal_padding": 1, "date_color": "C00000"},
                "data_source": {"name": "milestones"},
            }
        ],
    }
    payload = {
        "milestones": [
            {"label": "准备", "date": "04-20", "status": "done"},
            {"label": "验收", "date": "07-05", "status": "pending"},
        ]
    }

    output = generate_report(_template_bytes(), mapping, payload, PostProcessingRegistry())
    doc = PptxDocument.open(output)
    index = doc.shape_index()
    first_node = index["milestone.delivery.item_1.node"].shape
    second_node = index["milestone.delivery.item_2.node"].shape
    first_date = index["milestone.delivery.item_1.date"].shape
    first_label = index["milestone.delivery.item_1.label"].shape
    second_date = index["milestone.delivery.item_2.date"].shape
    second_label = index["milestone.delivery.item_2.label"].shape
    date_run = first_date.text_frame.paragraphs[0].runs[0]
    first_node_center = first_node.left + first_node.width / 2
    second_node_center = second_node.left + second_node.width / 2

    assert abs((first_node.left + first_node.width / 2) - Inches(1.7)) < Inches(0.05)
    assert abs((second_node.left + second_node.width / 2) - Inches(7.7)) < Inches(0.05)
    assert first_date.left + first_date.width / 2 == first_node_center
    assert first_label.left + first_label.width / 2 == first_node_center
    assert second_date.left + second_date.width / 2 == second_node_center
    assert second_label.left + second_label.width / 2 == second_node_center
    assert date_run.font.color.rgb == RGBColor(0xC0, 0x00, 0x00)
