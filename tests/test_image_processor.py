import base64
from io import BytesIO

from PIL import Image

from report_generator.components.image import apply_image
from report_generator.models import ComponentMapping
from report_generator.pptx.document import PptxDocument


def image_component(**config):
    return ComponentMapping.model_validate(
        {
            "location": "image.company_logo",
            "semantic_description": "公司 Logo",
            "type": "Image",
            "config": config,
        }
    )


def png_data_uri() -> str:
    image = Image.new("RGB", (10, 10), color=(255, 0, 0))
    output = BytesIO()
    image.save(output, format="PNG")
    encoded = base64.b64encode(output.getvalue()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def test_apply_image_replaces_region_with_picture(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["image.company_logo"].shape

    new_shape = apply_image(doc, shape, image_component(fit="contain"), png_data_uri())

    assert new_shape.name == "image.company_logo"
    assert "PICTURE" in str(new_shape.shape_type)
