from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt


def main() -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(8), Inches(0.8))
    title.name = "text.report_title"
    title.text_frame.text = "报告标题"
    title.text_frame.paragraphs[0].runs[0].font.size = Pt(28)

    table_shape = slide.shapes.add_table(2, 2, Inches(0.5), Inches(1.6), Inches(8), Inches(2.5))
    table_shape.name = "table.top_risks"
    table = table_shape.table
    table.cell(0, 0).text = "风险类型"
    table.cell(0, 1).text = "风险描述"
    table.cell(1, 0).text = "示例类型"
    table.cell(1, 1).text = "示例描述"

    output = Path(__file__).with_name("sample_template.pptx")
    prs.save(output)
    print(output)


if __name__ == "__main__":
    main()
