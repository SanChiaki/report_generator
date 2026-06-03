from __future__ import annotations

from typing import Any

from pptx.util import Pt

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping
from report_generator.pptx.document import PptxDocument


def apply_table(
    doc: PptxDocument,
    shape: Any,
    component: ComponentMapping,
    value: Any,
) -> Any:
    if not getattr(shape, "has_table", False):
        raise ReportGenerationError(
            ErrorCode.TYPE_MISMATCH,
            f"组件 {component.location} 不是表格组件",
            component,
        )

    columns, rows = _normalize_table(value, component)
    max_rows = int(component.config.get("max_rows", 30))
    max_columns = int(component.config.get("max_columns", 10))
    if len(rows) > max_rows:
        raise ReportGenerationError(
            ErrorCode.TABLE_OVERFLOW,
            f"组件 {component.location} 的数据行数 {len(rows)} 超过限制 {max_rows}",
            component,
        )
    if len(columns) > max_columns:
        raise ReportGenerationError(
            ErrorCode.TABLE_OVERFLOW,
            f"组件 {component.location} 的列数 {len(columns)} 超过限制 {max_columns}",
            component,
        )

    row_count = len(rows) + 1
    column_count = len(columns)
    if row_count == 0 or column_count == 0:
        raise ReportGenerationError(
            ErrorCode.DATA_SOURCE_INVALID,
            f"组件 {component.location} 的表格数据为空且无法推断列",
            component,
        )

    x, y, cx, cy = shape.left, shape.top, shape.width, shape.height
    slide = shape.part.slide
    doc.remove_shape(shape)
    new_shape = slide.shapes.add_table(row_count, column_count, x, y, cx, cy)
    new_shape.name = component.location
    table = new_shape.table

    for index in range(column_count):
        table.columns[index].width = int(cx / column_count)
    for index in range(row_count):
        table.rows[index].height = int(cy / row_count)

    font_size = _table_font_size(component, row_count)
    for col_index, column in enumerate(columns):
        _set_cell_text(table.cell(0, col_index), column["label"], font_size, bold=True)
    for row_index, row in enumerate(rows, start=1):
        for col_index, column in enumerate(columns):
            _set_cell_text(table.cell(row_index, col_index), row.get(column["key"], ""), font_size)

    return new_shape


def _normalize_table(value: Any, component: ComponentMapping) -> tuple[list[dict[str, str]], list[dict[str, Any]]]:
    if isinstance(value, dict) and "columns" in value and "rows" in value:
        columns = [
            {"key": str(column.get("key")), "label": str(column.get("label", column.get("key")))}
            for column in value["columns"]
        ]
        rows = [dict(row) for row in value["rows"]]
        return columns, rows

    if isinstance(value, list):
        rows = [dict(row) for row in value]
        order = component.config.get("order")
        if order:
            keys = [str(key) for key in order]
        else:
            keys = list(rows[0].keys()) if rows else []
        columns = [{"key": key, "label": key} for key in keys]
        return columns, rows

    raise ReportGenerationError(
        ErrorCode.DATA_SOURCE_INVALID,
        f"组件 {component.location} 的表格数据必须是对象数组或 columns/rows 对象",
        component,
    )


def _table_font_size(component: ComponentMapping, row_count: int) -> int:
    minimum = int(component.config.get("min_font_size", 8))
    preferred = int(component.config.get("font_size", 12))
    if row_count <= 8:
        return preferred
    return max(minimum, preferred - (row_count - 8))


def _set_cell_text(cell: Any, value: Any, font_size: int, bold: bool = False) -> None:
    cell.text = "" if value is None else str(value)
    for paragraph in cell.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(font_size)
            run.font.bold = bold
