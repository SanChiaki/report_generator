from __future__ import annotations

from copy import deepcopy
from collections.abc import Mapping
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

    if component.config.get("preserve_style"):
        return _apply_table_preserving_style(shape, component, columns, rows)

    row_count = len(rows) + 1
    column_count = len(columns)
    if row_count == 0 or column_count == 0:
        raise ReportGenerationError(
            ErrorCode.DATA_SOURCE_INVALID,
            f"组件 {component.location} 的表格数据为空且无法推断列",
            component,
        )

    x, y, cx, cy = shape.left, shape.top, shape.width, shape.height
    source_table = shape.table
    slide = shape.part.slide
    new_shape = slide.shapes.add_table(row_count, column_count, x, y, cx, cy)
    new_shape.name = component.location
    table = new_shape.table
    _copy_table_style(source_table, table, cx, cy)
    doc.remove_shape(shape)

    for col_index, column in enumerate(columns):
        _replace_cell_text_preserving_style(table.cell(0, col_index), column["label"])
    for row_index, row in enumerate(rows, start=1):
        for col_index, column in enumerate(columns):
            _replace_cell_text_preserving_style(table.cell(row_index, col_index), row.get(column["key"], ""))

    if "font_size" in component.config:
        _apply_font_size(table, int(component.config["font_size"]))

    return new_shape


def _copy_table_style(source_table: Any, target_table: Any, width: int, height: int) -> None:
    _copy_dimensions(source_table, target_table, width, height)
    source_row_count = len(source_table.rows)
    source_col_count = len(source_table.columns)
    for row_index in range(len(target_table.rows)):
        source_row_index = _source_index(row_index, source_row_count, repeat_after_first=True)
        for col_index in range(len(target_table.columns)):
            source_col_index = _source_index(col_index, source_col_count, repeat_after_first=False)
            _copy_cell_style(
                source_table.cell(source_row_index, source_col_index),
                target_table.cell(row_index, col_index),
            )


def _copy_dimensions(source_table: Any, target_table: Any, width: int, height: int) -> None:
    column_widths = [
        source_table.columns[_source_index(index, len(source_table.columns), repeat_after_first=False)].width
        for index in range(len(target_table.columns))
    ]
    row_heights = [
        source_table.rows[_source_index(index, len(source_table.rows), repeat_after_first=True)].height
        for index in range(len(target_table.rows))
    ]
    _assign_scaled_sizes(target_table.columns, column_widths, width)
    _assign_scaled_sizes(target_table.rows, row_heights, height)


def _assign_scaled_sizes(collection: Any, source_sizes: list[int], total_size: int) -> None:
    source_total = sum(source_sizes)
    if source_total <= 0:
        even_size = int(total_size / len(collection))
        for item in collection:
            if hasattr(item, "width"):
                item.width = even_size
            else:
                item.height = even_size
        return
    remaining = total_size
    last_index = len(source_sizes) - 1
    for index, item in enumerate(collection):
        if index == last_index:
            size = remaining
        else:
            size = int(total_size * source_sizes[index] / source_total)
            remaining -= size
        if hasattr(item, "width"):
            item.width = size
        else:
            item.height = size


def _source_index(index: int, source_count: int, repeat_after_first: bool) -> int:
    if source_count <= 1:
        return 0
    if index < source_count:
        return index
    if repeat_after_first:
        return 1 + ((index - 1) % (source_count - 1))
    return source_count - 1


def _copy_cell_style(source_cell: Any, target_cell: Any) -> None:
    source_tc = source_cell._tc
    target_tc = target_cell._tc
    _replace_xml_child(target_tc, target_tc.txBody, deepcopy(source_tc.txBody))
    _replace_xml_child(target_tc, target_tc.tcPr, deepcopy(source_tc.get_or_add_tcPr()))


def _replace_xml_child(parent: Any, old_child: Any, new_child: Any) -> None:
    index = parent.index(old_child)
    parent.remove(old_child)
    parent.insert(index, new_child)


def _apply_table_preserving_style(
    shape: Any,
    component: ComponentMapping,
    columns: list[dict[str, str]],
    rows: list[dict[str, Any]],
) -> Any:
    table = shape.table
    row_count = len(rows) + 1
    column_count = len(columns)
    if row_count > len(table.rows):
        raise ReportGenerationError(
            ErrorCode.TABLE_OVERFLOW,
            f"组件 {component.location} 的数据行数 {len(rows)} 超过模板预留行数 {len(table.rows) - 1}",
            component,
        )
    if column_count > len(table.columns):
        raise ReportGenerationError(
            ErrorCode.TABLE_OVERFLOW,
            f"组件 {component.location} 的列数 {len(columns)} 超过模板预留列数 {len(table.columns)}",
            component,
        )

    for col_index, column in enumerate(columns):
        _replace_cell_text_preserving_style(table.cell(0, col_index), column["label"])
    for row_index, row in enumerate(rows, start=1):
        for col_index, column in enumerate(columns):
            _replace_cell_text_preserving_style(table.cell(row_index, col_index), row.get(column["key"], ""))

    for row_index in range(row_count, len(table.rows)):
        for col_index in range(len(table.columns)):
            _replace_cell_text_preserving_style(table.cell(row_index, col_index), "")
    for row_index in range(row_count):
        for col_index in range(column_count, len(table.columns)):
            _replace_cell_text_preserving_style(table.cell(row_index, col_index), "")

    return shape


def _normalize_table(value: Any, component: ComponentMapping) -> tuple[list[dict[str, str]], list[dict[str, Any]]]:
    if isinstance(value, dict) and "columns" in value and "rows" in value:
        if not isinstance(value["columns"], list) or not isinstance(value["rows"], list):
            raise _invalid_table_data(component)
        for column in value["columns"]:
            if not isinstance(column, Mapping):
                raise _invalid_table_data(component)
        for row in value["rows"]:
            if not isinstance(row, Mapping):
                raise _invalid_table_data(component)
        columns = [
            {"key": str(column.get("key")), "label": str(column.get("label", column.get("key")))}
            for column in value["columns"]
        ]
        rows = [dict(row) for row in value["rows"]]
        return columns, rows

    if isinstance(value, list):
        for row in value:
            if not isinstance(row, Mapping):
                raise _invalid_table_data(component)
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


def _invalid_table_data(component: ComponentMapping) -> ReportGenerationError:
    return ReportGenerationError(
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


def _apply_font_size(table: Any, font_size: int) -> None:
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)


def _replace_cell_text_preserving_style(cell: Any, value: Any) -> None:
    text = "" if value is None else str(value)
    text_frame = cell.text_frame
    paragraph = text_frame.paragraphs[0]
    if paragraph.runs:
        paragraph.runs[0].text = text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run().text = text
    for extra_paragraph in text_frame.paragraphs[1:]:
        for run in extra_paragraph.runs:
            run.text = ""
