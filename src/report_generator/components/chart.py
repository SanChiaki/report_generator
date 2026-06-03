from __future__ import annotations

from numbers import Real
from typing import Any

from pptx.chart.data import CategoryChartData

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping


def apply_chart(shape: Any, component: ComponentMapping, value: Any) -> None:
    if not getattr(shape, "has_chart", False):
        raise ReportGenerationError(
            ErrorCode.TYPE_MISMATCH,
            f"组件 {component.location} 不是图表组件",
            component,
        )
    if not isinstance(value, dict):
        raise ReportGenerationError(
            ErrorCode.CHART_DATA_INVALID,
            f"组件 {component.location} 的图表数据必须是对象",
            component,
        )
    categories = value.get("categories", [])
    series = value.get("series", [])
    max_categories = int(component.config.get("max_categories", 24))
    max_series = int(component.config.get("max_series", 6))
    if len(categories) > max_categories or len(series) > max_series:
        raise ReportGenerationError(
            ErrorCode.CHART_DATA_INVALID,
            f"组件 {component.location} 的图表数据超过分类或系列数量限制",
            component,
        )

    chart_data = CategoryChartData()
    chart_data.categories = [str(category) for category in categories]
    for item in series:
        if not isinstance(item, dict):
            raise ReportGenerationError(
                ErrorCode.CHART_DATA_INVALID,
                f"组件 {component.location} 的图表系列必须是对象",
                component,
            )
        values = item.get("values", [])
        if len(values) != len(categories):
            raise ReportGenerationError(
                ErrorCode.CHART_DATA_INVALID,
                f"组件 {component.location} 的系列 {item.get('name')} 数据长度与分类数量不一致",
                component,
            )
        chart_data.add_series(str(item.get("name", "")), tuple(_numeric_value(component, value) for value in values))

    try:
        shape.chart.replace_data(chart_data)
    except Exception as exc:
        raise ReportGenerationError(
            ErrorCode.CHART_DATA_INVALID,
            f"组件 {component.location} 的图表数据替换失败: {exc}",
            component,
        ) from exc


def _numeric_value(component: ComponentMapping, value: Any) -> int | float:
    if isinstance(value, bool) or not isinstance(value, Real):
        raise ReportGenerationError(
            ErrorCode.CHART_DATA_INVALID,
            f"组件 {component.location} 的图表数值必须是数字",
            component,
        )
    return value
