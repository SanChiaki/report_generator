from __future__ import annotations

from collections.abc import Mapping
from datetime import datetime
from typing import Any

from report_generator.post_processing import PostProcessingRegistry


TASK_COLUMNS = [
    "阶段名称",
    "任务名称",
    "产品线",
    "任务状态",
    "计划完成时间",
    "实际完成时间",
    "责任人",
]

SEVERITY_ORDER = {"紧急": 0, "重要": 1, "一般": 2}


def default_registry() -> PostProcessingRegistry:
    registry = PostProcessingRegistry()
    registry.register("general_project_delivery_plan", general_project_delivery_plan)
    registry.register("general_project_milestones", general_project_milestones)
    registry.register("general_stage_task_details", general_project_delivery_plan)
    registry.register("general_top_risks_and_issues", general_top_risks_and_issues)
    registry.register("general_top_risks_and_issues_v2", general_top_risks_and_issues_v2)
    return registry


def general_project_delivery_plan(sr_api_data: Mapping[str, Any]) -> list[dict[str, Any]]:
    rows = _project_stages(sr_api_data, prefer_period=False)
    return [{column: _text(row.get(column)) for column in TASK_COLUMNS} for row in rows]


def general_project_milestones(sr_api_data: Mapping[str, Any]) -> dict[str, list[dict[str, str]]]:
    stages: dict[str, dict[str, Any]] = {}
    for row in _project_stages(sr_api_data, prefer_period=False):
        label = _text(row.get("阶段名称"))
        if not label:
            continue
        item = stages.setdefault(label, {"label": label, "date": "", "status_rank": 0})
        item["date"] = max(str(item["date"]), _stage_date(row))
        item["status_rank"] = max(int(item["status_rank"]), _stage_status_rank(row.get("任务状态")))
    return {
        "items": [
            {"label": str(item["label"]), "date": str(item["date"]), "status": _stage_status(item["status_rank"])}
            for item in stages.values()
            if str(item["date"])
        ]
    }


def general_top_risks_and_issues(sr_api_data: Mapping[str, Any]) -> list[dict[str, Any]]:
    rows = []
    for issue in _top_issue_rows(sr_api_data):
        rows.append(
            {
                "风险类型": issue["severity"],
                "风险描述": issue["description"],
                "责任人": issue["owner"],
                "问题状态": issue["status"],
                "计划解决时间": issue["due_date"],
                "解决措施与进展": issue["action"],
            }
        )
    return rows


def general_top_risks_and_issues_v2(sr_api_data: Mapping[str, Any], limit: int = 3) -> dict[str, list[dict[str, Any]]]:
    return {"items": _top_issue_rows(sr_api_data)[:limit]}


def _project_stages(sr_api_data: Mapping[str, Any], *, prefer_period: bool) -> list[Mapping[str, Any]]:
    period_rows = _nested_list(sr_api_data, "周期数据", "项目阶段")
    all_rows = _nested_list(sr_api_data, "全量数据", "项目阶段")
    rows = period_rows if prefer_period and period_rows else all_rows
    return [row for row in rows if isinstance(row, Mapping)]


def _top_issue_rows(sr_api_data: Mapping[str, Any]) -> list[dict[str, Any]]:
    issues = _nested_list(sr_api_data, "周期数据", "问题与风险") or _nested_list(sr_api_data, "全量数据", "问题与风险")
    rows = [_normalize_issue(issue) for issue in issues if isinstance(issue, Mapping)]
    rows.sort(key=_issue_sort_key)
    return rows


def _normalize_issue(issue: Mapping[str, Any]) -> dict[str, Any]:
    action = _text(issue.get("问题解决措施与进展")) or "待补充处理措施"
    return {
        "severity": _normalize_severity(issue.get("紧急程度")),
        "created_at": _text(issue.get("问题产生时间")),
        "description": _text(issue.get("问题及风险描述") or issue.get("风险描述") or issue.get("问题描述")),
        "action": action,
        "owner": _text(issue.get("责任人")) or "暂未分配",
        "status": _text(issue.get("问题状态")) or "处理中",
        "due_date": _date_only(issue.get("要求解决日期")),
    }


def _issue_sort_key(issue: Mapping[str, Any]) -> tuple[int, int, str, str]:
    severity = SEVERITY_ORDER.get(str(issue.get("severity", "一般")), SEVERITY_ORDER["一般"])
    status_rank = 1 if str(issue.get("status", "")).strip() in {"已关闭", "关闭", "已解决"} else 0
    due_date = str(issue.get("due_date", ""))
    created_at = str(issue.get("created_at", ""))
    return (status_rank, severity, due_date or "9999-12-31", created_at)


def _normalize_severity(value: Any) -> str:
    severity = _text(value)
    if severity in SEVERITY_ORDER:
        return severity
    return "一般"


def _stage_date(row: Mapping[str, Any]) -> str:
    return _date_only(row.get("实际完成时间")) or _date_only(row.get("计划完成时间"))


def _stage_status_rank(value: Any) -> int:
    status = _text(value)
    if status in {"已完成", "完成", "已结束", "结束"}:
        return 1
    if status in {"进行中", "处理中", "执行中"}:
        return 2
    return 0


def _stage_status(rank: Any) -> str:
    if rank == 2:
        return "active"
    if rank == 1:
        return "done"
    return "pending"


def _nested_list(data: Mapping[str, Any], section: str, key: str) -> list[Any]:
    section_value = data.get(section)
    if not isinstance(section_value, Mapping):
        return []
    value = section_value.get(key)
    return value if isinstance(value, list) else []


def _date_only(value: Any) -> str:
    text = _text(value)
    if not text:
        return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return text[:10]


def _text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()
