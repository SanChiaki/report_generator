from report_generator.builtin_processors import (
    general_project_delivery_plan,
    general_project_milestones,
    general_top_risks_and_issues,
    general_top_risks_and_issues_v2,
)


def test_general_project_delivery_plan_returns_task_rows_from_full_data():
    payload = {
        "全量数据": {
            "项目阶段": [
                {
                    "阶段名称": "调测",
                    "任务名称": "配置检查",
                    "产品线": "交付",
                    "任务状态": "进行中",
                    "计划完成时间": "2026-06-14",
                    "实际完成时间": "",
                    "责任人": "张三",
                    "额外字段": "不输出",
                }
            ]
        }
    }

    rows = general_project_delivery_plan(payload)

    assert rows == [
        {
            "阶段名称": "调测",
            "任务名称": "配置检查",
            "产品线": "交付",
            "任务状态": "进行中",
            "计划完成时间": "2026-06-14",
            "实际完成时间": "",
            "责任人": "张三",
        }
    ]


def test_general_project_milestones_returns_full_date_items_by_stage():
    payload = {
        "全量数据": {
            "项目阶段": [
                {
                    "阶段名称": "设计",
                    "任务状态": "已完成",
                    "计划完成时间": "2026-06-26 08:00:00",
                    "实际完成时间": "2026-06-03 12:00:00",
                },
                {
                    "阶段名称": "设计",
                    "任务状态": "进行中",
                    "计划完成时间": "2026-06-28",
                    "实际完成时间": "",
                },
                {
                    "阶段名称": "验收",
                    "任务状态": "未开始",
                    "计划完成时间": "2026-07-09",
                    "实际完成时间": "",
                },
            ]
        }
    }

    data = general_project_milestones(payload)

    assert data == {
        "items": [
            {"label": "设计", "date": "2026-06-28", "status": "active"},
            {"label": "验收", "date": "2026-07-09", "status": "pending"},
        ]
    }


def test_general_project_milestones_skips_stages_without_dates():
    payload = {
        "全量数据": {
            "项目阶段": [
                {
                    "阶段名称": "维护",
                    "任务状态": "待分配",
                    "计划完成时间": "",
                    "实际完成时间": "",
                },
                {
                    "阶段名称": "验收",
                    "任务状态": "未开始",
                    "计划完成时间": "2026-07-09",
                    "实际完成时间": "",
                },
            ]
        }
    }

    data = general_project_milestones(payload)

    assert data == {
        "items": [
            {"label": "验收", "date": "2026-07-09", "status": "pending"},
        ]
    }


def test_general_top_risks_and_issues_v2_returns_top_issue_card_items():
    payload = {
        "周期数据": {
            "问题与风险": [
                {
                    "问题及风险描述": "一般风险",
                    "紧急程度": "一般",
                    "问题产生时间": "2026-06-03 16:44:00",
                    "要求解决日期": "2026-06-18 16:45:19",
                    "问题状态": "处理中",
                    "问题解决措施与进展": "",
                    "责任人": "",
                },
                {
                    "问题及风险描述": "紧急风险",
                    "紧急程度": "紧急",
                    "问题产生时间": "2026-06-04 16:42:00",
                    "要求解决日期": "2026-06-19 16:42:58",
                    "问题状态": "处理中",
                    "问题解决措施与进展": "已协调处理",
                    "责任人": "李四",
                },
            ]
        }
    }

    data = general_top_risks_and_issues_v2(payload)

    assert data == {
        "items": [
            {
                "severity": "紧急",
                "created_at": "2026-06-04 16:42:00",
                "description": "紧急风险",
                "action": "已协调处理",
                "owner": "李四",
                "status": "处理中",
                "due_date": "2026-06-19",
            },
            {
                "severity": "一般",
                "created_at": "2026-06-03 16:44:00",
                "description": "一般风险",
                "action": "待补充处理措施",
                "owner": "暂未分配",
                "status": "处理中",
                "due_date": "2026-06-18",
            },
        ]
    }


def test_general_top_risks_and_issues_returns_table_rows():
    payload = {
        "周期数据": {
            "问题与风险": [
                {
                    "问题及风险描述": "客户验收排期待确认",
                    "紧急程度": "重要",
                    "要求解决日期": "2026-06-18",
                    "问题状态": "",
                    "问题解决措施与进展": "",
                    "责任人": "",
                }
            ]
        }
    }

    rows = general_top_risks_and_issues(payload)

    assert rows == [
        {
            "风险类型": "重要",
            "风险描述": "客户验收排期待确认",
            "责任人": "暂未分配",
            "问题状态": "处理中",
            "计划解决时间": "2026-06-18",
            "解决措施与进展": "待补充处理措施",
        }
    ]
