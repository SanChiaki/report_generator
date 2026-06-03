import json

from fastapi.testclient import TestClient

from report_generator.api import app


def test_reports_pptx_endpoint_returns_pptx(simple_template_bytes):
    client = TestClient(app)
    mapping = {
        "template_id": "project-monthly-report-ppt-v1",
        "component_list": [
            {
                "location": "text.report_title",
                "semantic_description": "报告标题",
                "type": "Text",
                "data_source": {
                    "name": "api_data",
                    "template": "{{ 全量数据.项目概览.项目名称 }}-报告",
                },
            }
        ],
    }
    payload = {"api_data": {"全量数据": {"项目概览": {"项目名称": "智慧园区"}}}}

    response = client.post(
        "/reports/pptx",
        files={
            "template": (
                "template.pptx",
                simple_template_bytes,
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ),
            "mapping": (
                "mapping.json",
                json.dumps(mapping).encode("utf-8"),
                "application/json",
            ),
            "payload": (
                "payload.json",
                json.dumps(payload).encode("utf-8"),
                "application/json",
            ),
        },
    )

    assert response.status_code == 200
    assert response.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
    assert response.content[:2] == b"PK"


def test_reports_pptx_endpoint_returns_structured_error(simple_template_bytes):
    client = TestClient(app)
    mapping = {
        "template_id": "project-monthly-report-ppt-v1",
        "component_list": [
            {
                "location": "text.missing",
                "semantic_description": "缺失组件",
                "type": "Text",
                "data_source": {"name": "api_data"},
            }
        ],
    }

    response = client.post(
        "/reports/pptx",
        files={
            "template": (
                "template.pptx",
                simple_template_bytes,
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ),
            "mapping": (
                "mapping.json",
                json.dumps(mapping).encode("utf-8"),
                "application/json",
            ),
            "payload": (
                "payload.json",
                json.dumps({"api_data": "x"}).encode("utf-8"),
                "application/json",
            ),
        },
    )

    assert response.status_code == 400
    assert response.json()["detail"]["error_code"] == "COMPONENT_NOT_FOUND"
