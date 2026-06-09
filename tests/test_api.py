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


def test_reports_pptx_endpoint_accepts_llm_concurrency(monkeypatch, simple_template_bytes):
    calls = []

    def fake_generate_report(template_bytes, mapping_json, payload_json, registry=None, processor=None, llm_concurrency=None):
        calls.append(llm_concurrency)
        return simple_template_bytes

    monkeypatch.setattr("report_generator.api.generate_report", fake_generate_report)
    client = TestClient(app)
    mapping = {"template_id": "project-monthly-report-ppt-v1", "component_list": []}
    payload = {}

    response = client.post(
        "/reports/pptx",
        data={"llmConcurrency": "7"},
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
    assert calls == [7]


def test_async_reports_pptx_endpoint_accepts_llm_concurrency(monkeypatch, simple_template_bytes):
    calls = []

    def fake_generate_report(template_bytes, mapping_json, payload_json, registry=None, processor=None, llm_concurrency=None):
        calls.append(llm_concurrency)
        return simple_template_bytes

    monkeypatch.setattr("report_generator.api.generate_report", fake_generate_report)
    client = TestClient(app)
    mapping = {"template_id": "project-monthly-report-ppt-v1", "component_list": []}
    payload = {}

    response = client.post(
        "/reports/pptx/tasks",
        data={"llmConcurrency": "3"},
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

    assert response.status_code == 202
    task_id = response.json()["taskId"]

    status = client.get("/reports/pptx/tasks", params={"taskId": task_id})
    assert status.status_code == 200
    assert status.json()["status"] == "succeeded"
    assert calls == [3]


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


def test_async_reports_pptx_endpoint_completes_and_downloads(simple_template_bytes):
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
        "/reports/pptx/tasks",
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

    assert response.status_code == 202
    assert "task_id" not in response.json()
    task_id = response.json()["taskId"]

    status = client.get("/reports/pptx/tasks", params={"taskId": task_id})
    assert status.status_code == 200
    assert status.json()["taskId"] == task_id
    assert status.json()["status"] in {"pending", "running", "succeeded"}

    for _ in range(20):
        status = client.get("/reports/pptx/tasks", params={"taskId": task_id})
        if status.json()["status"] == "succeeded":
            break
    assert status.json()["status"] == "succeeded"

    download = client.get("/reports/pptx/tasks/download", params={"taskId": task_id})
    assert download.status_code == 200
    assert download.content[:2] == b"PK"
