from __future__ import annotations

import json
from io import BytesIO
from dataclasses import dataclass
from threading import Lock
from typing import Any
from uuid import uuid4

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.responses import StreamingResponse
from pydantic import ValidationError

from report_generator.errors import ReportGenerationError
from report_generator.generator import generate_report

PPTX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

app = FastAPI(title="PPT Report Generator")


@dataclass
class ReportTask:
    status: str
    output: bytes | None = None
    error: dict[str, Any] | None = None


_tasks: dict[str, ReportTask] = {}
_tasks_lock = Lock()


@app.post("/reports/pptx")
async def create_pptx_report(
    template: UploadFile = File(...),
    mapping: UploadFile = File(...),
    payload: UploadFile = File(...),
    llm_concurrency: int | None = Form(default=None, alias="llmConcurrency"),
) -> StreamingResponse:
    try:
        template_bytes = await template.read()
        mapping_json = _loads_json(await mapping.read(), "mapping")
        payload_json = _loads_json(await payload.read(), "payload")
        output = generate_report(template_bytes, mapping_json, payload_json, llm_concurrency=llm_concurrency)
    except ReportGenerationError as exc:
        raise HTTPException(status_code=400, detail=exc.to_response()) from exc
    except ValidationError as exc:
        raise HTTPException(
            status_code=400,
            detail={
                "error_code": "DATA_SOURCE_INVALID",
                "message": "mapping JSON 校验失败",
                "details": {"errors": exc.errors()},
            },
        ) from exc

    return StreamingResponse(
        BytesIO(output),
        media_type=PPTX_MEDIA_TYPE,
        headers={"Content-Disposition": 'attachment; filename="report.pptx"'},
    )


@app.post("/reports/pptx/tasks", status_code=202)
async def create_pptx_report_task(
    background_tasks: BackgroundTasks,
    template: UploadFile = File(...),
    mapping: UploadFile = File(...),
    payload: UploadFile = File(...),
    llm_concurrency: int | None = Form(default=None, alias="llmConcurrency"),
) -> dict[str, str]:
    template_bytes = await template.read()
    mapping_json = _loads_json(await mapping.read(), "mapping")
    payload_json = _loads_json(await payload.read(), "payload")
    task_id = uuid4().hex
    with _tasks_lock:
        _tasks[task_id] = ReportTask(status="pending")
    background_tasks.add_task(
        _run_report_task,
        task_id,
        template_bytes,
        mapping_json,
        payload_json,
        llm_concurrency,
    )
    return {"taskId": task_id, "status": "pending"}


@app.get("/reports/pptx/tasks")
async def get_pptx_report_task(task_id: str = Query(..., alias="taskId")) -> dict[str, Any]:
    task = _get_task(task_id)
    response: dict[str, Any] = {"taskId": task_id, "status": task.status}
    if task.error is not None:
        response["error"] = task.error
    return response


@app.get("/reports/pptx/tasks/download")
async def download_pptx_report_task(
    task_id: str = Query(..., alias="taskId"),
) -> StreamingResponse:
    task = _get_task(task_id)
    if task.status != "succeeded" or task.output is None:
        raise HTTPException(
            status_code=409,
            detail={
                "error_code": "REPORT_NOT_READY",
                "message": f"任务 {task_id} 尚未生成完成",
            },
        )
    return StreamingResponse(
        BytesIO(task.output),
        media_type=PPTX_MEDIA_TYPE,
        headers={"Content-Disposition": 'attachment; filename="report.pptx"'},
    )


def _run_report_task(
    task_id: str,
    template_bytes: bytes,
    mapping_json: dict[str, Any],
    payload_json: dict[str, Any],
    llm_concurrency: int | None,
) -> None:
    _update_task(task_id, status="running")
    try:
        output = generate_report(
            template_bytes,
            mapping_json,
            payload_json,
            llm_concurrency=llm_concurrency,
        )
    except ReportGenerationError as exc:
        _update_task(task_id, status="failed", error=exc.to_response())
    except ValidationError as exc:
        _update_task(
            task_id,
            status="failed",
            error={
                "error_code": "DATA_SOURCE_INVALID",
                "message": "mapping JSON 校验失败",
                "details": {"errors": exc.errors()},
            },
        )
    except Exception as exc:
        _update_task(
            task_id,
            status="failed",
            error={
                "error_code": "PPTX_RENDER_FAILED",
                "message": f"报告生成任务失败: {exc}",
            },
        )
    else:
        _update_task(task_id, status="succeeded", output=output)


def _get_task(task_id: str) -> ReportTask:
    with _tasks_lock:
        task = _tasks.get(task_id)
    if task is None:
        raise HTTPException(
            status_code=404,
            detail={
                "error_code": "TASK_NOT_FOUND",
                "message": f"任务 {task_id} 不存在",
            },
        )
    return task


def _update_task(
    task_id: str,
    *,
    status: str,
    output: bytes | None = None,
    error: dict[str, Any] | None = None,
) -> None:
    with _tasks_lock:
        task = _tasks[task_id]
        task.status = status
        if output is not None:
            task.output = output
        if error is not None:
            task.error = error


def _loads_json(raw: bytes, label: str) -> dict[str, Any]:
    try:
        value = json.loads(raw.decode("utf-8"))
    except Exception as exc:
        raise HTTPException(
            status_code=400,
            detail={
                "error_code": "DATA_SOURCE_INVALID",
                "message": f"{label} 不是合法 JSON",
            },
        ) from exc
    if not isinstance(value, dict):
        raise HTTPException(
            status_code=400,
            detail={
                "error_code": "DATA_SOURCE_INVALID",
                "message": f"{label} 必须是 JSON 对象",
            },
        )
    return value
