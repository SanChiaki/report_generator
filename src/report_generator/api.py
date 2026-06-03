from __future__ import annotations

import json
from io import BytesIO
from typing import Any

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import StreamingResponse
from pydantic import ValidationError

from report_generator.errors import ReportGenerationError
from report_generator.generator import generate_report

PPTX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

app = FastAPI(title="PPT Report Generator")


@app.post("/reports/pptx")
async def create_pptx_report(
    template: UploadFile = File(...),
    mapping: UploadFile = File(...),
    payload: UploadFile = File(...),
) -> StreamingResponse:
    try:
        template_bytes = await template.read()
        mapping_json = _loads_json(await mapping.read(), "mapping")
        payload_json = _loads_json(await payload.read(), "payload")
        output = generate_report(template_bytes, mapping_json, payload_json)
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
