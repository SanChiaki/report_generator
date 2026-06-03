from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from typing import Any

from report_generator.models import ComponentMapping


class ErrorCode(StrEnum):
    COMPONENT_NOT_FOUND = "COMPONENT_NOT_FOUND"
    DUPLICATE_COMPONENT_NAME = "DUPLICATE_COMPONENT_NAME"
    TYPE_MISMATCH = "TYPE_MISMATCH"
    DATA_SOURCE_NOT_FOUND = "DATA_SOURCE_NOT_FOUND"
    DATA_SOURCE_INVALID = "DATA_SOURCE_INVALID"
    POST_PROCESSING_FAILED = "POST_PROCESSING_FAILED"
    TEXT_OVERFLOW = "TEXT_OVERFLOW"
    TABLE_OVERFLOW = "TABLE_OVERFLOW"
    CHART_DATA_INVALID = "CHART_DATA_INVALID"
    IMAGE_LOAD_FAILED = "IMAGE_LOAD_FAILED"
    PPTX_PARSE_FAILED = "PPTX_PARSE_FAILED"
    PPTX_RENDER_FAILED = "PPTX_RENDER_FAILED"


@dataclass
class ReportGenerationError(Exception):
    error_code: ErrorCode
    message: str
    component: ComponentMapping | None = None
    details: dict[str, Any] | None = None

    def to_response(self) -> dict[str, Any]:
        response: dict[str, Any] = {
            "error_code": self.error_code.value,
            "message": self.message,
        }
        if self.component is not None:
            response["component"] = {
                "location": self.component.location,
                "type": self.component.type,
                "semantic_description": self.component.semantic_description,
            }
        if self.details:
            response["details"] = self.details
        return response
