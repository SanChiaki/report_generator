from __future__ import annotations

import json
import os
from typing import Any, Protocol

import httpx

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping


class ComponentDataProcessor(Protocol):
    def process(self, component: ComponentMapping, value: Any) -> Any:
        ...


class OpenAICompletionProcessor:
    def __init__(
        self,
        *,
        api_key: str | None = None,
        base_url: str | None = None,
        model: str | None = None,
        completion_mode: str | None = None,
        timeout: float = 60.0,
    ) -> None:
        dotenv = _load_dotenv()
        self.api_key = api_key or _env("API_KEY", dotenv) or _env("OPENAI_API_KEY", dotenv)
        self.base_url = (base_url or _env("BASE_URL", dotenv) or _env("OPENAI_BASE_URL", dotenv) or "").rstrip("/")
        self.model = model or _env("MODEL", dotenv) or _env("OPENAI_MODEL", dotenv) or "gpt-4o-mini"
        self.completion_mode = (
            completion_mode
            or _env("COMPLETION_MODE", dotenv)
            or _env("OPENAI_COMPLETION_MODE", dotenv)
            or "chat"
        )
        self.timeout = timeout

    def process(self, component: ComponentMapping, value: Any) -> Any:
        if not self.api_key or not self.base_url:
            raise ReportGenerationError(
                ErrorCode.POST_PROCESSING_FAILED,
                "大模型后处理缺少 API_KEY 或 BASE_URL 配置",
                component,
            )

        response_text = self._complete(component, value)
        return _parse_processor_response(component, response_text)

    def _complete(self, component: ComponentMapping, value: Any) -> str:
        if self.completion_mode in {"completion", "completions", "legacy"}:
            return self._legacy_complete(component, value)
        return self._chat_complete(component, value)

    def _chat_complete(self, component: ComponentMapping, value: Any) -> str:
        url = f"{self.base_url}/chat/completions"
        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "system",
                    "content": (
                        "你是 PPT 报告组件数据处理器。根据组件说明、提示词、示例和原始数据，"
                        "生成可以直接传给该组件渲染器的数据。只返回 JSON，格式必须是 {\"data\": ...}。"
                    ),
                },
                {
                    "role": "user",
                    "content": json.dumps(
                        {
                            "component": {
                                "location": component.location,
                                "type": component.type,
                                "semantic_description": component.semantic_description,
                                "prompt": component.prompt,
                                "data_example": component.data_example,
                                "config": component.config,
                            },
                            "input_data": value,
                        },
                        ensure_ascii=False,
                    ),
                },
            ],
            "temperature": 0,
        }
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        try:
            response = httpx.post(url, json=payload, headers=headers, timeout=self.timeout)
            response.raise_for_status()
            body = response.json()
            return str(body["choices"][0]["message"]["content"])
        except ReportGenerationError:
            raise
        except Exception as exc:
            raise ReportGenerationError(
                ErrorCode.POST_PROCESSING_FAILED,
                f"大模型后处理请求失败: {exc}",
                component,
            ) from exc

    def _legacy_complete(self, component: ComponentMapping, value: Any) -> str:
        url = f"{self.base_url}/completions"
        payload = {
            "model": self.model,
            "prompt": (
                "你是 PPT 报告组件数据处理器。根据组件说明、提示词、示例和原始数据，"
                "生成可以直接传给该组件渲染器的数据。只返回 JSON，格式必须是 {\"data\": ...}。\n\n"
                + json.dumps(
                    {
                        "component": {
                            "location": component.location,
                            "type": component.type,
                            "semantic_description": component.semantic_description,
                            "prompt": component.prompt,
                            "data_example": component.data_example,
                            "config": component.config,
                        },
                        "input_data": value,
                    },
                    ensure_ascii=False,
                )
            ),
            "temperature": 0,
        }
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        try:
            response = httpx.post(url, json=payload, headers=headers, timeout=self.timeout)
            response.raise_for_status()
            body = response.json()
            return str(body["choices"][0]["text"])
        except ReportGenerationError:
            raise
        except Exception as exc:
            raise ReportGenerationError(
                ErrorCode.POST_PROCESSING_FAILED,
                f"大模型后处理请求失败: {exc}",
                component,
            ) from exc


def _parse_processor_response(component: ComponentMapping, response_text: str) -> Any:
    try:
        parsed = json.loads(_strip_code_fence(response_text))
    except Exception as exc:
        raise ReportGenerationError(
            ErrorCode.POST_PROCESSING_FAILED,
            f"大模型后处理返回内容不是合法 JSON: {response_text[:200]}",
            component,
        ) from exc
    if not isinstance(parsed, dict) or "data" not in parsed:
        raise ReportGenerationError(
            ErrorCode.POST_PROCESSING_FAILED,
            "大模型后处理返回 JSON 必须包含 data 字段",
            component,
        )
    return parsed["data"]


def _strip_code_fence(text: str) -> str:
    stripped = text.strip()
    if not stripped.startswith("```"):
        return stripped
    lines = stripped.splitlines()
    if lines and lines[0].startswith("```"):
        lines = lines[1:]
    if lines and lines[-1].strip() == "```":
        lines = lines[:-1]
    return "\n".join(lines).strip()


def _env(name: str, dotenv: dict[str, str]) -> str | None:
    return os.getenv(name) or dotenv.get(name)


def _load_dotenv(path: str = ".env") -> dict[str, str]:
    values: dict[str, str] = {}
    try:
        with open(path, encoding="utf-8") as handle:
            for line in handle:
                stripped = line.strip()
                if not stripped or stripped.startswith("#") or "=" not in stripped:
                    continue
                key, value = stripped.split("=", 1)
                key = key.strip()
                value = value.strip().strip('"').strip("'")
                if key:
                    values[key] = value
    except FileNotFoundError:
        return {}
    return values
