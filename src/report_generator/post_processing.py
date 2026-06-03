from __future__ import annotations

from collections.abc import Callable
from typing import Any


PostProcessor = Callable[..., Any]


class PostProcessingRegistry:
    def __init__(self) -> None:
        self._processors: dict[str, PostProcessor] = {}

    def register(self, name: str, processor: PostProcessor) -> None:
        self._processors[name] = processor

    def has(self, name: str) -> bool:
        return name in self._processors

    def call(self, name: str, **params: Any) -> Any:
        return self._processors[name](**params)
