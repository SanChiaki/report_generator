from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field


ComponentType = Literal[
    "Text",
    "Image",
    "Table",
    "Chart",
    "Shape",
    "Milestone",
    "GanttChart",
]


class DataSource(BaseModel):
    model_config = ConfigDict(extra="allow")

    name: str | None = None
    index: str | None = None
    template: str | None = None
    params: dict[str, Any] = Field(default_factory=dict)
    needs_post_processing: bool = False


class ComponentMapping(BaseModel):
    model_config = ConfigDict(extra="allow")

    location: str
    semantic_description: str | None = None
    type: ComponentType
    prompt: str | None = None
    data_example: Any | None = None
    config: dict[str, Any] = Field(default_factory=dict)
    data_source: DataSource | None = None
    visible: bool | None = None


class ReportMapping(BaseModel):
    model_config = ConfigDict(extra="allow")

    template_id: str
    component_list: list[ComponentMapping]
