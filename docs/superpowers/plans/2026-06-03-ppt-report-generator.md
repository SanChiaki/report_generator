# PPT Report Generator Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a FastAPI service that accepts a PPTX template, an Excel-compatible component mapping JSON, and a business payload JSON, then returns a generated PPTX report.

**Architecture:** Use a layered Python package. FastAPI handles upload and response formatting; mapping models validate the Excel-compatible JSON; a data source resolver extracts or renders component values; PPTX utilities scan named shapes; component processors mutate text, images, tables, charts, and simple shape visibility; a generator orchestrates the full flow.

**Tech Stack:** Python 3.11+, FastAPI, Pydantic v2, python-pptx, jsonpath-ng, Jinja2, Pillow, httpx, pytest.

---

## Scope Check

The approved spec is one coherent subsystem: synchronous PPTX report generation over HTTP. Template registration, visual editing, thumbnail preview, and async jobs are outside the first implementation. This plan implements the MVP from the spec and leaves those extensions out of code.

## File Structure

Create this structure:

```text
pyproject.toml
README.md
src/report_generator/__init__.py
src/report_generator/api.py
src/report_generator/datasource.py
src/report_generator/errors.py
src/report_generator/generator.py
src/report_generator/models.py
src/report_generator/post_processing.py
src/report_generator/pptx/__init__.py
src/report_generator/pptx/document.py
src/report_generator/components/__init__.py
src/report_generator/components/chart.py
src/report_generator/components/image.py
src/report_generator/components/shape.py
src/report_generator/components/table.py
src/report_generator/components/text.py
tests/conftest.py
tests/test_api.py
tests/test_chart_processor.py
tests/test_datasource.py
tests/test_generator.py
tests/test_image_processor.py
tests/test_mapping_models.py
tests/test_pptx_document.py
tests/test_table_processor.py
tests/test_text_processor.py
examples/build_sample_template.py
examples/mapping.json
examples/payload.json
```

Responsibilities:

- `models.py`: Pydantic schema for the Excel-compatible mapping JSON.
- `errors.py`: Structured error codes and response serialization.
- `datasource.py`: JSONPath, Jinja template rendering, and post-processing dispatch.
- `post_processing.py`: Local registry for named data functions used by mappings.
- `pptx/document.py`: PPTX opening, named shape scanning, duplicate detection, shape removal, and output serialization.
- `components/*.py`: Type-specific component mutation.
- `generator.py`: End-to-end generation orchestration.
- `api.py`: FastAPI app and `/reports/pptx` endpoint.
- `tests/conftest.py`: In-memory PPTX template builders used by tests.
- `examples/*`: Small runnable sample template builder, mapping, and payload.

---

### Task 1: Project Scaffold

**Files:**
- Create: `pyproject.toml`
- Create: `README.md`
- Create: `src/report_generator/__init__.py`
- Test: `tests/test_mapping_models.py`

- [ ] **Step 1: Create `pyproject.toml`**

```toml
[build-system]
requires = ["setuptools>=69", "wheel"]
build-backend = "setuptools.build_meta"

[project]
name = "report-generator"
version = "0.1.0"
description = "HTTP PPTX report generator with Excel-compatible component mapping"
requires-python = ">=3.11"
dependencies = [
  "fastapi",
  "httpx",
  "jinja2",
  "jsonpath-ng",
  "pillow",
  "pydantic>=2",
  "python-multipart",
  "python-pptx",
  "uvicorn[standard]"
]

[project.optional-dependencies]
dev = [
  "pytest",
  "pytest-cov"
]

[tool.setuptools.packages.find]
where = ["src"]

[tool.pytest.ini_options]
testpaths = ["tests"]
pythonpath = ["src"]
addopts = "-q"
```

- [ ] **Step 2: Create `README.md`**

```markdown
# Report Generator

HTTP service for generating PPTX reports from:

- a PowerPoint template,
- an Excel-compatible component mapping JSON,
- a business payload JSON.

The PPT mapping keeps `template_id` and `component_list`. In PPT mappings, `location` is the PowerPoint selection pane shape name.

## Development

```bash
python -m pip install -e ".[dev]"
python -m pytest
```

## Run

```bash
uvicorn report_generator.api:app --reload
```
```

- [ ] **Step 3: Create `src/report_generator/__init__.py`**

```python
"""PPTX report generator package."""

__all__ = ["__version__"]

__version__ = "0.1.0"
```

- [ ] **Step 4: Add a temporary import smoke test**

Create `tests/test_mapping_models.py` with:

```python
from report_generator import __version__


def test_package_imports():
    assert __version__ == "0.1.0"
```

- [ ] **Step 5: Run test to verify the scaffold works**

Run:

```bash
python -m pip install -e ".[dev]"
python -m pytest tests/test_mapping_models.py -q
```

Expected: one passing test.

- [ ] **Step 6: Commit**

```bash
git add pyproject.toml README.md src/report_generator/__init__.py tests/test_mapping_models.py
git commit -m "chore: scaffold report generator package"
```

---

### Task 2: Mapping Models and Structured Errors

**Files:**
- Create: `src/report_generator/models.py`
- Create: `src/report_generator/errors.py`
- Modify: `tests/test_mapping_models.py`

- [ ] **Step 1: Replace `tests/test_mapping_models.py` with model tests**

```python
import pytest
from pydantic import ValidationError

from report_generator.models import ComponentMapping, DataSource, ReportMapping


def test_report_mapping_accepts_excel_compatible_shape():
    mapping = ReportMapping.model_validate(
        {
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
                    "config": {"min_font_size": 12},
                }
            ],
        }
    )

    assert mapping.template_id == "project-monthly-report-ppt-v1"
    assert mapping.component_list[0].location == "text.report_title"
    assert mapping.component_list[0].type == "Text"
    assert mapping.component_list[0].data_source.template.startswith("{{")


def test_component_requires_known_type():
    with pytest.raises(ValidationError):
        ComponentMapping.model_validate(
            {
                "location": "unknown.component",
                "semantic_description": "未知组件",
                "type": "Unknown",
            }
        )


def test_data_source_keeps_params_and_post_processing_flag():
    source = DataSource.model_validate(
        {
            "name": "general_top_risks_and_issues",
            "params": {"sr_api_data": "sr_api_data"},
            "needs_post_processing": True,
        }
    )

    assert source.name == "general_top_risks_and_issues"
    assert source.params == {"sr_api_data": "sr_api_data"}
    assert source.needs_post_processing is True
```

- [ ] **Step 2: Run model tests and verify they fail**

Run:

```bash
python -m pytest tests/test_mapping_models.py -q
```

Expected: import failure for `report_generator.models`.

- [ ] **Step 3: Create `src/report_generator/models.py`**

```python
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
```

- [ ] **Step 4: Create `src/report_generator/errors.py`**

```python
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
```

- [ ] **Step 5: Run model tests**

Run:

```bash
python -m pytest tests/test_mapping_models.py -q
```

Expected: all tests pass.

- [ ] **Step 6: Commit**

```bash
git add src/report_generator/models.py src/report_generator/errors.py tests/test_mapping_models.py
git commit -m "feat: add mapping models and structured errors"
```

---

### Task 3: Data Source Resolution

**Files:**
- Create: `src/report_generator/datasource.py`
- Create: `src/report_generator/post_processing.py`
- Create: `tests/test_datasource.py`

- [ ] **Step 1: Write failing data source tests**

Create `tests/test_datasource.py`:

```python
import pytest

from report_generator.datasource import resolve_component_value
from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping
from report_generator.post_processing import PostProcessingRegistry


def component(data_source):
    return ComponentMapping.model_validate(
        {
            "location": "text.report_title",
            "semantic_description": "报告标题",
            "type": "Text",
            "data_source": data_source,
        }
    )


def test_resolves_jsonpath_against_named_source():
    payload = {
        "api_data": {
            "全量数据": {
                "项目概览": {
                    "项目名称": "智慧园区",
                }
            }
        }
    }
    result = resolve_component_value(
        component(
            {
                "name": "api_data",
                "index": "$['全量数据']['项目概览']['项目名称']",
            }
        ),
        payload,
        PostProcessingRegistry(),
    )

    assert result == "智慧园区"


def test_renders_template_against_named_source():
    payload = {
        "api_data": {
            "全量数据": {
                "项目概览": {
                    "项目名称": "智慧园区",
                }
            }
        }
    }
    result = resolve_component_value(
        component(
            {
                "name": "api_data",
                "template": "{{ 全量数据.项目概览.项目名称 }}-报告",
            }
        ),
        payload,
        PostProcessingRegistry(),
    )

    assert result == "智慧园区-报告"


def test_calls_named_post_processing_function_with_params():
    registry = PostProcessingRegistry()
    registry.register("join_names", lambda users: ", ".join(item["name"] for item in users))
    payload = {"users": [{"name": "张三"}, {"name": "李四"}]}

    result = resolve_component_value(
        component(
            {
                "name": "join_names",
                "params": {"users": "users"},
                "needs_post_processing": True,
            }
        ),
        payload,
        registry,
    )

    assert result == "张三, 李四"


def test_missing_source_raises_structured_error():
    with pytest.raises(ReportGenerationError) as exc:
        resolve_component_value(
            component({"name": "api_data", "index": "$.missing"}),
            {},
            PostProcessingRegistry(),
        )

    assert exc.value.error_code == ErrorCode.DATA_SOURCE_NOT_FOUND
```

- [ ] **Step 2: Run tests and verify they fail**

Run:

```bash
python -m pytest tests/test_datasource.py -q
```

Expected: import failure for `report_generator.datasource`.

- [ ] **Step 3: Create `src/report_generator/post_processing.py`**

```python
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
```

- [ ] **Step 4: Create `src/report_generator/datasource.py`**

```python
from __future__ import annotations

from typing import Any

from jinja2 import StrictUndefined
from jinja2.sandbox import SandboxedEnvironment
from jsonpath_ng import parse

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping, DataSource
from report_generator.post_processing import PostProcessingRegistry


def resolve_component_value(
    component: ComponentMapping,
    payload: dict[str, Any],
    registry: PostProcessingRegistry,
) -> Any:
    source = component.data_source
    if source is None:
        return None

    if source.needs_post_processing:
        return _resolve_post_processed(component, source, payload, registry)

    if source.template:
        base = _source_base(component, source, payload)
        return _render_template(source.template, base)

    if source.index:
        base = _source_base(component, source, payload)
        return _resolve_jsonpath(component, source.index, base)

    if source.name and source.name in payload:
        return payload[source.name]

    raise ReportGenerationError(
        ErrorCode.DATA_SOURCE_INVALID,
        f"组件 {component.location} 的 data_source 缺少 index、template 或可用的 name",
        component,
    )


def _source_base(component: ComponentMapping, source: DataSource, payload: dict[str, Any]) -> Any:
    if source.name is None:
        return payload
    if source.name not in payload:
        raise ReportGenerationError(
            ErrorCode.DATA_SOURCE_NOT_FOUND,
            f"组件 {component.location} 引用了不存在的数据源 {source.name}",
            component,
        )
    return payload[source.name]


def _resolve_jsonpath(component: ComponentMapping, expression: str, base: Any) -> Any:
    matches = [match.value for match in parse(expression).find(base)]
    if not matches:
        raise ReportGenerationError(
            ErrorCode.DATA_SOURCE_NOT_FOUND,
            f"组件 {component.location} 的 JSONPath 没有匹配到数据: {expression}",
            component,
        )
    if len(matches) == 1:
        return matches[0]
    return matches


def _render_template(template: str, base: Any) -> str:
    env = SandboxedEnvironment(undefined=StrictUndefined, autoescape=False)
    return env.from_string(template).render(base)


def _resolve_post_processed(
    component: ComponentMapping,
    source: DataSource,
    payload: dict[str, Any],
    registry: PostProcessingRegistry,
) -> Any:
    if not source.name or not registry.has(source.name):
        raise ReportGenerationError(
            ErrorCode.POST_PROCESSING_FAILED,
            f"组件 {component.location} 引用了未注册的后处理函数 {source.name}",
            component,
        )
    params: dict[str, Any] = {}
    for param_name, payload_key in source.params.items():
        if payload_key not in payload:
            raise ReportGenerationError(
                ErrorCode.DATA_SOURCE_NOT_FOUND,
                f"组件 {component.location} 的参数 {param_name} 引用了不存在的数据 {payload_key}",
                component,
            )
        params[param_name] = payload[payload_key]
    try:
        return registry.call(source.name, **params)
    except Exception as exc:
        raise ReportGenerationError(
            ErrorCode.POST_PROCESSING_FAILED,
            f"组件 {component.location} 的后处理失败: {exc}",
            component,
        ) from exc
```

- [ ] **Step 5: Run data source tests**

Run:

```bash
python -m pytest tests/test_datasource.py -q
```

Expected: all tests pass.

- [ ] **Step 6: Commit**

```bash
git add src/report_generator/datasource.py src/report_generator/post_processing.py tests/test_datasource.py
git commit -m "feat: resolve Excel-compatible data sources"
```

---

### Task 4: PPTX Document Utilities

**Files:**
- Create: `src/report_generator/pptx/__init__.py`
- Create: `src/report_generator/pptx/document.py`
- Create: `tests/conftest.py`
- Create: `tests/test_pptx_document.py`

- [ ] **Step 1: Write failing PPTX document tests**

Create `tests/test_pptx_document.py`:

```python
import pytest

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.pptx.document import PptxDocument


def test_scans_named_shapes(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    index = doc.shape_index()

    assert "text.report_title" in index
    assert "table.top_risks" in index
    assert "chart.revenue_trend" in index


def test_duplicate_shape_names_raise_error(duplicate_name_template_bytes):
    doc = PptxDocument.open(duplicate_name_template_bytes)

    with pytest.raises(ReportGenerationError) as exc:
        doc.shape_index()

    assert exc.value.error_code == ErrorCode.DUPLICATE_COMPONENT_NAME
```

- [ ] **Step 2: Create `tests/conftest.py`**

```python
from __future__ import annotations

from io import BytesIO

import pytest
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt


def _save(prs: Presentation) -> bytes:
    output = BytesIO()
    prs.save(output)
    return output.getvalue()


@pytest.fixture
def simple_template_bytes() -> bytes:
    prs = Presentation()
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)

    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(6), Inches(0.6))
    title.name = "text.report_title"
    title.text_frame.text = "旧标题"
    title.text_frame.paragraphs[0].runs[0].font.size = Pt(24)

    summary = slide.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(4), Inches(0.7))
    summary.name = "text.summary"
    summary.text_frame.text = "旧摘要"
    summary.text_frame.paragraphs[0].runs[0].font.size = Pt(18)

    table_shape = slide.shapes.add_table(2, 2, Inches(0.5), Inches(2), Inches(5), Inches(1.2))
    table_shape.name = "table.top_risks"
    table = table_shape.table
    table.cell(0, 0).text = "风险类型"
    table.cell(0, 1).text = "风险描述"
    table.cell(1, 0).text = "旧类型"
    table.cell(1, 1).text = "旧描述"

    chart_data = CategoryChartData()
    chart_data.categories = ["Q1", "Q2"]
    chart_data.add_series("收入", (10, 20))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(6),
        Inches(2),
        Inches(3),
        Inches(2),
        chart_data,
    )
    chart.name = "chart.revenue_trend"

    badge = slide.shapes.add_shape(1, Inches(6), Inches(0.5), Inches(1.2), Inches(0.5))
    badge.name = "shape.status_badge"
    badge.text = "旧状态"

    return _save(prs)


@pytest.fixture
def duplicate_name_template_bytes() -> bytes:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    first = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(2), Inches(0.5))
    second = slide.shapes.add_textbox(Inches(0.5), Inches(1), Inches(2), Inches(0.5))
    first.name = "text.duplicate"
    second.name = "text.duplicate"
    return _save(prs)
```

- [ ] **Step 3: Run tests and verify they fail**

Run:

```bash
python -m pytest tests/test_pptx_document.py -q
```

Expected: import failure for `report_generator.pptx.document`.

- [ ] **Step 4: Create `src/report_generator/pptx/__init__.py`**

```python
"""PPTX helpers."""
```

- [ ] **Step 5: Create `src/report_generator/pptx/document.py`**

```python
from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from typing import Any

from pptx import Presentation

from report_generator.errors import ErrorCode, ReportGenerationError


@dataclass(frozen=True)
class ShapeRef:
    name: str
    slide_index: int
    shape: Any
    kind: str


class PptxDocument:
    def __init__(self, presentation: Presentation) -> None:
        self.presentation = presentation

    @classmethod
    def open(cls, template_bytes: bytes) -> "PptxDocument":
        try:
            return cls(Presentation(BytesIO(template_bytes)))
        except Exception as exc:
            raise ReportGenerationError(
                ErrorCode.PPTX_PARSE_FAILED,
                f"PPTX 模板解析失败: {exc}",
            ) from exc

    def shape_index(self) -> dict[str, ShapeRef]:
        index: dict[str, ShapeRef] = {}
        duplicates: list[str] = []
        for slide_index, slide in enumerate(self.presentation.slides):
            for shape in slide.shapes:
                name = getattr(shape, "name", "")
                if not name:
                    continue
                if name in index:
                    duplicates.append(name)
                    continue
                index[name] = ShapeRef(
                    name=name,
                    slide_index=slide_index,
                    shape=shape,
                    kind=infer_shape_kind(shape),
                )
        if duplicates:
            raise ReportGenerationError(
                ErrorCode.DUPLICATE_COMPONENT_NAME,
                "PPT 模板中存在重复组件名称",
                details={"duplicates": sorted(set(duplicates))},
            )
        return index

    def remove_shape(self, shape: Any) -> None:
        element = shape._element
        element.getparent().remove(element)

    def to_bytes(self) -> bytes:
        output = BytesIO()
        self.presentation.save(output)
        return output.getvalue()


def infer_shape_kind(shape: Any) -> str:
    if getattr(shape, "has_table", False):
        return "Table"
    if getattr(shape, "has_chart", False):
        return "Chart"
    if getattr(shape, "has_text_frame", False):
        return "Text"
    if getattr(shape, "shape_type", None) is not None:
        return "Shape"
    return "Unknown"
```

- [ ] **Step 6: Run PPTX document tests**

Run:

```bash
python -m pytest tests/test_pptx_document.py -q
```

Expected: all tests pass.

- [ ] **Step 7: Commit**

```bash
git add src/report_generator/pptx tests/conftest.py tests/test_pptx_document.py
git commit -m "feat: scan named PPTX shapes"
```

---

### Task 5: Text Processor

**Files:**
- Create: `src/report_generator/components/__init__.py`
- Create: `src/report_generator/components/text.py`
- Create: `tests/test_text_processor.py`

- [ ] **Step 1: Write failing text processor tests**

Create `tests/test_text_processor.py`:

```python
import pytest
from pptx.util import Pt

from report_generator.components.text import apply_text
from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping
from report_generator.pptx.document import PptxDocument


def text_component(**config):
    return ComponentMapping.model_validate(
        {
            "location": "text.summary",
            "semantic_description": "整体状态",
            "type": "Text",
            "config": config,
        }
    )


def test_apply_text_replaces_content(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["text.summary"].shape

    apply_text(shape, text_component(min_font_size=10), "新的摘要")

    assert shape.text_frame.text == "新的摘要"


def test_apply_text_shrinks_long_content(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["text.summary"].shape

    apply_text(shape, text_component(min_font_size=8), "这是一段较长的摘要内容，用来触发字号缩小逻辑。" * 3)

    assert shape.text_frame.paragraphs[0].runs[0].font.size <= Pt(18)
    assert shape.text_frame.paragraphs[0].runs[0].font.size >= Pt(8)


def test_apply_text_raises_when_content_cannot_fit(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["text.summary"].shape

    with pytest.raises(ReportGenerationError) as exc:
        apply_text(shape, text_component(min_font_size=16), "超长内容" * 400)

    assert exc.value.error_code == ErrorCode.TEXT_OVERFLOW
```

- [ ] **Step 2: Run tests and verify they fail**

Run:

```bash
python -m pytest tests/test_text_processor.py -q
```

Expected: import failure for `report_generator.components.text`.

- [ ] **Step 3: Create `src/report_generator/components/__init__.py`**

```python
"""Component processors."""
```

- [ ] **Step 4: Create `src/report_generator/components/text.py`**

```python
from __future__ import annotations

import math
from typing import Any

from pptx.util import Pt

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping

EMU_PER_INCH = 914400


def apply_text(shape: Any, component: ComponentMapping, value: Any) -> None:
    if not getattr(shape, "has_text_frame", False):
        raise ReportGenerationError(
            ErrorCode.TYPE_MISMATCH,
            f"组件 {component.location} 不是文本组件",
            component,
        )

    text = "" if value is None else str(value)
    min_font_size = int(component.config.get("min_font_size", 10))
    start_font_size = _existing_font_size(shape) or int(component.config.get("font_size", 18))
    fitted_font_size = _fit_font_size(shape, text, start_font_size, min_font_size)
    if fitted_font_size is None:
        raise ReportGenerationError(
            ErrorCode.TEXT_OVERFLOW,
            f"组件 {component.location} 的文本在最小字号 {min_font_size} 下仍无法放入模板区域",
            component,
        )

    text_frame = shape.text_frame
    text_frame.clear()
    lines = text.splitlines() or [""]
    for index, line in enumerate(lines):
        paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
        run = paragraph.add_run()
        run.text = line
        run.font.size = Pt(fitted_font_size)


def _existing_font_size(shape: Any) -> int | None:
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if run.font.size is not None:
                return int(round(run.font.size.pt))
    return None


def _fit_font_size(shape: Any, text: str, start: int, minimum: int) -> int | None:
    for font_size in range(start, minimum - 1, -1):
        if _fits(shape, text, font_size):
            return font_size
    return None


def _fits(shape: Any, text: str, font_size: int) -> bool:
    width_in = max(shape.width / EMU_PER_INCH, 0.1)
    height_in = max(shape.height / EMU_PER_INCH, 0.1)
    chars_per_line = max(1, int((width_in * 72) / (font_size * 0.55)))
    source_lines = text.splitlines() or [""]
    needed_lines = 0
    for line in source_lines:
        needed_lines += max(1, math.ceil(len(line) / chars_per_line))
    line_height_in = (font_size * 1.25) / 72
    capacity = max(1, math.floor(height_in / line_height_in))
    return needed_lines <= capacity
```

- [ ] **Step 5: Run text processor tests**

Run:

```bash
python -m pytest tests/test_text_processor.py -q
```

Expected: all tests pass.

- [ ] **Step 6: Commit**

```bash
git add src/report_generator/components tests/test_text_processor.py
git commit -m "feat: add text component processor"
```

---

### Task 6: Table Processor

**Files:**
- Create: `src/report_generator/components/table.py`
- Create: `tests/test_table_processor.py`

- [ ] **Step 1: Write failing table processor tests**

Create `tests/test_table_processor.py`:

```python
import pytest

from report_generator.components.table import apply_table
from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping
from report_generator.pptx.document import PptxDocument


def table_component(**config):
    return ComponentMapping.model_validate(
        {
            "location": "table.top_risks",
            "semantic_description": "TOP问题与风险",
            "type": "Table",
            "config": config,
        }
    )


def test_apply_table_rebuilds_rows_and_columns(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["table.top_risks"].shape
    data = [
        {"风险类型": "延期风险", "风险描述": "设备到货存在延期风险"},
        {"风险类型": "质量风险", "风险描述": "测试通过率偏低"},
    ]

    new_shape = apply_table(
        doc,
        shape,
        table_component(order=["风险类型", "风险描述"], min_font_size=8),
        data,
    )

    table = new_shape.table
    assert len(table.rows) == 3
    assert len(table.columns) == 2
    assert table.cell(0, 0).text == "风险类型"
    assert table.cell(2, 1).text == "测试通过率偏低"


def test_apply_table_rejects_too_many_rows(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["table.top_risks"].shape
    data = [{"风险类型": str(index), "风险描述": "x"} for index in range(3)]

    with pytest.raises(ReportGenerationError) as exc:
        apply_table(
            doc,
            shape,
            table_component(order=["风险类型", "风险描述"], max_rows=2),
            data,
        )

    assert exc.value.error_code == ErrorCode.TABLE_OVERFLOW


def test_apply_table_accepts_columns_rows_object(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["table.top_risks"].shape
    value = {
        "columns": [
            {"key": "type", "label": "风险类型"},
            {"key": "desc", "label": "风险描述"},
        ],
        "rows": [{"type": "延期风险", "desc": "设备到货存在延期风险"}],
    }

    new_shape = apply_table(doc, shape, table_component(), value)

    assert new_shape.table.cell(0, 0).text == "风险类型"
    assert new_shape.table.cell(1, 1).text == "设备到货存在延期风险"
```

- [ ] **Step 2: Run tests and verify they fail**

Run:

```bash
python -m pytest tests/test_table_processor.py -q
```

Expected: import failure for `report_generator.components.table`.

- [ ] **Step 3: Create `src/report_generator/components/table.py`**

```python
from __future__ import annotations

from typing import Any

from pptx.util import Pt

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping
from report_generator.pptx.document import PptxDocument


def apply_table(
    doc: PptxDocument,
    shape: Any,
    component: ComponentMapping,
    value: Any,
) -> Any:
    if not getattr(shape, "has_table", False):
        raise ReportGenerationError(
            ErrorCode.TYPE_MISMATCH,
            f"组件 {component.location} 不是表格组件",
            component,
        )

    columns, rows = _normalize_table(value, component)
    max_rows = int(component.config.get("max_rows", 30))
    max_columns = int(component.config.get("max_columns", 10))
    if len(rows) > max_rows:
        raise ReportGenerationError(
            ErrorCode.TABLE_OVERFLOW,
            f"组件 {component.location} 的数据行数 {len(rows)} 超过限制 {max_rows}",
            component,
        )
    if len(columns) > max_columns:
        raise ReportGenerationError(
            ErrorCode.TABLE_OVERFLOW,
            f"组件 {component.location} 的列数 {len(columns)} 超过限制 {max_columns}",
            component,
        )

    row_count = len(rows) + 1
    column_count = len(columns)
    if row_count == 0 or column_count == 0:
        raise ReportGenerationError(
            ErrorCode.DATA_SOURCE_INVALID,
            f"组件 {component.location} 的表格数据为空且无法推断列",
            component,
        )

    x, y, cx, cy = shape.left, shape.top, shape.width, shape.height
    slide = shape.part.slide
    doc.remove_shape(shape)
    new_shape = slide.shapes.add_table(row_count, column_count, x, y, cx, cy)
    new_shape.name = component.location
    table = new_shape.table

    for index in range(column_count):
        table.columns[index].width = int(cx / column_count)
    for index in range(row_count):
        table.rows[index].height = int(cy / row_count)

    font_size = _table_font_size(component, row_count)
    for col_index, column in enumerate(columns):
        _set_cell_text(table.cell(0, col_index), column["label"], font_size, bold=True)
    for row_index, row in enumerate(rows, start=1):
        for col_index, column in enumerate(columns):
            _set_cell_text(table.cell(row_index, col_index), row.get(column["key"], ""), font_size)

    return new_shape


def _normalize_table(value: Any, component: ComponentMapping) -> tuple[list[dict[str, str]], list[dict[str, Any]]]:
    if isinstance(value, dict) and "columns" in value and "rows" in value:
        columns = [
            {"key": str(column.get("key")), "label": str(column.get("label", column.get("key")))}
            for column in value["columns"]
        ]
        rows = [dict(row) for row in value["rows"]]
        return columns, rows

    if isinstance(value, list):
        rows = [dict(row) for row in value]
        order = component.config.get("order")
        if order:
            keys = [str(key) for key in order]
        else:
            keys = list(rows[0].keys()) if rows else []
        columns = [{"key": key, "label": key} for key in keys]
        return columns, rows

    raise ReportGenerationError(
        ErrorCode.DATA_SOURCE_INVALID,
        f"组件 {component.location} 的表格数据必须是对象数组或 columns/rows 对象",
        component,
    )


def _table_font_size(component: ComponentMapping, row_count: int) -> int:
    minimum = int(component.config.get("min_font_size", 8))
    preferred = int(component.config.get("font_size", 12))
    if row_count <= 8:
        return preferred
    return max(minimum, preferred - (row_count - 8))


def _set_cell_text(cell: Any, value: Any, font_size: int, bold: bool = False) -> None:
    cell.text = "" if value is None else str(value)
    for paragraph in cell.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(font_size)
            run.font.bold = bold
```

- [ ] **Step 4: Run table processor tests**

Run:

```bash
python -m pytest tests/test_table_processor.py -q
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add src/report_generator/components/table.py tests/test_table_processor.py
git commit -m "feat: add fixed-region table processor"
```

---

### Task 7: Image Processor

**Files:**
- Create: `src/report_generator/components/image.py`
- Create: `tests/test_image_processor.py`

- [ ] **Step 1: Add image shape fixture to `tests/conftest.py`**

Modify `simple_template_bytes` by adding this shape before `return _save(prs)`:

```python
    image_box = slide.shapes.add_textbox(Inches(6), Inches(1.1), Inches(1.2), Inches(0.7))
    image_box.name = "image.company_logo"
    image_box.text = "logo"
```

- [ ] **Step 2: Write failing image processor tests**

Create `tests/test_image_processor.py`:

```python
import base64
from io import BytesIO

from PIL import Image

from report_generator.components.image import apply_image
from report_generator.models import ComponentMapping
from report_generator.pptx.document import PptxDocument


def image_component(**config):
    return ComponentMapping.model_validate(
        {
            "location": "image.company_logo",
            "semantic_description": "公司 Logo",
            "type": "Image",
            "config": config,
        }
    )


def png_data_uri() -> str:
    image = Image.new("RGB", (10, 10), color=(255, 0, 0))
    output = BytesIO()
    image.save(output, format="PNG")
    encoded = base64.b64encode(output.getvalue()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def test_apply_image_replaces_region_with_picture(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["image.company_logo"].shape

    new_shape = apply_image(doc, shape, image_component(fit="contain"), png_data_uri())

    assert new_shape.name == "image.company_logo"
    assert "PICTURE" in str(new_shape.shape_type)
```

- [ ] **Step 3: Run tests and verify they fail**

Run:

```bash
python -m pytest tests/test_image_processor.py -q
```

Expected: import failure for `report_generator.components.image`.

- [ ] **Step 4: Create `src/report_generator/components/image.py`**

```python
from __future__ import annotations

import base64
from io import BytesIO
from typing import Any

import httpx

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping
from report_generator.pptx.document import PptxDocument


def apply_image(
    doc: PptxDocument,
    shape: Any,
    component: ComponentMapping,
    value: Any,
) -> Any:
    image_bytes = _load_image_bytes(component, value)
    x, y, cx, cy = shape.left, shape.top, shape.width, shape.height
    slide = shape.part.slide
    doc.remove_shape(shape)
    new_shape = slide.shapes.add_picture(BytesIO(image_bytes), x, y, width=cx, height=cy)
    new_shape.name = component.location
    return new_shape


def _load_image_bytes(component: ComponentMapping, value: Any) -> bytes:
    src = value.get("src") if isinstance(value, dict) else value
    if not isinstance(src, str) or not src:
        raise ReportGenerationError(
            ErrorCode.IMAGE_LOAD_FAILED,
            f"组件 {component.location} 的图片地址为空",
            component,
        )

    try:
        if src.startswith("data:"):
            return base64.b64decode(src.split(",", 1)[1])
        if src.startswith("http://") or src.startswith("https://"):
            response = httpx.get(src, timeout=10)
            response.raise_for_status()
            return response.content
        with open(src, "rb") as handle:
            return handle.read()
    except Exception as exc:
        raise ReportGenerationError(
            ErrorCode.IMAGE_LOAD_FAILED,
            f"组件 {component.location} 的图片加载失败: {exc}",
            component,
        ) from exc
```

- [ ] **Step 5: Run image processor tests**

Run:

```bash
python -m pytest tests/test_image_processor.py -q
```

Expected: all tests pass.

- [ ] **Step 6: Commit**

```bash
git add src/report_generator/components/image.py tests/conftest.py tests/test_image_processor.py
git commit -m "feat: add image component processor"
```

---

### Task 8: Chart and Shape Processors

**Files:**
- Create: `src/report_generator/components/chart.py`
- Create: `src/report_generator/components/shape.py`
- Create: `tests/test_chart_processor.py`

- [ ] **Step 1: Write failing chart and shape tests**

Create `tests/test_chart_processor.py`:

```python
from report_generator.components.chart import apply_chart
from report_generator.components.shape import apply_shape
from report_generator.models import ComponentMapping
from report_generator.pptx.document import PptxDocument


def chart_component(**config):
    return ComponentMapping.model_validate(
        {
            "location": "chart.revenue_trend",
            "semantic_description": "收入趋势",
            "type": "Chart",
            "config": config,
        }
    )


def shape_component(**config):
    return ComponentMapping.model_validate(
        {
            "location": "shape.status_badge",
            "semantic_description": "项目状态",
            "type": "Shape",
            "config": config,
        }
    )


def test_apply_chart_replaces_native_chart_data(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["chart.revenue_trend"].shape
    value = {
        "categories": ["Q1", "Q2", "Q3"],
        "series": [{"name": "收入", "values": [10, 20, 30]}],
    }

    apply_chart(shape, chart_component(max_categories=4, max_series=2), value)

    reopened = PptxDocument.open(doc.to_bytes())
    chart = reopened.shape_index()["chart.revenue_trend"].shape.chart
    assert len(chart.series[0].points) == 3


def test_apply_shape_updates_text(simple_template_bytes):
    doc = PptxDocument.open(simple_template_bytes)
    shape = doc.shape_index()["shape.status_badge"].shape

    apply_shape(shape, shape_component(), "风险")

    assert shape.text == "风险"
```

- [ ] **Step 2: Run tests and verify they fail**

Run:

```bash
python -m pytest tests/test_chart_processor.py -q
```

Expected: import failure for `report_generator.components.chart`.

- [ ] **Step 3: Create `src/report_generator/components/chart.py`**

```python
from __future__ import annotations

from typing import Any

from pptx.chart.data import CategoryChartData

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping


def apply_chart(shape: Any, component: ComponentMapping, value: Any) -> None:
    if not getattr(shape, "has_chart", False):
        raise ReportGenerationError(
            ErrorCode.TYPE_MISMATCH,
            f"组件 {component.location} 不是图表组件",
            component,
        )
    if not isinstance(value, dict):
        raise ReportGenerationError(
            ErrorCode.CHART_DATA_INVALID,
            f"组件 {component.location} 的图表数据必须是对象",
            component,
        )
    categories = value.get("categories", [])
    series = value.get("series", [])
    max_categories = int(component.config.get("max_categories", 24))
    max_series = int(component.config.get("max_series", 6))
    if len(categories) > max_categories or len(series) > max_series:
        raise ReportGenerationError(
            ErrorCode.CHART_DATA_INVALID,
            f"组件 {component.location} 的图表数据超过分类或系列数量限制",
            component,
        )

    chart_data = CategoryChartData()
    chart_data.categories = [str(category) for category in categories]
    for item in series:
        values = item.get("values", [])
        if len(values) != len(categories):
            raise ReportGenerationError(
                ErrorCode.CHART_DATA_INVALID,
                f"组件 {component.location} 的系列 {item.get('name')} 数据长度与分类数量不一致",
                component,
            )
        chart_data.add_series(str(item.get("name", "")), tuple(values))

    shape.chart.replace_data(chart_data)
```

- [ ] **Step 4: Create `src/report_generator/components/shape.py`**

```python
from __future__ import annotations

from typing import Any

from pptx.dml.color import RGBColor

from report_generator.models import ComponentMapping


def apply_shape(shape: Any, component: ComponentMapping, value: Any) -> None:
    if value is not None and hasattr(shape, "text"):
        shape.text = str(value)

    styles = component.config.get("state_styles", {})
    state_style = styles.get(str(value), {}) if isinstance(styles, dict) else {}
    fill = state_style.get("fill") or component.config.get("fill")
    if fill and hasattr(shape, "fill"):
        shape.fill.solid()
        shape.fill.fore_color.rgb = _rgb(fill)

    line = state_style.get("line") or component.config.get("line")
    if line and hasattr(shape, "line"):
        shape.line.color.rgb = _rgb(line)


def _rgb(value: str) -> RGBColor:
    normalized = value.lstrip("#")
    return RGBColor(
        int(normalized[0:2], 16),
        int(normalized[2:4], 16),
        int(normalized[4:6], 16),
    )
```

- [ ] **Step 5: Run chart and shape tests**

Run:

```bash
python -m pytest tests/test_chart_processor.py -q
```

Expected: all tests pass.

- [ ] **Step 6: Commit**

```bash
git add src/report_generator/components/chart.py src/report_generator/components/shape.py tests/test_chart_processor.py
git commit -m "feat: add chart and shape processors"
```

---

### Task 9: Generation Orchestrator

**Files:**
- Create: `src/report_generator/generator.py`
- Create: `tests/test_generator.py`

- [ ] **Step 1: Write failing generator tests**

Create `tests/test_generator.py`:

```python
import pytest

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.generator import generate_report
from report_generator.post_processing import PostProcessingRegistry
from report_generator.pptx.document import PptxDocument


def test_generate_report_populates_text_and_table(simple_template_bytes):
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
            },
            {
                "location": "table.top_risks",
                "semantic_description": "TOP问题与风险",
                "type": "Table",
                "config": {"order": ["风险类型", "风险描述"]},
                "data_source": {
                    "name": "risks",
                },
            },
        ],
    }
    payload = {
        "api_data": {"全量数据": {"项目概览": {"项目名称": "智慧园区"}}},
        "risks": [{"风险类型": "延期风险", "风险描述": "设备到货存在延期风险"}],
    }

    output = generate_report(simple_template_bytes, mapping, payload, PostProcessingRegistry())
    doc = PptxDocument.open(output)
    index = doc.shape_index()

    assert index["text.report_title"].shape.text_frame.text == "智慧园区-报告"
    assert index["table.top_risks"].shape.table.cell(1, 1).text == "设备到货存在延期风险"


def test_generate_report_rejects_missing_component(simple_template_bytes):
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

    with pytest.raises(ReportGenerationError) as exc:
        generate_report(simple_template_bytes, mapping, {"api_data": "value"}, PostProcessingRegistry())

    assert exc.value.error_code == ErrorCode.COMPONENT_NOT_FOUND
```

- [ ] **Step 2: Run tests and verify they fail**

Run:

```bash
python -m pytest tests/test_generator.py -q
```

Expected: import failure for `report_generator.generator`.

- [ ] **Step 3: Create `src/report_generator/generator.py`**

```python
from __future__ import annotations

from typing import Any

from report_generator.components.chart import apply_chart
from report_generator.components.image import apply_image
from report_generator.components.shape import apply_shape
from report_generator.components.table import apply_table
from report_generator.components.text import apply_text
from report_generator.datasource import resolve_component_value
from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.models import ComponentMapping, ReportMapping
from report_generator.post_processing import PostProcessingRegistry
from report_generator.pptx.document import PptxDocument, ShapeRef


def generate_report(
    template_bytes: bytes,
    mapping_payload: dict[str, Any],
    business_payload: dict[str, Any],
    registry: PostProcessingRegistry | None = None,
) -> bytes:
    registry = registry or PostProcessingRegistry()
    mapping = ReportMapping.model_validate(mapping_payload)
    doc = PptxDocument.open(template_bytes)
    index = doc.shape_index()

    for component in mapping.component_list:
        ref = _find_component(component, index)
        if component.visible is False:
            doc.remove_shape(ref.shape)
            continue
        value = resolve_component_value(component, business_payload, registry)
        _apply_component(doc, ref, component, value)
        index = doc.shape_index()

    return doc.to_bytes()


def _find_component(component: ComponentMapping, index: dict[str, ShapeRef]) -> ShapeRef:
    ref = index.get(component.location)
    if ref is None:
        raise ReportGenerationError(
            ErrorCode.COMPONENT_NOT_FOUND,
            f"模板中未找到组件 {component.location}",
            component,
        )
    return ref


def _apply_component(
    doc: PptxDocument,
    ref: ShapeRef,
    component: ComponentMapping,
    value: Any,
) -> None:
    if component.type == "Text":
        apply_text(ref.shape, component, value)
        return
    if component.type == "Image":
        apply_image(doc, ref.shape, component, value)
        return
    if component.type == "Table":
        apply_table(doc, ref.shape, component, value)
        return
    if component.type == "Chart":
        apply_chart(ref.shape, component, value)
        return
    if component.type == "Shape":
        apply_shape(ref.shape, component, value)
        return
    raise ReportGenerationError(
        ErrorCode.TYPE_MISMATCH,
        f"组件 {component.location} 的类型 {component.type} 尚未支持",
        component,
    )
```

- [ ] **Step 4: Run generator tests**

Run:

```bash
python -m pytest tests/test_generator.py -q
```

Expected: all tests pass.

- [ ] **Step 5: Run component regression tests**

Run:

```bash
python -m pytest tests/test_datasource.py tests/test_pptx_document.py tests/test_text_processor.py tests/test_table_processor.py tests/test_image_processor.py tests/test_chart_processor.py tests/test_generator.py -q
```

Expected: all tests pass.

- [ ] **Step 6: Commit**

```bash
git add src/report_generator/generator.py tests/test_generator.py
git commit -m "feat: orchestrate PPT report generation"
```

---

### Task 10: FastAPI Endpoint

**Files:**
- Create: `src/report_generator/api.py`
- Create: `tests/test_api.py`

- [ ] **Step 1: Write failing API tests**

Create `tests/test_api.py`:

```python
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
            "template": ("template.pptx", simple_template_bytes, "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
            "mapping": ("mapping.json", json.dumps(mapping).encode("utf-8"), "application/json"),
            "payload": ("payload.json", json.dumps(payload).encode("utf-8"), "application/json"),
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
            "template": ("template.pptx", simple_template_bytes, "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
            "mapping": ("mapping.json", json.dumps(mapping).encode("utf-8"), "application/json"),
            "payload": ("payload.json", json.dumps({"api_data": "x"}).encode("utf-8"), "application/json"),
        },
    )

    assert response.status_code == 400
    assert response.json()["detail"]["error_code"] == "COMPONENT_NOT_FOUND"
```

- [ ] **Step 2: Run tests and verify they fail**

Run:

```bash
python -m pytest tests/test_api.py -q
```

Expected: import failure for `report_generator.api`.

- [ ] **Step 3: Create `src/report_generator/api.py`**

```python
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
```

- [ ] **Step 4: Run API tests**

Run:

```bash
python -m pytest tests/test_api.py -q
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add src/report_generator/api.py tests/test_api.py
git commit -m "feat: expose PPT report generation endpoint"
```

---

### Task 11: Examples and Documentation

**Files:**
- Create: `examples/build_sample_template.py`
- Create: `examples/mapping.json`
- Create: `examples/payload.json`
- Modify: `README.md`

- [ ] **Step 1: Create `examples/mapping.json`**

```json
{
  "template_id": "project-monthly-report-ppt-v1",
  "component_list": [
    {
      "location": "text.report_title",
      "semantic_description": "报告标题",
      "type": "Text",
      "data_source": {
        "name": "api_data",
        "template": "{{ 全量数据.项目概览.项目名称 }}-报告"
      },
      "config": {
        "min_font_size": 12
      }
    },
    {
      "location": "table.top_risks",
      "semantic_description": "TOP问题与风险",
      "type": "Table",
      "config": {
        "order": ["风险类型", "风险描述"],
        "fit": "fixed_region",
        "min_font_size": 8,
        "max_rows": 20,
        "max_columns": 8
      },
      "data_source": {
        "name": "risks"
      }
    }
  ]
}
```

- [ ] **Step 2: Create `examples/payload.json`**

```json
{
  "api_data": {
    "全量数据": {
      "项目概览": {
        "项目名称": "智慧园区"
      }
    }
  },
  "risks": [
    {
      "风险类型": "延期风险",
      "风险描述": "设备到货存在延期风险"
    },
    {
      "风险类型": "质量风险",
      "风险描述": "测试通过率偏低"
    }
  ]
}
```

- [ ] **Step 3: Create `examples/build_sample_template.py`**

```python
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt


def main() -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(8), Inches(0.8))
    title.name = "text.report_title"
    title.text_frame.text = "报告标题"
    title.text_frame.paragraphs[0].runs[0].font.size = Pt(28)

    table_shape = slide.shapes.add_table(2, 2, Inches(0.5), Inches(1.6), Inches(8), Inches(2.5))
    table_shape.name = "table.top_risks"
    table = table_shape.table
    table.cell(0, 0).text = "风险类型"
    table.cell(0, 1).text = "风险描述"
    table.cell(1, 0).text = "示例类型"
    table.cell(1, 1).text = "示例描述"

    output = Path(__file__).with_name("sample_template.pptx")
    prs.save(output)
    print(output)


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Extend `README.md`**

Append:

```markdown
## Example

Build the sample template:

```bash
python examples/build_sample_template.py
```

Start the API:

```bash
uvicorn report_generator.api:app --reload
```

Generate a PPTX:

```bash
curl -X POST http://127.0.0.1:8000/reports/pptx \
  -F template=@examples/sample_template.pptx \
  -F mapping=@examples/mapping.json \
  -F payload=@examples/payload.json \
  --output examples/output_report.pptx
```
```

- [ ] **Step 5: Verify the example script**

Run:

```bash
python examples/build_sample_template.py
```

Expected: `examples/sample_template.pptx` is created.

- [ ] **Step 6: Run full test suite**

Run:

```bash
python -m pytest -q
```

Expected: all tests pass.

- [ ] **Step 7: Commit**

```bash
git add README.md examples/build_sample_template.py examples/mapping.json examples/payload.json
git commit -m "docs: add PPT generation example"
```

Do not commit `examples/sample_template.pptx` unless the team wants binary fixtures in git.

---

### Task 12: Final Verification

**Files:**
- No new files expected.

- [ ] **Step 1: Run complete tests**

Run:

```bash
python -m pytest -q
```

Expected: all tests pass.

- [ ] **Step 2: Build a sample template**

Run:

```bash
python examples/build_sample_template.py
```

Expected: command prints `examples/sample_template.pptx`.

- [ ] **Step 3: Run the API locally**

Run:

```bash
uvicorn report_generator.api:app --host 127.0.0.1 --port 8000
```

Expected: server starts and logs that it is listening on `http://127.0.0.1:8000`.

- [ ] **Step 4: Generate an example report from another terminal**

Run:

```bash
curl -sS -X POST http://127.0.0.1:8000/reports/pptx \
  -F template=@examples/sample_template.pptx \
  -F mapping=@examples/mapping.json \
  -F payload=@examples/payload.json \
  --output /tmp/report-generator-output.pptx
```

Expected: command exits with status 0.

- [ ] **Step 5: Verify output is a PPTX package**

Run:

```bash
python - <<'PY'
from pathlib import Path
path = Path("/tmp/report-generator-output.pptx")
assert path.exists(), "output file missing"
assert path.read_bytes()[:2] == b"PK", "output is not a zipped PPTX package"
print(path.stat().st_size)
PY
```

Expected: a positive file size is printed.

- [ ] **Step 6: Stop the API server**

Stop the `uvicorn` process with `Ctrl-C`.

- [ ] **Step 7: Check git status**

Run:

```bash
git status --short
```

Expected: no uncommitted tracked source changes. `examples/sample_template.pptx` may appear untracked if it was generated during verification.
