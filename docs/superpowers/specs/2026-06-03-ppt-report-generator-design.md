# PPT Report Generator Design

Date: 2026-06-03
Workspace: `/Users/oam/Workspace/demos/report_generator`

## Goal

Build an HTTP service that generates a final `.pptx` report from:

- A PowerPoint template.
- A component mapping JSON.
- A business payload JSON.

The mapping JSON should stay as close as possible to the existing Excel report generator format. The top-level structure remains `template_id` plus `component_list`; each component keeps fields such as `location`, `semantic_description`, `type`, `prompt`, `data_example`, `config`, and `data_source`.

## Non-Goals

- Do not create a visual template editing backend in the first version.
- Do not automatically duplicate slides or sections based on data length.
- Do not let components push, resize, or reflow neighboring components.
- Do not replace the existing Excel report generation protocol with a separate PPT-only protocol.

## Core Principles

- The PPT template owns layout, pages, component positions, z-order, theme, and visual style.
- JSON owns component data, data source resolution, prompts, post-processing, visibility, and limited rendering constraints.
- Components are bound by PowerPoint selection pane names.
- The `location` field keeps its name for Excel compatibility, but its PPT meaning is the PowerPoint shape name instead of an Excel range.
- The template page count is fixed.
- Text fits inside the original text box. It can shrink to a configured minimum font size; below that, generation fails.
- Tables may grow rows and columns, but only inside the original template table region. If they cannot fit within readable limits, generation fails.
- Charts are generated as native editable PPT charts where possible. Complex charts may fall back to image rendering when configured.

## JSON Protocol

The PPT mapping JSON keeps the Excel report generator shape:

```json
{
  "template_id": "project-monthly-report-ppt-v1",
  "component_list": [
    {
      "location": "text.report_title",
      "semantic_description": "报告标题",
      "type": "Text",
      "data_example": "{项目名称}-报告",
      "config": {
        "min_font_size": 12,
        "overflow": "shrink_then_error"
      },
      "data_source": {
        "name": "api_data",
        "template": "{{ 全量数据.项目概览.项目名称 }}-报告"
      }
    },
    {
      "location": "table.top_risks",
      "semantic_description": "TOP问题与风险",
      "type": "Table",
      "prompt": "无风险时仍输出表格结构，保留列头。",
      "config": {
        "order": ["风险类型", "风险描述", "责任人", "解决日期", "状态"],
        "fit": "fixed_region",
        "min_font_size": 8,
        "max_rows": 20,
        "max_columns": 8,
        "header_from_template": true,
        "body_style_from_template": true
      },
      "data_source": {
        "name": "general_top_risks_and_issues",
        "params": {
          "sr_api_data": "sr_api_data"
        },
        "needs_post_processing": true
      }
    }
  ]
}
```

### Field Meanings

- `template_id`: Logical template identifier.
- `component_list`: Ordered list of template components to populate.
- `location`: In Excel this is a cell range, such as `D4:F4`. In PPT this is the PowerPoint shape name, such as `text.report_title` or `table.top_risks`.
- `semantic_description`: Human-readable business meaning of the component.
- `type`: Component type. First-version types are `Text`, `Image`, `Table`, `Chart`, and `Shape`. Business-specific types such as `Milestone` and `GanttChart` are reserved for later specialized processors.
- `prompt`: Optional instruction for post-processing or LLM summarization.
- `data_example`: Optional example that constrains output shape.
- `config`: Component-specific rendering and validation constraints.
- `data_source`: Existing Excel-compatible data source descriptor.

### Data Source Compatibility

The PPT service should reuse the existing Excel data source semantics:

- `data_source.index`: Resolve a value from the business payload by JSONPath.
- `data_source.template`: Render a template expression against the business payload.
- `data_source.name`: Resolve a named source or post-processing function.
- `data_source.params`: Provide named inputs for a post-processing function.
- `data_source.needs_post_processing`: Route the extracted data through post-processing before rendering.

The service should not require callers to flatten business payloads into PPT-specific component values.

## HTTP API

### Generate PPT Report

```http
POST /reports/pptx
Content-Type: multipart/form-data

template: .pptx
mapping: .json
payload: .json
```

Request parts:

- `template`: PowerPoint template file.
- `mapping`: Component mapping JSON using the protocol above.
- `payload`: Business data JSON containing sources such as `api_data` and `sr_api_data`.

Successful response:

- `200 OK`
- `Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation`
- Body is the generated `.pptx`.

Failure response:

- Structured JSON error with component context when applicable.

### Future API Extensions

- `GET /templates/{template_id}/components`: Scan a template and return detected shape names and inferred component types.
- `POST /reports/pptx/validate`: Validate template, mapping, and payload without generating the final file.
- `POST /reports/pptx/preview`: Generate slide thumbnails for review.
- `POST /templates`: Register templates server-side instead of uploading on every request.

## Template Conventions

Template authors must name controllable components in PowerPoint's selection pane.

Recommended naming:

- `text.report_title`
- `text.project_summary`
- `image.company_logo`
- `table.top_risks`
- `chart.revenue_trend`
- `shape.status_badge`

Rules:

- Every `component_list.location` must match exactly one shape name in the template.
- Duplicate controllable names are invalid in the first version.
- The shape's position and dimensions define its rendering region.
- The shape's existing style is the default style source.
- Non-listed shapes remain unchanged.

## Generation Flow

1. Receive `template`, `mapping`, and `payload`.
2. Validate that the mapping JSON has `template_id` and `component_list`.
3. Open the PPTX package and scan all slides for shape names.
4. Build a shape index keyed by selection pane name.
5. Validate that each `component_list.location` exists and matches the declared component type where this can be inferred.
6. Resolve data for each component using the Excel-compatible `data_source` semantics.
7. Run post-processing when `needs_post_processing` is true.
8. Dispatch each component to a type-specific processor.
9. Save the generated PPTX.
10. Run final checks for missing components, unresolved placeholders, text overflow, table overflow, chart data errors, and image load failures.
11. Return the generated `.pptx` or a structured error.

## Component Processors

### Text

Responsibilities:

- Replace text content in the named text shape.
- Preserve the template's font family, color, alignment, paragraph style, and text box geometry.
- Support plain text and multiline text.
- Wrap text inside the existing text box.
- Shrink font size until content fits or until `config.min_font_size` is reached.

Failure conditions:

- The content still does not fit after shrinking to `config.min_font_size`.
- The target shape is not a text-capable shape.

Default config:

```json
{
  "overflow": "shrink_then_error",
  "min_font_size": 10
}
```

### Image

Responsibilities:

- Replace the named image region with an image from URL, local file reference, or base64 payload.
- Preserve the original geometry.
- Support `contain`, `cover`, and `stretch` fit modes.
- Optionally hide the component when data is missing and `config.missing` allows it.

Default config:

```json
{
  "fit": "contain",
  "missing": "error"
}
```

Failure conditions:

- The image cannot be loaded.
- The image format is unsupported.
- The target component cannot be used as an image region.

### Table

Responsibilities:

- Treat the template table as a style sample and fixed rendering region.
- Rebuild rows and columns from resolved data.
- Apply configured column order from `config.order` when present.
- Preserve header style, body style, borders, fills, alignment, and number formatting where possible.
- Fit the generated table inside the original region by adjusting column widths, row heights, and font size.

First-version table input should be an array of row objects or an object containing `columns` plus `rows`.

Example resolved data:

```json
[
  {
    "风险类型": "延期风险",
    "风险描述": "设备到货存在延期风险",
    "责任人": "张三",
    "解决日期": "2026.06.10",
    "状态": "Open"
  }
]
```

Failure conditions:

- Row count exceeds `config.max_rows`.
- Column count exceeds `config.max_columns`.
- The table cannot fit inside the original region above `config.min_font_size`.
- The target shape is not a table or table-like placeholder.

Default config:

```json
{
  "fit": "fixed_region",
  "min_font_size": 8,
  "max_rows": 30,
  "max_columns": 10,
  "header_from_template": true,
  "body_style_from_template": true
}
```

### Chart

Responsibilities:

- Prefer native editable PPT chart output.
- Preserve template chart type, style, colors, title, axes, legend, and geometry.
- Replace categories, series names, and values.
- Optionally fall back to a rendered image chart when configured.

Example resolved data:

```json
{
  "categories": ["Q1", "Q2", "Q3", "Q4"],
  "series": [
    {
      "name": "收入",
      "values": [100, 140, 130, 180]
    }
  ]
}
```

Default config:

```json
{
  "mode": "native",
  "fallback": "image",
  "max_categories": 24,
  "max_series": 6
}
```

Failure conditions:

- Native chart data cannot be updated.
- Category or series counts exceed configured limits and fallback is disabled.
- Values are not numeric where numeric series are required.

### Shape

Responsibilities:

- Control simple visual state while preserving geometry.
- Support text, fill color, line color, transparency, and visibility.
- Support business-specific state mapping through config, such as status badge colors.

Example config:

```json
{
  "state_styles": {
    "正常": { "fill": "#2E7D32", "text_color": "#FFFFFF" },
    "风险": { "fill": "#F9A825", "text_color": "#111111" },
    "异常": { "fill": "#C62828", "text_color": "#FFFFFF" }
  }
}
```

Failure conditions:

- Unsupported style keys are supplied in strict mode.
- The target shape cannot apply the requested style.

## Visibility

All component processors should support optional visibility control:

```json
{
  "location": "shape.optional_badge",
  "type": "Shape",
  "visible": false
}
```

If `visible` is false, the service hides or removes the target shape. This must not trigger layout changes in other components.

## Error Semantics

Errors should be structured and include component context where possible:

```json
{
  "error_code": "TEXT_OVERFLOW",
  "message": "组件 text.summary 的文本在最小字号 10 下仍无法放入模板区域",
  "component": {
    "location": "text.summary",
    "type": "Text",
    "semantic_description": "整体状态"
  }
}
```

Required error codes:

- `COMPONENT_NOT_FOUND`
- `DUPLICATE_COMPONENT_NAME`
- `TYPE_MISMATCH`
- `DATA_SOURCE_NOT_FOUND`
- `DATA_SOURCE_INVALID`
- `POST_PROCESSING_FAILED`
- `TEXT_OVERFLOW`
- `TABLE_OVERFLOW`
- `CHART_DATA_INVALID`
- `IMAGE_LOAD_FAILED`
- `PPTX_PARSE_FAILED`
- `PPTX_RENDER_FAILED`

## MVP Scope

The first implementation should include:

- FastAPI HTTP service.
- `POST /reports/pptx` synchronous generation endpoint.
- Excel-style mapping JSON compatibility.
- PPT shape name scanning.
- Component lookup by `location`.
- Data source resolution for `index`, `template`, and named post-processing hooks.
- `Text`, `Image`, and `Table` processors.
- Basic `Chart` processor for native chart data replacement when the template chart is simple.
- Image fallback for chart only when explicitly configured.
- Structured errors.
- Example PPT template, mapping JSON, payload JSON, and output fixture.
- Automated tests for data source resolution, component matching, text overflow, table growth, and error responses.

## Implementation Notes

Use a layered design:

- `api`: HTTP request parsing, response formatting, and error mapping.
- `mapping`: schema validation for Excel-compatible component JSON.
- `datasource`: JSONPath, template expression rendering, and post-processing function dispatch.
- `pptx`: low-level PPTX package scanning and saving.
- `components`: type-specific processors.
- `validation`: pre-generation and post-generation checks.

The underlying PPTX manipulation can be a hybrid of high-level libraries and direct Open XML editing. The design should avoid depending on a single library abstraction for all PPTX features because native charts and shape names often require direct XML access.

## Testing Strategy

Unit tests:

- Validate mapping schema.
- Resolve JSONPath values.
- Render template expressions.
- Dispatch data source functions with params.
- Detect missing or duplicate component names.
- Validate text overflow behavior.
- Validate table row and column limit behavior.

Integration tests:

- Generate a PPTX from a sample template and sample payload.
- Verify that expected text appears in the output.
- Verify that a dynamic table has the expected row and column counts.
- Verify that missing components return `COMPONENT_NOT_FOUND`.
- Verify that oversized text returns `TEXT_OVERFLOW`.
- Verify that oversized tables return `TABLE_OVERFLOW`.

Manual QA for first templates:

- Open generated PPTX in PowerPoint or LibreOffice.
- Check that text is not clipped.
- Check that tables stay within their original region.
- Check that images preserve intended aspect ratio.
- Check that charts remain editable when native mode is used.

## Open Implementation Decisions

These do not block the design but should be decided during implementation planning:

- Exact Python PPTX/Open XML library mix.
- Whether missing optional images should default to `error`, `hide`, or placeholder.
- Whether the service stores uploaded files temporarily only or also supports registered templates.
- Whether chart image fallback uses ECharts, Matplotlib, or another renderer.
- Whether post-processing functions are local Python callables, HTTP callbacks, or both.
