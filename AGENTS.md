# AGENTS.md

本仓库是一个 PPTX 报告生成服务：输入 PowerPoint 模板、mapping JSON 和业务 payload JSON，输出生成后的 PPTX 报告。后续 agent 修改本项目时，应优先保持模板样式不变，只替换数据；除非 mapping 显式配置样式，否则不要主动改字体、颜色、边框、行高或布局。

## 常用命令

```bash
uv run --extra dev python -m pytest -q
uv run uvicorn report_generator.api:app --reload
```

大模型后处理使用 OpenAI-compatible chat completions 接口，默认从 `.env` 或环境变量读取：

```dotenv
API_KEY=...
BASE_URL=https://example.com/v1
MODEL=gpt-4o-mini
COMPLETION_MODE=chat
LLM_CONCURRENCY=4
```

`COMPLETION_MODE=chat` 使用 `/chat/completions`；`COMPLETION_MODE=completions` 使用 legacy `/completions`。
`LLM_CONCURRENCY` 控制大模型后处理默认总并发数，默认值为 `4`。同步和异步接口也支持 multipart 表单字段 `llmConcurrency` 覆盖本次请求的并发数。

API 入口：

```http
POST /reports/pptx
multipart/form-data:
  template=@template.pptx
  mapping=@mapping.json
  payload=@payload.json
  llmConcurrency=4
```

异步生成接口：

```http
POST /reports/pptx/tasks
GET /reports/pptx/tasks?taskId=...
GET /reports/pptx/tasks/download?taskId=...
```

核心生成函数：

```python
from report_generator.generator import generate_report

output_bytes = generate_report(template_bytes, mapping_json, payload_json)
```

## 核心流程

1. 模板中每个可替换组件必须在 PowerPoint 选择窗格里设置唯一名称。
2. mapping 的 `component_list[].location` 必须与模板 shape 名称完全一致。
3. 生成时先按 mapping 找到模板组件，再通过 `data_source` 从 payload 取值，最后按组件类型写回 PPT。
4. `visible: false` 会删除该组件。
5. 映射到同一批组件里的 shape 名称不能重复，否则会报 `DUPLICATE_COMPONENT_NAME`。

## Mapping 结构

顶层结构保持与 Excel 报告生成能力接近：

```json
{
  "template_id": "project-report-v1",
  "component_list": [
    {
      "location": "text.report_title",
      "semantic_description": "报告标题",
      "type": "Text",
      "config": {
        "preserve_style": true
      },
      "data_source": {
        "name": "project",
        "template": "{{ 项目名称 }}项目进展报告"
      }
    }
  ]
}
```

组件字段：

- `location`: PPT 选择窗格里的 shape 名称。
- `semantic_description`: 组件语义说明，便于理解和生成 mapping。
- `type`: 组件类型。当前实际支持 `Text`、`Image`、`Table`、`Chart`、`Shape`、`TopIssues`、`Milestone`。`GanttChart` 只是模型中预留的类型，当前生成器未实现，mapping 中不要使用。
- `config`: 组件配置。默认应少配样式，保留模板原样式。
- `data_source`: 数据来源描述。
- `visible`: 为 `false` 时删除组件。
- `prompt`、`data_example`: 可保留为上游生成 mapping 的辅助字段，运行时不强依赖。

`data_source` 支持：

- `name`: 从 payload 顶层读取同名数据。
- `index`: 在 `name` 指定的数据源内执行 JSONPath；未设置 `name` 时基于整个 payload。
- `template`: 用 Jinja 模板渲染文本；未设置 `name` 时基于整个 payload。
- `needs_post_processing`: 为 `true` 时必须经过大模型处理。若 `name` 命中注册函数，则先调用函数精简/整理数据，再把函数返回值结合 `semantic_description`、`prompt`、`data_example` 传给大模型，生成符合组件入参的数据；未命中注册函数时，则把 `index`、`template` 或 `name` 解析出的数据直接传给大模型。
- `params`: 注册函数参数映射，格式为 `{"函数参数名": "payload顶层字段名"}`。`"$"` 表示传入整个 payload；为兼容 Excel 报告，`{"sr_api_data": "sr_api_data"}` 在 payload 没有顶层 `sr_api_data` 时会回退传入整个 payload。

内置注册函数：

- `general_project_delivery_plan`: 从 `全量数据.项目阶段` 输出项目交付计划/任务表对象数组。
- `general_top_risks_and_issues`: 从 `周期数据.问题与风险` 或 `全量数据.问题与风险` 输出表格行。
- `general_top_risks_and_issues_v2`: 从问题与风险输出 `TopIssues` 组件需要的 `{ "items": [...] }` 卡片数据。

示例：

```json
{
  "data_source": {
    "name": "project",
    "index": "$.基础信息.项目名称"
  }
}
```

```json
{
  "data_source": {
    "name": "project",
    "template": "报告周期：{{ 开始日期 }}至{{ 结束日期 }}"
  }
}
```

大模型后处理示例：

```json
{
  "location": "table.top_issues_risks",
  "semantic_description": "TOP 问题及风险动态表格",
  "prompt": "从原始项目日志中提取最重要的两个问题或风险，输出 columns/rows 表格数据。",
  "data_example": {
    "columns": [
      {"key": "序号", "label": "序号"},
      {"key": "问题风险与描述", "label": "问题风险与描述"}
    ],
    "rows": [
      {"序号": "1", "问题风险与描述": "验收排期需确认"}
    ]
  },
  "type": "Table",
  "data_source": {
    "name": "raw_project_logs",
    "needs_post_processing": true
  }
}
```

函数型数据源直接渲染示例：

```json
{
  "location": "top_issues.cards",
  "semantic_description": "TOP问题与风险动态卡片列表",
  "type": "TopIssues",
  "data_source": {
    "name": "general_top_risks_and_issues_v2",
    "params": {
      "sr_api_data": "sr_api_data"
    }
  }
}
```

函数型数据源作为大模型前置处理示例：

```json
{
  "location": "text.key_issue_summary",
  "semantic_description": "关键问题摘要",
  "type": "Text",
  "prompt": "根据输入的问题风险列表，提炼一句关键问题摘要。",
  "data_source": {
    "name": "general_top_risks_and_issues_v2",
    "params": {
      "sr_api_data": "sr_api_data"
    },
    "needs_post_processing": true
  }
}
```

## 支持的组件

### Text

用途：替换文本框、标题、占位符中的文本。

数据：普通值会转成字符串。富文本使用 `rich_text` 数组；为兼容 PPT run 语义，也支持 `runs` 作为别名。

常用配置：

- `preserve_style: true`: 保留原 run 样式，仅替换文本。若原字号放不下新文本，会报 `TEXT_OVERFLOW`。
- `font_size`: 未保留样式时的起始字号。
- `min_font_size`: 自动缩小字号时的最小字号。

建议：模板文本组件默认使用 `preserve_style: true`，让字号、颜色、粗细由模板决定。

富文本示例：

```json
{
  "rich_text": [
    {
      "text": "关键结论：",
      "color": "0052CC",
      "font_size": 16,
      "font_name": "Microsoft YaHei",
      "bold": true
    },
    {
      "text": "整体推进中",
      "color": "333333",
      "font_size": 14
    }
  ]
}
```

支持的片段字段：`text`、`color`、`font_size`、`font_name`、`bold`、`italic`、`underline`。换行直接写在任意片段的 `text` 中，例如 `"关键结论：第一行\n第二行"`；生成时会拆成 PPT 段落，并把该片段样式应用到换行后的文本。

### Table

用途：替换表格数据。表格默认优先在原表格内填充并动态增行，尽量保留模板表格样式。

常用配置：

- `max_rows`: 最大数据行数，默认 `30`。
- `max_columns`: 最大列数，默认 `10`。
- `order`: 当数据是对象数组时，指定列顺序。
- `font_size`: 写入后强制设置字号。
- `preserve_style: true`: 严格原位填充，不增行、不增列；数据超过模板预留行列会报 `TABLE_OVERFLOW`。
- `mode: "placeholders"`: 表格内占位符填充模式，不重建表格。

表格有三种数据形态：

1. 普通动态表格，适合列表数据：

```json
{
  "columns": [
    {"key": "task", "label": "任务"},
    {"key": "owner", "label": "负责人"}
  ],
  "rows": [
    {"task": "接口联调", "owner": "张明"},
    {"task": "验收测试", "owner": "李强"}
  ]
}
```

2. 对象数组，列由 `config.order` 或第一行 key 推断：

```json
[
  {"阶段名称": "开发联调", "任务名称": "接口联调", "任务状态": "完成"},
  {"阶段名称": "验收测试", "任务名称": "测试执行", "任务状态": "进行中"}
]
```

3. `cells` 矩阵，适合少量非表头网格；固定键值表单更推荐占位符模式：

```json
{
  "cells": [
    ["项目名称", "智能制造中台项目", "项目编码", "PRJ-2025-1225"],
    ["客户名称", "星河科技集团", "项目 PD", "张明"]
  ]
}
```

固定键值表单推荐使用占位符模式。模板表格的值单元格写入：

```text
{{ 项目名称 }}
{{ 项目编码 }}
{{ 客户名称 }}
```

mapping：

```json
{
  "location": "table.project_info",
  "semantic_description": "项目信息键值表单，占位符填充",
  "type": "Table",
  "config": {
    "mode": "placeholders"
  },
  "data_source": {
    "name": "project_info"
  }
}
```

payload：

```json
{
  "project_info": {
    "项目名称": "智能制造中台项目",
    "项目编码": "PRJ-2025-1225",
    "客户名称": "星河科技集团"
  }
}
```

### Image

用途：用图片替换模板中的占位 shape。

数据可以是：

- 本地文件路径。
- HTTP/HTTPS URL。
- base64 data URL。
- 对象形式：`{"src": "..."}`。

生成时会按原 shape 的 `left/top/width/height` 插入图片，并删除原 shape。图片占位符的边框、填充等形状样式不会保留；需要视觉框时，应在模板中使用独立背景/边框形状。

### Chart

用途：替换 PowerPoint 原生图表的数据，尽量保留原图表样式。

数据结构：

```json
{
  "categories": ["一月", "二月", "三月"],
  "series": [
    {"name": "完成数", "values": [10, 18, 24]},
    {"name": "延期数", "values": [1, 2, 1]}
  ]
}
```

常用配置：

- `max_categories`: 最大分类数，默认 `24`。
- `max_series`: 最大系列数，默认 `6`。

要求：每个 series 的 `values` 长度必须与 `categories` 一致，且 values 必须是数字。

### Shape

用途：更新普通形状文本，或根据状态改形状颜色/线条。

数据：任意值，若 shape 有 `text` 属性，会转成字符串写入。

配置：

```json
{
  "fill": "#00A896",
  "line": "#028090",
  "state_styles": {
    "正常": {"fill": "#00A896"},
    "风险": {"fill": "#F96167", "line": "#990011"}
  }
}
```

只有配置了 `fill`、`line` 或 `state_styles` 时才改样式。

### TopIssues

用途：生成“TOP 问题与风险”动态卡片列表。

模板：在 PPT 中放一个可见的示例卡片或普通 shape 作为预览锚点，命名为例如 `top_issues.cards`。如需让用户在模板里看到完整卡片效果，可在锚点上方放示例色条、严重程度、描述、措施、元信息等元素，并统一命名为 `top_issues.cards.preview.*`。生成时使用 `preview_mode: "replace"` 删除锚点和所有 `top_issues.cards.preview.*` 预览元素，并从锚点的 `left/top/width/height` 开始垂直绘制真实卡片。卡片数量按数据动态生成；超出页面空间时继续向下排布，不自动分页。

数据可以是数组，也可以是包含 `items` 的对象：

```json
[
  {
    "severity": "紧急",
    "created_at": "2026-06-06 10:30",
    "description": "客户验收窗口尚未最终确认，可能影响验收排期。",
    "action": "已协调客户项目经理锁定评审窗口。",
    "owner": "张明",
    "status": "跟踪中",
    "due_date": "2026-06-14"
  }
]
```

内置三种严重等级样式：`紧急`、`重要`、`一般`。可通过 `config.styles` 覆盖颜色：

```json
{
  "location": "top_issues.cards",
  "type": "TopIssues",
  "config": {
    "preview_mode": "replace",
    "card_height": 0.66,
    "card_gap": 0.1,
    "styles": {
      "紧急": {"accent": "D64545"},
      "重要": {"accent": "FF8A00"},
      "一般": {"accent": "F5C400"}
    }
  },
  "data_source": {"name": "top_issues"}
}
```

默认情况下，`问题描述：`、`解决措施与进展：` 标签使用蓝色粗体，标签后的具体内容使用黑色常规字重。`description_color`、`action_color` 兼容作为标签颜色；也可用 `description_label_color`、`description_value_color`、`action_label_color`、`action_value_color` 分别配置标签和内容颜色。

可配置 `description_template`、`action_template`、`meta_template` 自定义卡片文本。模板上下文就是单条 issue 对象。

### Milestone

用途：生成动态里程碑时间轴，节点数量按数据动态增减。

模板：在 PPT 中放一个可见的示例时间轴或普通 shape 作为预览锚点，命名为例如 `milestone.delivery`。设置 `preview_mode: "replace"` 时，生成时会删除或改名该预览锚点，并在同一区域内等距绘制真实连线、节点、日期和阶段名称。若锚点本身有填充或线条样式，会保留为 `milestone.delivery.background` 背景 shape，因此适合把黄色底色、圆角、阴影等容器样式放在这个命名锚点上。未设置 `preview_mode` 时，为兼容旧模板，默认保留锚点并清空其文字。

数据可以是数组，也可以是包含 `items` 的对象：

```json
{
  "items": [
    {"label": "准备", "date": "04-20", "status": "done"},
    {"label": "安装", "date": "05-18", "status": "done"},
    {"label": "调测", "date": "06-20", "status": "active"},
    {"label": "验收", "date": "07-05", "status": "pending"}
  ]
}
```

内置状态样式：`done`、`active`、`pending`。可通过 `config.status_styles` 覆盖节点填充、线条和文字颜色：

```json
{
  "location": "milestone.delivery",
  "type": "Milestone",
  "config": {
    "preview_mode": "replace",
    "node_size": 0.14,
    "date_width": 0.9,
    "text_axis_gap": 0.09,
    "status_styles": {
      "risk": {"fill": "FFF5D6", "line": "D99A00", "text": "333333"}
    }
  },
  "data_source": {"name": "milestones"}
}
```

`date_width` 控制日期文本框宽度，适合 `YYYY-MM-DD` 这类较长日期；`text_axis_gap` 控制日期文本框底部、阶段名文本框顶部到坐标轴的垂直距离，默认上下相等，避免日期和阶段名看起来不对称。

日期和阶段名默认字号都是 14pt；阶段名默认字体是 `Microsoft YaHei`，可通过 `label_font_name` 覆盖。

节点默认不使用 shape 描边线，避免小尺寸圆点在导出 PDF/PNG 时边缘发虚。`active` 和 `pending` 默认画成“外圆 + 内圆”的空心节点，其中外圆使用 `line` 色，内圆使用 `fill` 色；`done` 默认画成实心节点。可通过 `hollow_statuses` 指定哪些状态使用空心节点，通过 `node_inner_ratio` 控制内圆相对外圆的尺寸；如确实需要描边，可显式配置 `node_outline_width`。

```json
{
  "config": {
    "hollow_statuses": ["active", "pending", "risk"],
    "node_inner_ratio": 0.52,
    "status_styles": {
      "risk": {"fill": "FFF5D6", "line": "D99A00", "text": "333333", "hollow": true}
    }
  }
}
```

## 如何构造模板

### 通用规则

1. 先在 PowerPoint 中完成全部视觉设计。
2. 需要替换的组件在选择窗格中改成稳定名称，例如 `text.cover_title`、`table.task_details`、`chart.progress`。
3. 同一个 mapping 里引用的组件名必须唯一。
4. 不要依赖默认的 `TextBox 1`、`Table 4` 这类自动名称。
5. 固定位置固定字段优先用占位符；动态列表优先用动态表格。
6. 模板样式应该由 PPT 决定，mapping 只表达数据和必要限制。

### 文本模板

文本框先放一段示例文本，设置好字号、颜色、粗细和位置。mapping 中使用 `Text` + `preserve_style: true`。

```json
{
  "location": "text.cover_title",
  "type": "Text",
  "config": {"preserve_style": true},
  "data_source": {
    "name": "project",
    "template": "{{ 项目名称 }}项目进展报告"
  }
}
```

### 固定表单模板

适合项目信息、客户信息、基本信息等键值表格：

1. 在 PPT 表格中保留固定行列。
2. 标签单元格写固定文字，如 `项目名称`。
3. 值单元格写 Jinja 占位符，如 `{{ 项目名称 }}`。
4. mapping 使用 `Table` + `config.mode = "placeholders"`。
5. payload 提供一个对象，key 与占位符变量一致。

这种模式不会重建表格，最能保持原表格边框、行高、列宽和填充。

### 动态表格模板

适合任务列表、风险列表、交付计划等行数/列数可能变化的数据：

1. 在 PPT 中放一个示例表格，至少包含表头行和一行正文样式。
2. 表头行负责提供表头样式。
3. 正文行负责提供新增行复制样式。
4. mapping 使用 `Table`，不要设置 `preserve_style`。
5. 数据使用 `columns/rows` 或对象数组。

默认动态模式会优先原位填充表格。当数据列数不超过模板列数时，生成器会保留原表格 shape，按需复制模板最后一行追加数据行，因此能更好保留表头、正文行、边框、填充、行高和主题样式。只有当数据列数超过模板列数、需要动态增列时，才会退回到重建表格并复制可用单元格样式。

如果表格行列必须固定，设置 `preserve_style: true`，但这样不会动态增行/增列。

### 图表模板

1. 在 PPT 中插入真实 PowerPoint 图表。
2. 设置好图表颜色、坐标轴、图例、标题等样式。
3. 在选择窗格中命名为 `chart.xxx`。
4. mapping 使用 `Chart`，payload 传入 `categories` 和 `series`。

### 图片模板

1. 在 PPT 中放一个图片占位 shape 或示例图片。
2. 在选择窗格中命名为 `image.xxx`。
3. mapping 使用 `Image`。
4. payload 提供图片路径、URL 或 data URL。

注意：图片组件会删除原 shape 并插入新图片；若需要固定边框或背景，请把边框/背景做成单独 shape，不要依赖图片占位 shape 自身样式。

### TOP 问题与风险模板

适合问题/风险卡片数量变化的页面：

1. 在需要开始绘制卡片的位置放一个可见卡片背景 shape。
2. 将这个定位锚点命名为 `top_issues.cards`。
3. 如需让用户在模板中看到示例卡片内容，可在锚点上方放示例色条、严重程度、描述、措施、元信息，并统一命名为 `top_issues.cards.preview.*`。
4. mapping 使用 `TopIssues`。
5. mapping 中设置 `config.preview_mode = "replace"`。
6. payload 传入 issue 数组。

`TopIssues` 当前是程序生成卡片，预览元素只用于模板可视化和定位，不复制模板中的复杂组合卡片；如需完全复刻 PPT 手工设计的卡片内部布局，需要后续扩展模板组复制模式。

### 里程碑模板

适合交付计划、项目阶段、关键节点：

1. 在时间轴区域放一个可见容器 shape，例如黄色圆角矩形。
2. 将该容器 shape 命名为 `milestone.delivery`，它同时负责定位和提供背景样式。
3. 如需让用户在模板中看到示例节点，可在容器上方放示例线条、节点、日期、标签，并统一命名为 `milestone.delivery.preview.*`。
4. mapping 使用 `Milestone`。
5. mapping 中设置 `config.preview_mode = "replace"`。
6. payload 传入里程碑数组。

`Milestone` 当前按预览 shape 宽度等距分布节点，不按真实日期比例定位。`preview_mode: "replace"` 会删除 `milestone.delivery.preview.*` 示例元素；如果 `milestone.delivery` 有填充或线条样式，则保留并重命名为 `milestone.delivery.background`，真实节点绘制在它上方。

## 示例文件

- `examples/general_info_template.pptx`: 固定项目信息表单，占位符填充。
- `examples/general_info_mapping.json`: 占位符表格 mapping 示例。
- `examples/general_info_payload.json`: 占位符表格 mock payload。
- `examples/general_info_report.pptx`: 由上述模板和 payload 生成的报告。
- `examples/project_progress_template.pptx`: 多页项目进展报告模板。
- `examples/project_progress_mapping.json`: 多组件 mapping 示例。
- `examples/project_progress_payload.json`: 多组件 mock payload。
- `examples/project_progress_report.pptx`: 生成后的项目进展报告。

## 验证要求

修改代码或示例后至少运行：

```bash
uv run --extra dev python -m pytest -q
```

涉及 PPTX 模板或生成结果时，建议用 LibreOffice 渲染成 PDF/PNG 做视觉检查，重点看：

- 文本是否溢出或被裁切。
- 表格边框、行高、填充色是否保留。
- 占位符是否全部替换。
- 图片是否按预期位置和比例显示。
- 图表系列和分类是否正确。

不要把临时渲染产物提交进仓库，除非用户明确要求。
