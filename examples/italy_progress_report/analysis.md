# 进度报告模板 1 处理与能力缺口分析

## 处理结果

源文件：`/Users/oam/Desktop/意大利/PPT 生成/进度报告模板1.pptx`

已生成文件：

- `examples/italy_progress_report/source_progress_template1.pptx`: 源 PPT 备份，未改动。
- `examples/italy_progress_report/progress_template1_named.pptx`: 已处理后的 PPT 模板，关键可替换组件已在选择窗格中命名。
- `examples/italy_progress_report/progress_template1_mapping.json`: 组件 mapping。
- `examples/italy_progress_report/progress_template1_payload.json`: mock 业务数据。
- `examples/italy_progress_report/progress_template1_output.pptx`: 基于模板和 mock 数据生成的报告。
- `examples/italy_progress_report/rendered_source/contact_sheet.png`: 源 PPT 预览。
- `examples/italy_progress_report/rendered_output/contact_sheet.png`: 生成结果预览。

源 PPT 共 8 页，16:9 横版。已命名并纳入 mapping 的核心组件共 49 个。

## 当前模板化范围

### 第 1 页：封面

已映射：

- `text.cover_title`: 报告标题。
- `text.cover_period`: 报告周期。

背景图片保持固定，未纳入替换。

### 第 2 页：项目信息概览

已映射：

- `text.project_overview_summary`: 项目概述。
- `text.project_name`: 项目名称。
- `text.project_code`: 项目编码。
- `text.customer_name`: 客户名称。
- `text.project_manager`: 项目经理或项目 PD。
- `text.planned_start_date`: 计划开始时间。
- `text.planned_end_date`: 计划结束时间。
- `text.report_date`: 报告日期。

这页主要是固定字段，当前 `Text` 组件可以覆盖。

### 第 3 页：项目仪表盘

已映射：

- `text.project_status_value`: 项目状态。
- `text.progress_status_value`: 进度状态。
- `text.project_progress_value`: 项目累计进度。
- `text.technical_status_value`: 技术状态。
- `text.resource_status_value`: 资源状态。
- `text.business_status_value`: 经营状态。
- `shape.progress_status_badge`: 进度状态徽标。
- `shape.technical_status_badge`: 技术状态徽标。
- `shape.resource_status_badge`: 资源状态徽标。

当前实现可以替换文本，也可以按 `state_styles` 修改普通 shape 的填充和线条。但它还不理解“仪表盘卡片”“状态徽标”“图标组”这些语义结构，只能把它们拆成独立文本和 shape 处理。

### 第 4 页：项目整体进展

已映射：

- `text.overall_progress_summary`: 项目整体概况。
- `text.key_conclusion`: 关键结论。
- `text.key_progress`: 本期关键进展。
- `text.risk_issue_summary`: 风险及问题管理摘要。

当前 `Text` 组件可以覆盖，但长文本仍受文本框大小限制。

### 第 5 页：项目交付计划

已映射：

- `table.delivery_plan`: 交付计划表格。
- `milestone.delivery`: 项目交付里程碑动态时间轴。

交付计划表格是当前生成器覆盖效果最好的复杂组件：数据列数不超过模板列数时，会优先保留原表格并追加行，尽量保留行高、列宽、边框、填充和字体样式。

时间轴已改为 `Milestone` 组件，节点数量可以随数据动态增减。当前按锚点宽度等距分布节点，不按真实日期比例定位。

### 第 6 页：TOP 问题与风险

已映射：

- `top_issues.cards`: TOP 问题与风险动态卡片列表。

当前已改为 `TopIssues` 组件，卡片数量可以随数据动态增减，内置 `紧急`、`重要`、`一般` 三种样式，并按锚点位置垂直分布。超出页面空间时继续向下排布，不自动分页。

### 第 7 页：下期计划

已映射：

- `text.next_plan_item_1`: 下期计划第一条。
- `text.next_plan_item_2`: 下期计划第二条。

当前仍是固定两条文本。若后续计划数量变化，需要重复项能力或列表文本能力。

### 第 8 页：结束页

结束页保持静态，未纳入替换。

## 当前生成能力覆盖情况

当前项目已经能覆盖这个 PPT 的大部分基础数据替换场景：

- 固定文本替换。
- 多字段项目信息填充。
- 简单状态徽标颜色切换。
- 动态表格增行，并尽量保留原表格样式。
- 动态数量的 TOP 问题与风险卡片。
- 固定数量的下期计划文本替换。

这次生成过程中暴露出一个实际约束：小尺寸日期标签使用 `preserve_style` 时，长日期容易触发 `TEXT_OVERFLOW`。现在时间轴已改为 `Milestone` 程序绘制，仍建议里程碑日期使用短格式以保证视觉稳定。

## 当前缺失或偏弱的生成能力

### 1. 模板组件命名与发现工具

源 PPT 中有大量默认 shape 名称和重复名称，例如 `AutoShape 11`、`AutoShape 14`、`Connector 9`。当前生成器要求 mapping 中的 `location` 与 PPT shape 名完全一致，且同一批组件名称不能重复。

影响：

- 原始 PPT 不能直接稳定生成，必须先手工或脚本重命名关键组件。
- 复杂页面里很难靠人工确认每个 shape 对应哪个视觉元素。

建议：

- 增加模板分析工具，输出 slide、shape name、shape type、文本摘要、位置尺寸和截图编号。
- 增加命名清单或交互式重命名辅助脚本。
- 对未命名或重复命名组件给出更可操作的诊断信息。

### 2. Group shape 递归识别与处理

这个模板里很多视觉元素是组合图形，例如仪表盘卡片、图标、TOP 问题卡片内部元素。当前 `shape_index()` 只遍历 slide 顶层 shapes，不递归进入 group shape。

影响：

- 如果可替换文本或图形位于组合内部，当前 mapping 找不到。
- 需要先在 PowerPoint 中取消组合或把可替换元素暴露到顶层。

建议：

- 支持递归扫描 group shape。
- 设计 group 内组件 location 规则，例如 `group.dashboard.card_1/text.value`。
- 在修改 group 内文本、填充、线条时保留原组合结构。

### 3. 模板组复制模式

第 6 页 TOP 问题与风险已支持 `TopIssues` 动态卡片，但当前卡片是程序绘制，不是复制 PPT 中手工设计好的复杂组合卡片。

影响：

- 如果客户模板里的卡片有复杂组合、阴影、图标或特殊排版，当前不能完全保留原模板卡片细节。
- 同类需求还会出现在“下期计划”“关键进展”等页面。

建议：

- 增加 `Repeater` 或 `GroupList` 组件。
- 支持选择一个模板组作为 item template，根据数组长度复制、布局、填充。
- 支持超出单页时分页，或配置最大显示数量。

### 4. Timeline / Milestone 增强能力

第 5 页时间轴已支持 `Milestone` 动态节点。当前能力仍偏基础。

影响：

- 无法按日期比例重新定位节点。
- 不能表达已完成进度线段、延期标记、阶段跨度等更复杂语义。

建议：

- 输入建议支持 `items: [{label, date, status, progress}]`。
- 支持固定等距模式和按日期比例分布模式。
- 支持状态驱动的节点/线条样式。

### 5. 文本自动适配能力

当前 `Text` 的 `preserve_style` 会严格保留原字号，放不下时直接报 `TEXT_OVERFLOW`；非保留样式模式可以缩字号，但会弱化模板样式。

影响：

- 小标签、标题、卡片正文容易因为业务文本稍长而失败。
- 很难在“默认保留模板样式”和“允许有限自动适配”之间取得平衡。

建议：

- 增加 `preserve_style` 下的可选策略，例如 `fit: shrink`、`fit: wrap`、`fit: truncate`。
- 缩字号时保留原字体、颜色、粗细，只调整字号。
- 支持最大行数和省略号。
- 对中文、英文、数字分别使用更准确的宽度估算。

### 6. 更丰富的状态徽标与进度指示器

当前 `Shape` 只支持写文本、改填充、改线条。仪表盘里的状态徽标、进度标签、色块等可以勉强用 `Shape` 覆盖，但语义较弱。

影响：

- 不能统一表达状态等级、颜色、图标、标签文本。
- 不能根据进度百分比调整进度条长度或仪表盘视觉。

建议：

- 增加 `StatusBadge` 或在 `Shape` 上扩展状态配置。
- 增加 `ProgressBar` / `MetricCard` 类组件。
- 支持按状态映射文本、填充、线条、图标、可见性。

### 7. Jinja 模板与数组上下文易用性

第 6 页 TOP 问题里，严重程度用了 JSONPath `top_issues[0]`，描述/措施/元信息则用了额外的 `top_issue_context.items[0]` 包装对象。

影响：

- mapping 写法不够自然。
- 列表数据在 `template` 场景下需要人为包装成对象，增加 payload 冗余。

建议：

- 允许 `template` 基于任意 JSON 值渲染，并自动提供 `value`、`root`、`item` 等上下文。
- 支持 `data_source.index` 后再套 `template`，例如先取 `$[0]`，再渲染该对象。
- 支持组件级 `repeat_index` 或列表 item 上下文。

### 8. 图片占位符样式保留

当前图片组件会删除原 shape 并插入新图片。这个模板的封面背景是固定图片，所以未触发问题；但如果后续要替换封面图或项目图片，原图片占位符上的边框、填充、阴影不会自动保留。

建议：

- 对图片 shape 支持替换 blip 数据，而不是删除后重插。
- 或约定图片框、边框、遮罩必须拆成独立 shape，并在模板构造说明中明确。

### 9. 幻灯片级动态能力

当前 mapping 面向单个 shape，不支持按数据动态复制、删除或重排幻灯片。

影响：

- TOP 问题超过一页、交付计划表格超过一页、项目阶段较多时，需要人工预置多页。
- 无法根据 payload 自动生成多页明细。

建议：

- 增加 slide-level repeat / conditional show-hide。
- 支持按组件溢出自动分页，例如表格分页、卡片分页。

## 推荐实现顺序

1. 模板分析与命名辅助工具：先降低把真实 PPT 接入系统的成本。
2. Group shape 递归识别：解决复杂商业模板中组件经常被组合的问题。
3. 模板组复制组件：优先覆盖需要完全保留手工设计卡片样式的列表。
4. 文本自动适配增强：降低真实业务数据导致生成失败的概率。
5. Milestone / Timeline 增强：支持按日期比例定位、完成进度线和延期状态。
6. 状态徽标、进度条、指标卡组件：提升仪表盘页的数据表达能力。
7. 更自然的列表/Jinja 上下文：减少 payload 冗余和 mapping 编写成本。
8. 幻灯片级复制与分页：支撑长列表、长表格和多页报告。

## 本次处理的已知取舍

- 为了最大限度保留模板样式，文本组件都使用了 `preserve_style: true`。
- 时间轴日期使用短格式 `MM-DD`，保证动态时间轴视觉稳定。
- TOP 问题页已改为动态卡片，但卡片样式由程序生成，不复制原 PPT 组合卡片。
- 下期计划页按模板固定的 2 条计划填充，没有实现动态列表。
- 封面背景图和结束页保持静态。
