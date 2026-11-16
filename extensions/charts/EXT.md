# 图表能力

图表能力默认开启，使用 `chart` block 时无需另行启用拓展。图表只表达资料中的真实数据；没有真实数据、来源或可核对口径时，不生成图表，也不凭空补数值、分类、趋势或结论。

## 使用规则

- 无模板的两个完全自主模式优先使用 `power_pptx` 创建 PowerPoint 原生图表，保留后续编辑能力。
- 完全遵循模板和半自主模式使用 OfficeCLI 管理模板页：保留模板已有图表；`template_fill.py` 可将顶层图片位或已有图表位同位替换为 native chart。任意坐标新增图表需另写明确的 OfficeCLI 操作，当前 map 不支持。配色优先继承模板实际可见颜色；先核对主题色与可见颜色是否一致，不能只因名为 `accent1` 就直接采用。
- OfficeCLI 的模板图表操作应放在一次 `batch` 中；复制页、删除原图片、添加 chart 和设置位置任一步失败，整批回滚，不保留半成品。
- 图表类型不受原生支持，或 Office 兼容性不稳定时，使用 `fallback: "svg"`。SVG 兜底依赖 `matplotlib` 与 `cairosvg`，已列入 `requirements.txt`。
- 完全自主渲染器的 `fallback` 可写 `auto`（默认，原生失败时转 SVG）、`native`（强制原生）或 `svg`（强制 SVG）。模板脚本当前只生成 native chart，不实现自动 SVG 回退；需要图片回退时生成 SVG/PNG，再通过 `fields.image` 替换顶层图片位，并记录可编辑性限制。模板输出必须确认 chart 对象实际存在，失败不得静默通过。
- 单个测量值没有比较关系，不生成柱状图、折线图或竞品比较图；改用 `stats`/`lead`，或补齐同口径、可比较的类别。`deck_check.py` 和 `template_fill.py` 都会在只有一个 category 时阻断图表。
- 每个图表必须提供 `source`、`table_view` 或 `accessibility_note` 至少一项，说明数据来源、口径、时间范围或读图方式；多系列图表尤其不能省略。
- 禁止使用 `secondary_axis`。需要双轴时改用同一坐标口径、拆成两个图表，或改用文字说明。
- OfficeCLI 当前最稳的是同尺寸新增 chart。模板已有原生图表需要改变系列结构或类型时，优先保留原样；若必须变化，使用同尺寸 SVG/PNG 替换并在校验报告记录失去 native editability 的回退。复杂嵌套 group 内的图片位不能直接 `add --from` 为页面根部 chart，应复制整组或选择顶层图片位。

## OfficeCLI 模板图表参数

模板模式的 `template_fill.py` 将 map 中的 chart 对象转换为 OfficeCLI 的 `add --type chart`，使用这些属性：

```text
chartType    column / stackedColumn / bar / line / area / pie / doughnut / radar
categories   逗号分隔的分类
data         系列名:数值,数值;另一系列:数值,数值
x y w h      沿用被替换图片或图表位的 EMU 坐标
colors       模板主题色名，如 accent1,accent2,accent3
```

也可以在 chart 对象中传入 `data_labels`、`label_pos`、`axis_font`、`gridlines` 和 `chart_area_fill`，脚本会映射到 OfficeCLI 的 `dataLabels`、`labelPos`、`axisfont`、`gridlines` 和 `chartareafill`。系列名不能含 `:` 或 `;`，分类不能含 `,` 或 `;`，以避免 OfficeCLI 的批量数据语法歧义。脚本在同一原子 batch 内添加同位图表并删除原图片，避免删除导致位置寻址前移；输出只保留新 chart。嵌套 group 内图片位需要使用整组复制或另写明确的图片回退流程。

## block 结构

```json
{
  "kind": "chart",
  "title": "来自资料的图表标题",
  "subtitle": "可选的口径或范围说明",
  "chart_type": "column_clustered",
  "categories": ["分类一", "分类二"],
  "series": [
    {"name": "系列名称", "values": ["<来自资料的数值一>", "<来自资料的数值二>"]}
  ],
  "unit": "可选单位",
  "source": "资料来源或数据表",
  "accessibility_note": "可选的读图与口径说明",
  "fallback": "auto",
  "note": "可选的页面说明"
}
```

上例只展示字段结构；实际 `categories`、`values` 和 `points` 必须替换为资料中的真实数据。分类图表的每个 `values` 长度必须与 `categories` 一致，系列名称必须唯一。

散点图、带线散点图和气泡图使用 `points`：

```json
{
  "kind": "chart",
  "chart_type": "xy_scatter",
  "series": [
    {"name": "系列名称", "points": [{"x": "<来自资料的真实数值>", "y": "<来自资料的真实数值>"}]}
  ],
  "source": "资料来源"
}
```

气泡图的每个点还必须提供正数 `size`。支持的原生 `chart_type` 为：

`column_clustered`、`column_stacked`、`column_stacked_100`、`bar_clustered`、`bar_stacked`、`line`、`line_markers`、`area`、`area_stacked`、`pie`、`doughnut`、`radar`、`radar_filled`、`xy_scatter`、`xy_scatter_lines`、`bubble`。

原生白名单之外的类型必须明确写 `fallback: "svg"`，并仍提供真实数据与来源说明。