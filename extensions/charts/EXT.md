# 图表能力

图表能力默认开启，使用 `chart` block 时无需另行启用拓展。图表只表达资料中的真实数据；没有真实数据、来源或可核对口径时，不生成图表，也不凭空补数值、分类、趋势或结论。

## 使用规则

- 优先使用 `power_pptx` 创建 PowerPoint 原生图表，保留后续编辑能力。
- 图表类型不受原生支持，或 Office 兼容性不稳定时，使用 `fallback: "svg"`。SVG 兜底依赖 `matplotlib` 与 `cairosvg`，已列入 `requirements.txt`。
- `fallback` 可写 `auto`（默认，原生失败时转 SVG）、`native`（强制原生）或 `svg`（强制 SVG）。
- 每个图表必须提供 `source`、`table_view` 或 `accessibility_note` 至少一项，说明数据来源、口径、时间范围或读图方式；多系列图表尤其不能省略。
- 禁止使用 `secondary_axis`。需要双轴时改用同一坐标口径、拆成两个图表，或改用文字说明。
- 模板模式只保留模板中已有的图表，不新建图表，也不改写已有图表的系列；图表系列变更需在模板外完成并重新确认模板适配。

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