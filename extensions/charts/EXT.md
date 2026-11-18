# 图表能力

图表能力默认可用，使用 `chart` block 时无需另行启用扩展。创建图表或 KPI 前加载 dataviz skill；视觉继承按 [design.md](../../references/design.md)。图表只表达资料中的真实数据；没有真实数据、来源或可核对口径时，不生成图表，也不凭空补数值、分类、趋势或结论。

## 使用规则

- 无模板的两个完全自主模式优先使用 `power_pptx` 创建 PowerPoint 原生图表，保留后续编辑能力。
- strict 保留模板已有图表；显式 custom 模式的 `template_fill.py` 可将顶层图片位或已有图表位同位替换为 native chart。半自主 canvas 复用核心渲染器，支持 chart、figure 和 formula block，并按安全区原生坐标绘制。任意坐标新增图表需另写明确的 OfficeCLI 操作，当前 map 不支持。配色优先继承模板实际可见颜色；先核对主题色与可见颜色是否一致，不能只因名为 `accent1` 就直接采用。
- 包级克隆先准备独立页；OfficeCLI 的删除原图片、添加 chart 和设置位置在一次原子 batch 中完成。任一步或最终校验失败，暂存输出不发布。
- 图表类型不受原生支持，或 Office 兼容性不稳定时，使用 `fallback: "svg"`。SVG 兜底依赖 `matplotlib` 与 `cairosvg`，列于 `requirements-optional.txt`，由 `requirements.txt` 汇总。
- 完全自主渲染器的 `fallback` 可写 `auto`（默认，原生失败时转 SVG）、`native`（强制原生）或 `svg`（强制 SVG）。模板脚本当前只生成 native chart，不实现自动 SVG 回退；需要图片回退时生成 SVG/PNG，再通过 `fields.image` 替换顶层图片位，并记录可编辑性限制。模板输出必须确认 chart 对象实际存在，失败不得静默通过。
- 单个测量值没有比较关系，不生成柱状图、折线图或竞品比较图；改用 `stats`/`lead`，或补齐同口径、可比较的类别。`deck_check.py` 和 `template_fill.py` 都会在只有一个 category 时阻断图表。
- 图表数据与来源字段的契约见 [schema](../../references/deck-schema.md#chart)。
- 禁止使用 `secondary_axis`。需要双轴时改用同一坐标口径、拆成两个图表，或改用文字说明。
- OfficeCLI 当前最稳的是同尺寸新增 chart。模板已有原生图表需要改变系列结构或类型时，先核对上文的同位原生替换范围；不支持的变化使用明确的 SVG/PNG 图片替换流程，并在校验报告记录可编辑性限制。复杂嵌套 group 内的图片位不能直接 `add --from` 为页面根部 chart，应复制整组或选择顶层图片位。

## OfficeCLI 模板图表参数

custom 模式的 `template_fill.py` 将 map 中的 chart 对象转换为 OfficeCLI 的 `add --type chart`，使用这些属性：

```text
chartType    column / stackedColumn / bar / line / area / pie / doughnut / radar
categories   逗号分隔的分类
data         系列名:数值,数值;另一系列:数值,数值
x y w h      沿用被替换图片或图表位的 EMU 坐标
colors       模板主题色名，如 accent1,accent2,accent3
```

也可以在 chart 对象中传入 `data_labels`、`label_pos`、`axis_font`、`gridlines` 和 `chart_area_fill`，脚本会映射到 OfficeCLI 的 `dataLabels`、`labelPos`、`axisfont`、`gridlines` 和 `chartareafill`。系列名不能含 `:` 或 `;`，分类不能含 `,` 或 `;`，以避免 OfficeCLI 的批量数据语法歧义。脚本在同一原子 batch 内添加同位图表并删除原图片，避免删除导致位置寻址前移；输出只保留新 chart。嵌套 group 内图片位需要使用整组复制或另写明确的图片回退流程。

## deck 字段与原生类型

`chart` block 的 categories/series、points、fallback 及来源字段统一见 [deck-schema.md](../../references/deck-schema.md#chart)，此处不重复定义。完全自主渲染器支持的原生 chart_type 为：

`column_clustered`、`column_stacked`、`column_stacked_100`、`bar_clustered`、`bar_stacked`、`line`、`line_markers`、`area`、`area_stacked`、`pie`、`doughnut`、`radar`、`radar_filled`、`xy_scatter`、`xy_scatter_lines`、`bubble`。

原生白名单之外的类型须明确 `fallback: "svg"`，且图像后端本身也要支持该类型；不能把任意名称视为可渲染。模板图表仅接受上文列出的 OfficeCLI 映射，不等同于此原生白名单。
