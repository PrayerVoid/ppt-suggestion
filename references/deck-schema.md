# deck.json 结构说明（写 deck.json 与跑 deck_check 前必读）

> 设计要求：页面标题必须是结论句；目录与章节层级清晰、编号前后一致；输出为可编辑 PPT。
> 视觉/配色/密度下限在 `design.md`；版式匹配在 `template-inherit.md`。

## 1. 根对象

```json
{
  "source": "参考资料.md · 2026.11",
  "slides": [ ]
}
```

`source` 写入正文页脚注。无模板时画布 16:9；有模板时由模板决定。deck.json 里没有颜色、主题或样式字段：
配色与基本版式在阶段三写进渲染器副本的顶部常量，不进 JSON。

**通用要求**：`items[]` 里的条目 `body` 都必须是能读出信息的句子（≥10 字）。不要为了填满页面硬凑条目。

## 2. slides[].type


| type      | 用途                                                |
| ----------- | ----------------------------------------------------- |
| `cover`   | 封面                                                |
| `catalog` | 目录，blocks 里放一个 `agenda`                      |
| `section` | 章节扉页（大号编号 + 章名），必须与目录条目一一对应 |
| `content` | 正文（结论句标题 + blocks）                         |
| `close`   | 结束页                                              |

### cover / close

```json
{
  "type": "cover",
  "kicker": "内部技术汇报 · 2026.11",
  "title": "把企业 AI 从项目堆砌做成可治理的生产线",
  "subtitle": "AI Factory 立项评审",
  "footer": "背景 · 现状 · 方案 · 落地"
}
```

`title` 宜短（≤18 字），长标题渲染器会自动缩小。`footer` 用 `·` 分隔 3–5 个短词。

### catalog

```json
{
  "type": "catalog",
  "title": "本场讲四件事",
  "blocks": [
    {
      "kind": "agenda",
      "items": [
        { "num": "01", "title": "先把算力收成统一底座", "note": "Token Factory" }
      ]
    }
  ]
}
```

条目 3–8 条。后面必须有同样数量、标题一致的 `section`。

### section

```json
{
  "type": "section",
  "num": "01",
  "title": "先把算力收成统一底座",
  "subtitle": "Token Factory"
}
```

### content

```json
{
  "type": "content",
  "title": "必须是完整结论句，不要关键词标题",
  "chapter": "03 · TOKEN FACTORY",
  "blocks": []
}
```

`chapter` 显示在右上。`title` ≤ 30 字，尽量 1 行。

## 3. blocks[].kind

同一页可叠 1–3 个 block。优先结构，不要一页塞满。

### lead

```json
{ "kind": "lead", "text": "一句话带出本页判断。" }
```

### insight

本页「所以呢」。浅灰横条 + 强调色左边条，全稿 3–6 条。

```json
{ "kind": "insight", "kicker": "结论", "text": "先建计量，再谈治理。" }
```

### stats

3 或 4 个关键数字。`note` 必写，说明相对谁、意味着什么。

```json
{
  "kind": "stats",
  "items": [
    { "value": "17", "label": "在用大模型接入点", "note": "重复接入，无统一配额" }
  ]
}
```

### agenda

目录页。不要用 columns 做目录。

### columns

2 或 3 列。`icon` 可选，单个字符（如 `01`、`◆`），不传则用序号。

```json
{
  "kind": "columns",
  "items": [
    { "kicker": "现状", "icon": "01", "title": "可选小标题", "body": "80–120 字" }
  ]
}
```

### steps

横向 3–5 步流程。

```json
{
  "kind": "steps",
  "items": [
    { "num": "01", "title": "提交", "body": "携带身份与预算标签发起调用" }
  ]
}
```

### chips

底部能力清单，6–10 个短词。仅当信息确实短时用。

```json
{ "kind": "chips", "items": ["GPU 池", "K8s", "沙箱"] }
```

### layers

架构页专用。每层必须带 `modules`（白底模块盒），禁止把模块名写成 `body` 里的顿号列表。层间自动画箭头。

```json
{
  "kind": "layers",
  "items": [
    {
      "kicker": "应用层",
      "title": "AI Application",
      "modules": [
        { "title": "Coding Agent" },
        { "title": "BI Agent" }
      ]
    },
    {
      "kicker": "生产层",
      "title": "Agent Factory",
      "modules": [
        { "title": "Agent Platform", "body": "生产 Agent" },
        { "title": "Skills Hub", "body": "能力资产中心" }
      ]
    }
  ]
}
```

### table

```json
{
  "kind": "table",
  "headers": ["层", "职责", "不是什么"],
  "rows": [["AI Infra", "计算即服务", "GPU 台账"]]
}
```

### split

左右不对称。`right.kind`：`list`（白底条目）或 `panel`（浅灰列表）。

```json
{
  "kind": "split",
  "left": { "kicker": "TOKEN LIFECYCLE", "title": "Token\n= 新的 CPU", "body": "管理 Token，比管理模型更重要。" },
  "right": { "kind": "panel", "kicker": "生命周期", "items": ["购买额度", "Quota", "Chargeback"] }
}
```

### callouts

纵向 2 条说明，适合补治理/评测。

```json
{
  "kind": "callouts",
  "items": [
    { "title": "Auto Benchmark", "body": "每天自动跑企业 Benchmark。" }
  ]
}
```

### cards

2×2 / 2×4 / 3×3 网格。`items` 数量必须能整行铺满（4/6/8/9），禁止 5 个留空。全稿 ≤ 2 页。

```json
{
  "kind": "cards",
  "columns": 3,
  "items": [{ "title": "Authentication", "body": "可选一句" }]
}
```

### rows

端到端泳道，自上而下。

```json
{
  "kind": "rows",
  "items": [
    { "tag": "应用", "text": "Coding Agent 发起任务" }
  ]
}
```

### compare

现状 vs 目标 / 不是什么 vs 是什么。左栏浅灰、右栏强调色高亮，条目自动带 × 和 ✓。两侧 `items` 各 3–5 条。

```json
{
  "kind": "compare",
  "left": {
    "kicker": "今天的做法",
    "title": "先建工具，再等业务来用",
    "items": ["每个业务线各自采购模型，额度无法互通", "Agent 散落在十几个仓库"]
  },
  "right": {
    "kicker": "工厂思维",
    "title": "先把生产要素标准化",
    "items": ["Token 统一计量，按团队做 Chargeback", "能力进 Skills Hub，一次生产多处复用"]
  }
}
```

### pipeline

端到端环节链，3–5 段，段间自动画箭头。`metric` 可选，显示在卡片底部。

```json
{
  "kind": "pipeline",
  "items": [
    { "kicker": "STAGE 01", "title": "任务提交", "body": "Agent 携带身份与预算标签发起调用", "metric": "P99 120ms" }
  ]
}
```

### hub

中枢辐射图。用于统一入口 / 控制平面 / 网关。`items` 3–6 条，左右交替排布。

```json
{
  "kind": "hub",
  "center": { "kicker": "CONTROL PLANE", "title": "Agent Gateway", "body": "所有调用的强制通道" },
  "items": [
    { "title": "身份与鉴权", "body": "统一接入企业账号体系" },
    { "title": "配额与限流", "body": "按团队分配预算并实时熔断" }
  ]
}
```

### timeline

分期路线图，3–5 期。`phase` 写阶段与时间，`metric` 可选，写验收口径。

```json
{
  "kind": "timeline",
  "items": [
    { "phase": "P1 · 1-2 月", "title": "收口流量", "body": "网关上线，存量调用完成迁移", "metric": "覆盖率 100%" }
  ]
}
```

### points

编号论点列表，3–5 条。比 `callouts` 更密，适合并列论证。

```json
{
  "kind": "points",
  "items": [
    { "title": "先计量再治理", "body": "没有台账，配额与成本归因都无法落地。" }
  ]
}
```

### formula

数学公式。`latex` 必填，直接写 LaTeX（`\frac{a}{b}`、`\sum_{i=1}^{n}`、`\sqrt{x}`、希腊字母等）；
`display` 为 `true` 时独占一行居中，`false` 为行内较小的一行；`caption` 可选，写公式说明。

渲染器先画一个名为 `FORMULA_<序号>` 的占位文本框（文字是 LaTeX 源码），并写出 `<输出名>.formulas.json`；
渲染后跑 `scripts/formula.py <输出.pptx>`，用 officecli 在原位置写入 PowerPoint 原生可编辑公式（OMML）并删除占位框。
写入失败整批回滚，不会降级成图片。公式页同样要满足正文字数与密度要求，公式前后写清变量含义与结论。

```json
{ "kind": "formula", "latex": "E = mc^2", "display": true, "caption": "质能方程" }
```

#### 行内公式 `$...$`

任意文本字段（段落、卡片、表格单元格、列表项等）里，数学符号用 `$LaTeX$` 标记，例如 `第 $l$ 层的神经元数 $n_l$`、`衰减系数 $\alpha_{t}$`。字面美元符号写成 `\$`。
渲染器照原样写入文字，`scripts/formula.py` 在渲染后统一把 `$...$` 换成 PowerPoint 原生可编辑的行内公式
（字号与颜色沿用所在文字，同一段里被拆成多个 run 的 `$...$` 也会合并处理）。`$` 必须成对，否则 `deck_check.py` 报错。
只含公式的一整段文字，PowerPoint 会按独占一行的样式排版（居中、上下标更大）；要行内效果就把公式放进一句话里。
`latex` 字段（`formula` block）不用 `$` 包裹。

### figure

引用资料中的原图，等比缩放放入，不拉伸。仅当用户在询问阶段允许复用资料图片时使用。`src` 相对 deck.json 所在目录；`side` 可选，是图旁的文字栏
（`kicker` / `title` / `items` 字符串列表）；`image_side` 为 `left` 或 `right`（默认 right）；
`image_ratio` 图占宽度比例（有 `side` 时默认 0.62）；`height` 图的最大高度（英寸），同页还有别的
block 时必须设置。图宽过小的截图（如 400px）不要放太大，否则会糊。

```json
{
  "kind": "figure",
  "src": "images/img2",
  "caption": "图：ECIES 原理流程（参考资料）",
  "image_side": "left",
  "image_ratio": 0.5,
  "height": 4.3,
  "side": { "kicker": "火山 · 原理", "title": "一句话标题", "items": ["要点一", "要点二"] }
}
```

### chart

图表能力默认开启。没有资料中的真实数据、来源或可核对口径时，不生成图表，不凭空补数值、分类、趋势或结论。优先使用 `power_pptx` 生成 PowerPoint 原生图表；类型不支持或 Office 兼容性不稳定时，明确写 `fallback: "svg"`，SVG 兜底依赖 `matplotlib` 与 `cairosvg`。

`fallback` 可写 `auto`（默认，原生失败时转 SVG）、`native`（强制原生）或 `svg`（强制 SVG）。分类图表使用 `categories` 与 `series[].values`，每个系列的数值数量必须与分类数量一致；`xy_scatter`、`xy_scatter_lines`、`bubble` 使用 `series[].points`，点分别需要 `x`、`y`，气泡图还需要正数 `size`。系列名称必须唯一。

```json
{
  "kind": "chart",
  "title": "来自资料的图表标题",
  "chart_type": "column_clustered",
  "categories": ["分类一", "分类二"],
  "series": [{ "name": "系列名称", "values": ["<来自资料的数值一>", "<来自资料的数值二>"] }],
  "unit": "可选单位",
  "source": "资料来源或数据表",
  "accessibility_note": "可选的口径与读图说明",
  "fallback": "auto"
}
```

上例只展示字段结构，数值必须替换为资料中的真实数据。每个图表必须提供 `source`、`table_view` 或 `accessibility_note` 至少一项，说明来源、口径、时间范围或读图方式；多系列图表尤其不能省略。禁止使用 `secondary_axis`；需要双轴时改用同一坐标口径、拆分图表或文字说明。原生白名单之外的类型必须使用 `fallback: "svg"`。支持的原生类型见 `extensions/charts/EXT.md`。

模板模式按四种模式区分图表职责：完全遵循模板可以保留已有图表，也可以用 OfficeCLI 将顶层图片/图表位替换成同尺寸 native chart，或在明确位置新增 chart；半自主可以从内容页母页扩展 chart，并使用模板主题色、字体和版式骨架；两个完全自主模式继续由 `power_pptx` 原生渲染，必要时使用 SVG fallback。复杂嵌套图表或系列结构大幅变化时必须明确回退，不得让图表静默丢失。

## 4. 质量检查清单（写 JSON 前过一遍，`deck_check.py` 会强制执行）

- 每页 `content.title` 是结论，不是模块名（≥8 字）
- 有 `section` 必须有 `catalog`，且 `section` 数量/标题与目录条目一致
- 全稿至少 **2 页**结构页（`layers` / `steps` / `rows` / `pipeline` / `hub` / `timeline` / `compare`）
- 全稿至少用到 **5 种**不同版式（`lead` / `insight` / `chips` 不计）
- 每页正文总字数 ≥ 110 字；`cards` / `columns` / `pipeline` / `points` / `callouts` / `timeline` 的每条 `body` ≥ 10 字
- 没有只放 `lead` / `insight` / `chips` 的页
- 等宽 `cards` 全稿 ≤ 2 页；禁止连续两页 `cards`/`columns`
- 不要给 block 加颜色字段（color / fontSize / hex 一律禁止）
- `formula` 的 `latex` 必填且非空
- `chart` 只能使用真实数据；必须有 `source`、`table_view` 或 `accessibility_note` 至少一项，禁止 `secondary_axis`
- 文本里的行内公式 `$...$` 的 `$` 必须成对，字面美元符号写成 `\$`
- 模板模式（`deck_check.py --template-mode`）：版式由模板决定，上面关于字数、版式种类、结构页、`cards` 与相邻页版式的检查全部跳过；
  仍检查目录与章节一致、标题 ≥8 字、禁止视觉字段、block kind 合法、`formula` 的 `latex`、`$` 配对
- 启用了特殊排版拓展时，`bignum` / `asym` / `quote` / `diagonal` / `radial` / `ring` / `marginalia` / `concentric` / `quadrant` 等新增 kind 要用 `deck_check.py --extra-kinds` 放行（字段见 `extensions/special-layout/EXT.md`）

## 5. deck_check.py 错误解读

`deck_check.py deck.json`（阶段三复制到项目里的渲染器副本里那份）失败会非 0 退出并打印每条原因，格式 `✗ <原因>`。按表修复后重跑：


| 报错                                       | 修法                                 |
| -------------------------------------------- | -------------------------------------- |
| 目录 N 条与 section M 个数量不一致         | 补/删 section 或目录条目，使两者相等 |
| 目录第 i 条「A」与 section「B」标题不一致  | 改标题使逐字一致                     |
| 第 N 页正文总字数 X < 110                  | 补内容（机制/代价/边界），别用空话凑 |
| 第 N 页「...」的 body 只有 X 字            | 把标签式短语扩成能读出信息的句子     |
| 第 N 页标题过短（X 字），疑似关键词标题    | 改成完整结论句                       |
| 第 N 页只放了 lead/insight/chips，没有结构 | 换结构版式或合并进其他页             |
| 第 N 页与上一页连续使用 cards/columns      | 换相邻页版式                         |
| cards 页共 N 页 > 2                        | 减少卡片页                           |
| 全稿版式种类过少                           | 增加结构版式种类                     |
| block 携带视觉字段 color/fontSize/hex      | 删除该字段（颜色由渲染器副本的常量控制） |
| 未知的 block kind                          | 改成已有 kind，或用 `--extra-kinds` 放行渲染器副本里新增的 kind |
| formula 缺少 latex                         | 补上 LaTeX 源码                      |
| 第 N 页 chart 缺少来源说明 | 补 `source`、`table_view` 或 `accessibility_note`，说明数据来源、口径或读图方式 |
| 第 N 页 chart 使用 `secondary_axis` | 删除该字段，改用同一坐标口径、拆分图表或文字说明 |
| 第 N 页 … 的 $ 数量为奇数 / 含空的行内公式 $$ | 补全 `$...$`；字面美元符号写成 `\$`  |
