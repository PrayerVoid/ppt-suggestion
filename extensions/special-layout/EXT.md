# 拓展：特殊排版

仅在用户选了“特殊排版”时使用。它给内容页增加少量有明确传播目的的版式 block，不修改 skill 原件里的 `scripts/render_pptx.py`；运行时只把代码粘进阶段三复制出的渲染器副本。

当前 block 按信息结构覆盖核心数字、非对称论点、引文、方向主张、中心关系、指标强调、主文边注、同心层级和四象限判断；内置 `columns`、`compare`、`timeline`、`hub`、`steps` 已能表达的页面不重复启用。

已有的 `design.md`、`deck-schema.md`、`validation.md` 优先级更高：特殊排版不能绕过白底极简、单套强调色、正文密度、相邻页版式不同、标题必须是结论句、文字不重叠和资料不编造等规则。

## 接入步骤

1. 打开 `extensions/special-layout/blocks.py`，把 `_point_texts` 到文件末尾的函数整体粘进**渲染器副本**，位置在 `BLOCK_RENDERERS = {` 之前。文件头列出的依赖在副本里都已有，不要重复 import。
2. 在副本的 `BLOCK_RENDERERS` 中登记：

   ```python
   "bignum": _bignum,
   "asym": _asym,
   "quote": _quote,
   "diagonal": _diagonal,
   "radial": _radial,
   "ring": _ring,
   "marginalia": _marginalia,
   "concentric": _concentric,
   "quadrant": _quadrant,
   ```

3. 渲染前校验时放行新增 kind：

   ```bash
   python deck_check.py deck.json --extra-kinds bignum,asym,quote,diagonal,radial,ring,marginalia,concentric,quadrant
   ```

4. 仍按正常顺序执行：

   ```text
   deck_check → 渲染器/render_pptx.py → formula.py（如有公式）→ geometry-deco（如启用）→ animation（如启用）→ validate_pptx → render_preview
   ```

## Block 选择

| kind | 适用信息 | 主要字段 | 约束 |
|---|---|---|---|
| `bignum` | 一个来自资料的核心数字 | `value`、`title`、`body`、`items` | 全稿最多 2 页；`items` 最多 3 条 |
| `asym` | 主论点 + 右侧依据 | `left`、`right` | 固定约 7:3；左侧正文不超过 4 行，右侧最多 4 条 |
| `quote` | 原文引语、用户明确要求保留的观点 | `text`、`source`、`mark` | 必须有可追溯来源或明确的说话者；不能把普通结论伪装成引语 |
| `diagonal` | 章节转场、方案/结果、封面式主张 | `kicker`、`title`、`body`、`note`、`side` | 只用浅色强调面做斜切方向提示；不创建深色整页背景 |
| `radial` | 一个中心能力与 3–6 个关系节点 | `center`、`items` | 节点必须是中心的真实组成或依赖关系，不用来散放装饰词 |
| `ring` | 一个比例/完成度/覆盖率指标 | `value`、`title`、`body`、`items` | `value` 必须来自资料；最多 3 条旁注，不替代完整图表 |
| `marginalia` | 主判断 + 定义、边界或补充证据 | `main`、`notes` | `notes` 最多 3 条；边注必须能回指主文判断 |
| `concentric` | 外部环境、能力层与核心对象的嵌套关系 | `center`、`rings` | `rings` 2–3 层；环内只放短标签，不塞长文本 |
| `quadrant` | 四类互斥判断、风险矩阵或 SWOT 类信息 | `center`、`items` | `items` 必须正好 4 条；每项短标题 + 一句说明 |

### bignum

适合真正需要观众先记住的数字。整页以一个数字为视觉锚点，右侧最多 3 条解释；数字不能由版式自行推算或编造。

```json
{
  "kind": "bignum",
  "kicker": "核心结论",
  "value": "37%",
  "title": "单次调用成本下降",
  "body": "对比统一计量前后三个月的账单，按同一口径折算。",
  "items": ["重复采购的卡数减少", "闲置配额重新入池", "业务开始按调用量优化"]
}
```

### asym

适合左侧形成完整判断、右侧列出证据或边界的页面。左侧 `title` 是结论句，右侧 `items` 每条只表达一个依据。

```json
{
  "kind": "asym",
  "left": {
    "kicker": "主论点",
    "title": "先建计量，再谈治理",
    "body": "没有统一台账，配额、成本和责任都无法落到团队。"
  },
  "right": {
    "kicker": "依据",
    "items": ["三个月账单完成对账", "九个业务接入调用日志"]
  }
}
```

### quote

只用于真正的引语或需要原样保留的观点。`source` 写说话者、文件或页码；没有来源时应改用 `lead`、`insight` 或普通结构 block。

```json
{
  "kind": "quote",
  "text": "排版的核心不是装饰，而是让信息的主次、分组与先后顺序一眼读懂。",
  "source": "版式原则 · 内部设计规范"
}
```

### diagonal

适合章节转场、方案/结果页或需要明显方向感的主张页。`side` 只能为 `left` 或 `right`，默认 `right`；斜切面使用渲染器的浅色强调色，不在 JSON 里写颜色。

```json
{
  "kind": "diagonal",
  "side": "right",
  "kicker": "方案 / 结果",
  "title": "把路线压缩成一条可执行主线",
  "body": "先收口入口，再统一计量，最后把能力沉淀为可复用资产。",
  "note": "页面只保留一个主张"
}
```

### radial

适合控制平面、统一入口、能力中心等中心—周边关系。中心对象放在 `center`，周边节点放在 `items`；每个节点最多一个标题和一小句说明。

```json
{
  "kind": "radial",
  "center": { "kicker": "CONTROL PLANE", "title": "统一入口" },
  "items": [
    { "title": "身份", "body": "统一账号与权限" },
    { "title": "配额", "body": "按团队分配预算" },
    { "title": "审计", "body": "保留调用证据" },
    { "title": "计费", "body": "形成成本归因" }
  ]
}
```

### ring

适合一个比例指标需要被快速读取、同时保留少量解释的页面。它是指标强调版式，不是完整饼图；有多个分类、时间序列或精确占比时应使用原生图表。

```json
{
  "kind": "ring",
  "value": "68%",
  "kicker": "覆盖率",
  "title": "统一入口已覆盖主要调用",
  "body": "覆盖率按已迁移业务数计算，仍有边缘系统待接入。",
  "items": [
    { "label": "已覆盖", "text": "核心业务" },
    { "label": "待接入", "text": "边缘系统" }
  ]
}
```

### marginalia

适合研究说明、定义、边界条件和补充证据。主文区承载完整判断，右侧窄栏只保留能回指主文的边注。

```json
{
  "kind": "marginalia",
  "main": {
    "kicker": "定义",
    "title": "治理边界必须先于工具选择明确",
    "body": "主文区承载完整判断，右侧只保留补充说明与边界条件。"
  },
  "notes": [
    { "label": "边界 01", "text": "不把供应商能力等同于组织治理能力。" },
    { "label": "边界 02", "text": "注释必须能回指主文中的具体判断。" }
  ]
}
```

### concentric

适合表达外部环境、能力层与核心对象的嵌套关系。环内只放短标签，解释写在右侧；需要精确比例时使用原生图表，不把同心圆当作进度图。

```json
{
  "kind": "concentric",
  "center": { "kicker": "层级关系", "title": "核心" },
  "rings": [
    { "label": "外部环境", "text": "政策、客户与供应约束" },
    { "label": "能力层", "text": "平台规则与可复用能力" },
    { "label": "核心对象", "text": "团队可直接调用的资产" }
  ]
}
```

### quadrant

适合四类互斥判断、风险矩阵或 SWOT 类信息。中心标题只承担分类主题，四个象限分别写一条判断，不用于替代 `compare` 的现状与目标两栏。

```json
{
  "kind": "quadrant",
  "center": { "title": "治理" },
  "items": [
    { "key": "01", "title": "高价值", "body": "优先投入并持续复盘" },
    { "key": "02", "title": "高风险", "body": "先补审计与责任边界" },
    { "key": "03", "title": "低成本", "body": "沉淀为标准能力" },
    { "key": "04", "title": "待观察", "body": "保留试点并收集证据" }
  ]
}
```

## 使用纪律

- 一页通常只放一个特殊 block；`bignum` 允许再配一个 `insight`，其余组合须有明确的信息关系。
- 特殊 block 不能连续机械重复；相邻页仍需满足版式不同规则。
- `quote`、`radial`、`ring`、`marginalia`、`concentric`、`quadrant` 都必须服务于事实或关系表达，不能用空话填满页面。
- 不给 block 增加 `color`、`fontSize`、`hex` 等视觉字段；配色只改渲染器副本顶部常量。
- 先用现有内置 block 能清楚表达时，不启用特殊 block；拓展只解决现有版式表达不了的页面。
- 运行 `validate_pptx.py` 后必须用 `render_preview.py` 逐页查看，重点检查斜切面是否压住文字、环形指标是否被页脚截断、放射节点是否相互碰撞、边注与四象限文字是否过密。
