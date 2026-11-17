# 扩展：特殊排版

仅在用户启用特殊排版时接入 `diagonal`、`radial`、`ring`、`concentric`。日常优先核心版式，选型和图片构图见 [layout-types.md](../../references/layout-types.md)，核心字段见 [deck-schema.md](../../references/deck-schema.md)。本文件负责四个扩展 kind 的接入、字段与容量，不重新定义视觉和密度规则。

## 接入步骤

1. 将 `scripts/render_pptx.py`、`scripts/text_measure.py`、`scripts/deck_check.py` 复制到项目的 `渲染器/`。
2. 将本目录 [blocks.py](blocks.py) 的全部函数放入副本的 `BLOCK_RENDERERS` 之前；依赖副本已有的 `math`、常量、形状工具和 `_layout_items`。仅追加以下四项注册，不重复登记同名函数：

```python
"diagonal": _diagonal,
"radial": _radial,
"ring": _ring,
"concentric": _concentric,
```

3. 使用 skill 虚拟环境执行（`<py>` 见 SKILL）：

```text
<py> 渲染器/deck_check.py deck.json --extra-kinds diagonal,radial,ring,concentric
<py> 渲染器/render_pptx.py --input deck.json --output 输出.pptx
```

`--extra-kinds` 只放行名称，不安装渲染函数或校验所有条目容量；渲染器遇到超量条目会报错，不能静默截断。模板模式按 map 填充已有模板位，不因启用本扩展而自动渲染自由坐标构图。

## Block 选择

| kind | 信息关系 / 字段 | 容量与边界 |
|---|---|---|
| `diagonal` | 方向主张；`kicker/title/body/note/side` | `side` 为 left/right；浅色斜切，无图片蒙版；文字过多就拆页 |
| `radial` | 中心与短节点；`center/items` | 3–6 节点，真实组成或依赖；日常优先 `hub` |
| `ring` | 单指标；`value/kicker/title/body/items` | 真实值，旁注最多 3 条；固定完整圆环不表示完成比例；不是数据饼图 |
| `concentric` | 嵌套层级；`center/rings` | 2–3 层，由外到内；环编号与右侧同序，圆大小不表示数值 |

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

适合一个真实指标需要被快速读取、同时保留少量解释的页面。它是固定圆环衬托数字的版式，不编码完成比例；有多个分类、时间序列或精确占比时应使用原生图表。以下仅展示字段；数值、判断和来源必须替换为资料中的真实信息。

```json
{
  "kind": "ring",
  "value": "<原始数值及单位>",
  "kicker": "关键指标",
  "title": "资料支持的指标判断",
  "body": "填写统计口径、时间范围与指标含义。",
  "items": [
    { "label": "来源", "text": "<文件名、数据表或页码>" }
  ]
}
```

### concentric

适合表达外部环境、能力层与核心对象的嵌套关系。`rings` 按从外到内排列，环上编号与右侧说明一一对应，`center.title` 只写共同的核心对象短标签；解释写在右侧。圆环大小固定，不编码数值；需要精确比例时使用原生图表，不把同心圆当作进度图。

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

## 使用纪律

- 一页通常一个特殊 block；相邻页不机械重复。数字、引文、关系均须来自资料，不把字段示例当作项目事实。
- 不给 JSON 添加 color/fontSize/hex 等视觉字段，配色来自渲染器副本常量。
- 环绕关系不表示图片等权；同心层级不表示进度，指标圆环不表示精确扇区比例。
- 必须运行 validate_pptx 并逐页预览，重点核对斜切面与文字、放射节点、圆环旁注、层级对应和页脚；脚本通过不能覆盖目视问题。

后处理顺序与最终检查按 [扩展索引](../README.md) 和 [validation.md](../../references/validation.md)；本扩展不自动启用图片裁剪、透视、3D 或 Morph。
