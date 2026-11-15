# 模板继承（用户选择"完全遵循模板"时必读）

设计要求：严格遵循模板。页面比例、母版、版式、主题色、分隔线/色块装饰、页眉页脚页码、图片放置逻辑
一律以模板为准，不另起视觉系统。做法是**复制模板原页并填充**：用 officecli 在模板副本里复制最接近的原页，
只替换文字、图片、表格内容，母版与装饰天然保留。模板原件不会被修改。

## 1. 流程

1. `analyze_template.py <模板.pptx> -o 模板摘要.md` 看版式与主题色，再加 `--slides-json` 看逐页清单
   （页号、版式名、每个形状含组合内子形状的 key/name/id/类型/占位符/位置/文字摘要；组合子形状的位置是组合内坐标）。
2. 按 deck.json 每页的类型与内容量，为它挑一页模板原页。同一模板页可复用多次；模板里没被选中的原页会删除。
3. 写 `map.json`（格式见 §2）。
4. `python template_fill.py --template 模板.pptx --deck deck.json --map map.json --output "<中文名>.pptx"`。
   一次批量编辑，任何一步失败整批回滚，不产出半成品；成功后自动重开校验页数与文字落位。
5. 有公式时：`python formula.py "<中文名>.pptx"`（读取 `<输出名>.formulas.json`，写入原生可编辑公式）。
6. `validate_pptx.py "<中文名>.pptx"` 与 officecli 出图复核（见 validation.md）。

模板模式不跑 `sample_style.py`，也不使用渲染器副本，配色全部来自模板主题。deck.json 仍要先过 `deck_check.py`。

## 2. map.json 格式

`slides` 数组与 deck.json 的 `slides` 一一对应、顺序一致，数量不等直接报错。

```json
{
  "slides": [
    {"template_slide": 1, "fields": {"标题": "封面标题", "副标题": {"text": "汇报人 · 2026.11", "size": 20}}},
    {"template_slide": 2, "fields": {"组合 35/文本框 18": "第一章标题", "组合 35/文本框 23": "第一章说明"}},
    {"template_slide": 3,
     "fields": {"标题": "结论句", "正文": "要点一\n要点二", "图片": {"image": "imgs/a.png"}},
     "remove": ["装饰图"],
     "extra_shapes": [{"text": "图 1 说明", "x": "7.6cm", "y": "13cm", "width": "12cm", "height": "1cm",
                       "size": 12, "color": "tx2"}],
     "formulas": [{"latex": "\\frac{a}{b}", "display": true, "x": "2cm", "y": "8cm", "width": "14cm", "height": "3cm"}]}
  ]
}
```

- `template_slide`：模板页号，从 1 开始。
- `fields`：键是模板页上的形状 key（用 `analyze_template.py --slides-json` 里每个形状的 `key`），值是：
  - 字符串：改文字，`\n` 分段；需要改字号时写 `{"text": "…", "size": 40}`。
  - `{"image": 路径}`：换图，路径相对 map.json；模板图片位的裁剪会被重置，新图的宽高比最好和图片框接近。
  - `{"table": [["表头"],["单元格"]]}`：只改模板里已有表格的单元格。
- `remove`：删掉该页不需要的形状（key 写法同上；填充完成后才删，不会影响 `fields` 的寻址）。
- `extra_shapes`：原页放不下内容时才新增的文本框。`color` 只能用主题色名
  （tx1 tx2 bg1 bg2 dk1 dk2 lt1 lt2 accent1–6），长度用 EMU 整数或 `cm/mm/in/pt`。
- `formulas`：在指定位置留公式占位（形状名 `FORMULA_<序号>`），由 `formula.py` 替换为原生公式。
- 引用不存在的 key、类型不符（给图片位填文字、给组合直接填文字等）、颜色不是主题色，均在动手前报错并列出可用 key，退出码 2。

### key 规则（含组合内的形状）

- 顶层形状：key 就是形状名，如 `标题 1`。
- 组合内的子形状：`组合名/子形状名`，嵌套多层继续用 `/` 连接，如 `组合 35/文本框 18`、`组合 7/组合 9/矩形 4`。
- 同一层有同名形状时，在名字后加 `#序号`（按 z 序从 1 起），如 `组合 70/流程图: 数据 9#3`；以 `analyze_template.py --slides-json` 打印的 key 为准，不要自己拼。
- `@id=N` 写模板页上原来的 id，会被换算成同一个形状；复制页后 officecli 会给所有形状重新编号，所以输出页上的 id 不能用。同页 id 重复时报错，改用 key。
- 组合本身、连接线、图表、艺术字不能填文字；要改组合里的文字，写子形状的 key。

## 3. 选页与填充原则

- 封面、目录、章节、结束页优先用模板里对应的专用页；正文页按"纯文字/图文/多栏"选相近的原页。
- 标题、正文放进模板的占位符，文字样式继承模板，不手动改字号、颜色、位置。
- 原页没有合适的位置时才用 `extra_shapes`；新增文本框只用主题色，不加阴影、发光、渐变。
- 不删除模板的页眉、页脚、页码、logo；这些在母版或原页里，复制时自然保留。
- 每页字数控制在模板占位符能放下的范围内，装不下就换更大的原页，不缩字号硬塞。

## 4. 已知限制

- 没有合适原页时，该页只能用最接近的原页加 `extra_shapes` 拼，版式丰富度不如非模板模式。
- 表格只支持填充模板已有表格的单元格，不新建；图表、SmartArt 不会被改写，保持模板原样。
- 公式字号由 PowerPoint 默认决定，放不下时调大 `formulas` 的 width/height。
- officecli 的 html 出图会多画一个日期和页码占位（PowerPoint 实际不显示）；用 `--render native` 更接近真实效果。
- 交付时说明：哪些模板页被复用、哪些页用了 `extra_shapes` 或回退，方便用户在 PowerPoint 里微调。
