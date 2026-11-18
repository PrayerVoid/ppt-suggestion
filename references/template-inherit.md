# 模板继承与 map 接口

本文件负责 `template_fill.py` 的 strict/custom 边界、选页、寻址与字段。半自主正文画布使用 [template-semi-autonomous.md](template-semi-autonomous.md)，不走此接口。两种模板路径均继承原模板的比例、主题、母版和字体配色，不运行采样器。半自主在清理后的安全区复用核心正文渲染器，不使用独立采样画布。

## 1. 完全遵循：strict

map 默认 `"mode": "strict"`。复制容量合适的原页，替换已有文字、图片或表格内容，保持全部对象数量、位置、尺寸、分组、连接线、编号和底衬。每个语义槽位都应填入真实的论据、机制、条件或边界；材料不足时换页或拆页。

strict 禁止 `remove`、`clone_groups`、`add_texts`、`extra_shapes`、`formulas` 和 `fields.chart`。文字可显式调整字号或主题色以修正模板缺陷，但不得靠持续缩字号挤入内容。对象结构的修改使用 §4 的 custom 接口，并准确标注模式。

## 2. 执行流程

命令从 skill 根目录运行，`<py>` 为项目虚拟环境解释器。

```text
<py> scripts/analyze_template.py 模板.pptx -o 模板摘要.md
<py> scripts/analyze_template.py 模板.pptx --slides-json > 模板清单.json
<py> scripts/deck_check.py deck.json --template-mode
<py> scripts/template_preflight.py --template 模板.pptx --map map.json --output preflight.json
<py> scripts/template_fill.py --template 模板.pptx --deck deck.json --map map.json --output 输出.pptx --animation-plan-only
<py> scripts/template_fill.py --template 模板.pptx --deck deck.json --map map.json --output 输出.pptx
```

1. 先逐页看实际预览，再结合 inventory 建立语义选页目录，记录封面、目录、章节、流程、对比、表格、并列等角色、适合的信息关系和保守容量。依资料分配事实与槽位，再选择原页，不能只凭对象数选页。`map.slides` 与 `deck.slides` 数量、顺序一一对应。同一原页可复用，未选原页不进入输出。
2. preflight 只读检查拟替换文字容量、原模板正文缺陷、几何 baseline 与声明的删除项。原模板的正文缺陷不能通过 baseline 豁免。
3. 生成器在暂存包中独立克隆所选页，重映射关系，为标签和备注创建独立部件；保留母版、版式与主题关联。纯文字字符串映射直接替换发生变化的槽位，保留原段落和首 run 属性，不重写未变化的富文本。其它对象操作经 OfficeCLI 批量编辑，完成后 flush/close 再由 Python 读取。
4. 应用[模板动画计划](template-animation.md)，写入[讲稿](deck-schema.md#讲稿)，复查内容落位、包关系、容量和布局。strict 最后对比原页的几何、对象数、分组和槽位结构；克隆元数据身份不参与几何比较。任何步骤失败均不替换已有成功输出，模板原件始终只读。
5. 按[最终校验](validation.md)完成后处理、逐页预览与适用计划复核。单独复核 strict 可运行：

```text
<py> scripts/template_contract.py --template 模板.pptx --output 输出.pptx --map map.json
```

脚本自动检查非空且名称匹配 `Text数字` / `Bullet数字` 的文字槽位是否映射；其它命名的语义槽位须依据 inventory 人工逐项核对。几何一致不能证明内容完整。

## 3. map 字段与 key

```json
{
  "mode": "strict",
  "slides": [
    {"template_slide": 1, "fields": {"标题": "封面标题", "副标题": {"text": "汇报人 · 日期", "size": 20}}},
    {"template_slide": 3, "fields": {"标题": "资料支持的完整结论句", "组合 35/文本框 18": "真实论据及成立条件"}}
  ]
}
```

- `template_slide` 为原模板页号，从 1 开始。
- `fields` 的键使用 `analyze_template.py --slides-json` 输出的 key。字符串替换文字，`\n` 分段；对象 `{"text": "…", "size": 20, "color": "tx1"}` 可显式设置字号和主题色。颜色只接受 `tx1/tx2/bg1/bg2/dk1/dk2/lt1/lt2/accent1–6`。
- `{"image": "imgs/a.png"}` 替换已有图片，路径相对 map 文件；重置裁剪，素材宽高比应接近原位。
- `{"table": [["表头"], ["单元格"]]}` 填写原有表格单元格，不创建表格。
- 顶层 key 为形状名；子形状按 `组合名/子形状名` 逐层连接。同层重名带 `#序号`，按 z 序从 1 起。直接复制 inventory 中的完整 key。
- `@id=N` 使用模板原始 ID 定位；同页重复 ID 时改用 key。填充操作按扫描结果转换寻址，不能拿输出页 ID 反推模板 key。
- 组合、连接线、艺术字和图表不能直接填文字；应定位组合内的文本对象。不存在的 key、类型不匹配、非法主题色在编辑前报错并列出可用 key。

## 4. custom 对象编辑

`"mode": "custom"` 必须同时提供非空 `layout_change_plan`，说明容器、连接线、编号、底衬如何一起调整、重新分布和居中。该字段要求存在；计划的语义完整性仍由人工审阅。custom 不称为完全遵循或半自主。

| 字段 | 行为与边界 |
|---|---|
| `remove` | 填充后删除列出的 key；图表同位替换自动加入原位删除清单 |
| `clone_groups` | 仅复制顶层 group；保留字体、颜色、边框、圆角和段落样式。`source` 指组 key；`x/y/width/height` 调整整体几何，`fields` 填写源组内子形状 key |
| `extra_shapes` | 新增文本框：`text/x/y/width/height/size/color`；颜色仅主题色，不加阴影、发光或渐变 |
| `formulas` | `latex/display/x/y/width/height` 创建 `FORMULA_` 占位框；生成后运行 `scripts/formula.py 输出.pptx` 转为原生公式 |
| `fields.chart` | 删除顶层图片或图表，在相同坐标新增原生图表；参数、类型与图片回退见 [charts EXT](../extensions/charts/EXT.md) |

长度接受 EMU 整数或带 `cm/mm/in/pt` 的值。`add_texts` 不是生成器支持的字段，不使用。

```json
{
  "mode": "custom",
  "layout_change_plan": "调整整组容器与连线，复核编号和底衬，按正文区域重新居中。",
  "slides": [{
    "template_slide": 3,
    "fields": {"标题": "资料支持的结论"},
    "clone_groups": [{"source": "组合 13", "x": "2.6cm", "y": "4.2cm", "width": "6.4cm", "height": "8.1cm", "fields": {"组合 13/文本框 8": "真实内容"}}],
    "extra_shapes": [{"text": "图注", "x": "2cm", "y": "13cm", "width": "12cm", "height": "1cm", "size": 12, "color": "tx2"}]
  }]
}
```

## 5. 容量与失败回退

文字预检使用原字号、内边距、换行和实际显示尺寸，包括竖排标签、不换行标题与缩放 group 子形状。替换图片检查原生像素，低于显示需求的 70% 时阻断；Pillow 缺失时仍能读取 PNG/JPEG 头，未知尺寸报错。

| 情况 | 处理 |
|---|---|
| 文字装不下 | 精简而不改变事实，细节放讲稿；仍不足换更大原页或拆页 |
| 原模板示例、英文副标或未用槽位溢出 | 显式填充修正；strict 保留槽位，custom 才能声明删除 |
| 固定图案叠加或出血 | 核对真实渲染后保留设计 baseline，新文字碰撞仍阻断 |
| 零尺寸或无意义对象 | strict 换页；经授权的 custom 在计划中整体清理。保留页眉、页脚、页码和 logo |
| 没有足够容量 | 换页、拆页，或采用用户已授权的半自主模式 |
| 执行或校验失败 | 丢弃暂存包，返回 deck/map/选页修正后重新生成 |

公式字号由目标应用决定，新增公式须给足区域；strict 的行内公式后处理若改变对象结构，仍须重新核验严格契约。交付记录复用原页、实际模式与回退项。预览引擎差异和原生打开限制统一见 validation。
