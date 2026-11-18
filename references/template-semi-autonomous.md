# 半自主：清理正文画布后自主排版

选择合适正文页，仅保留母页背景、边缘装饰、页眉页脚和字体配色。清除正文区域中的所有内容及旧布局，包括卡片、容器、连接线、编号、底衬和嵌套组。正文通过现有 `render_body`/`render_block` 核心渲染器在目标安全区原生坐标绘制；不得用 clone_groups 复制模板卡片作为半自主主路径。

## 可运行入口

```text
<py> scripts/template_canvas.py --template 模板.pptx --manifest mothers.json --output mothers.pptx
<py> scripts/template_canvas.py --template 模板.pptx --manifest mothers.json --deck deck.json --output 最终.pptx
```

同一清理函数生成母页预览与最终稿，避免预览清空但最终仍复制旧卡片。原模板不可覆盖，输出先写临时包并校验后原子发布。`template_mother.py` 为显式对象删除工具，接口见模板动画文档；它不执行三层正文清理契约。

## manifest 契约

mothers 是非空数组。每项含 name、template_slide（1 起）、content_region（EMU 的 x/y/w/h）、style（font/accent/text）、header（[层名,对象名] 数组）及 layers。layers 必须包含 slide、layout、master，每层明确列出 keep、remove、keep_reasons。全部顶层对象必须且只能归入一个清单，嵌套中央组整体移除。保留组混有正文时，先换合适原页或人工拆分并审计；不能把混合组直接豁免。背景通过原始 OOXML 关联继承，不按颜色猜测删除。继承层相同对象的决策必须一致，否则阻断。

keep_reasons 解释每个对象为何是固定装饰、页眉或页脚；保留对象与正文区域相交时，必须在 canvas_overlap_reviewed 中逐个记录实际预览确认，不能批量推断。页眉可通过 header 定位并替换章节文字，保留几何与原始文字样式。

style_extraction 输出模板三层原始字体大小、字体名与 run 样式 XML，以及母版 part；style 是据此并结合 theme 与实际预览确定的显式解析值，不能凭空另采样。输出 `.canvas-evidence.json` 包括保留/删除对象的层、文本、几何、原因与实际注入对象；逐页看母页确认中心确实空白、背景及装饰未丢失，检查继承层正文和旧占位文本是否已经处理。

## 正文和讲稿

正文页提供 mother、title、source 和非空 blocks；讲稿启用时必须有 notes。blocks 直接复用核心 26 个 kind，包括 table、rows、columns、split、steps、chart、figure 和 formula；字段契约见 [schema](deck-schema.md)，图表回退见 [charts EXT](../extensions/charts/EXT.md)。旧 items(title/body) 仅兼容转换为 rows，不另造简易排版。图片相对路径以 deck 所在目录解析。渲染器接收安全区英寸坐标并设置 `enlarge=False`，不事后仅缩放对象 bbox；字体、配色临时注入并在失败时恢复。容量不足就拆页或调整安全区，不缩字号硬塞。

封面、目录、章节、收束页可与正文混合：用 template_slide、fields、keep_text（key 到保留理由）替代 mother。所有非空原文字槽位必须被填写或明确保留；结构页检查原几何和分组不变。正文页先隔离 layout/master 再清理，结构页继续使用原继承层，避免污染共享母版。

讲稿及 notes_enabled 开关统一按 [schema](deck-schema.md#讲稿)；source 必须可回查材料段落。正文不能空壳。canvas-evidence 输出正文字符数、有效语义项、来源和逐项覆盖记录；数量只是证据，不能用凑字代替内容审阅。

## 动画与验证

动画处理与接口边界见 [template-animation.md 的正文画布入口](template-animation.md#正文画布入口)。最终包、布局、背景、对比度、清理证据与逐页目视要求见 [validation.md](validation.md)。
