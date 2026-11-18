# 快速模式路径

快速模式用于用户明确说“快速”“直接做”或选择快速出稿的场景。它减少规划、候选和确认步骤，但仍保留会阻止严重错误的结构与渲染校验。快速模式的结果适合先交付可编辑初稿，复杂任务应自动回退到普通 Quick 流程。

## 成本提示

完整模式会读取并整理更多材料、生成完整 `STORY.md`、进行多轮版式与大纲确认，并保留更多候选比较，因此会消耗更多 token 和时间，但可控性更强。快速模式减少这些步骤，适合信息边界清楚、版式要求明确或需要先得到初稿的任务；用户纠偏机会更少，交付后可能需要调整。

启用拓展功能会增加额外 token、处理时间和出错面。图表能力默认可用，但只有资料里存在真实的数据关系时才生成图表；几何底纹必须单独明确允许，正式汇报默认不使用。

## 模式矩阵

| 模式 | 无模板 | 有原始 PPTX 模板 | 保留的检查 |
| --- | --- | --- | --- |
| 完整 | 完整材料摘要、`STORY.md`、多个候选与确认 | 完整模板 inventory、逐页映射与确认 | 全部结构、版式、渲染和视觉检查 |
| Quick | 普通六阶段流程，减少候选数量 | 普通 inventory 后按页填充 | 全部 Hard 检查，允许少量 Soft |
| fast-lite | 一次事实清单、一个 seed、一个候选、直接生成 `deck.json` | 完全遵循：紧凑 inventory 后批量填充；半自主：三层 inventory、清理正文画布并注入自主布局；完全自主：按无模板路径采样 | 与完整模式相同的检查；目标应用不可用时如实记录未验证 |

## 无模板 fast-lite

1. 只整理一次紧凑事实清单：术语、单位、数字、来源、缺口和页数，不另写完整 `STORY.md`。
2. 生成一个可复现 seed，运行：

   ```text
   <py> scripts/sample_style.py --seed <seed> --count 1 --scene <场景> --pages <页数或类型序列>
   ```

   把 seed、候选序号、场景、页型序列、强制颜色和渲染器常量记录到 `style.json`，保证后续能够复现同一候选。
3. 直接取唯一候选，运行 `scripts/prepare_renderer.py 渲染器` 复制完整运行时，并生成 `deck.json`。不生成完整设计锁定文件，不比较多个候选。
4. 按资料中的真实事实生成页面；随机性只能影响强调色、小范围字号与边距、同等版式候选的 tie-break，不得影响事实、数字、公式、图表数据、模板主题或用户指定颜色。
5. 按顺序运行结构校验、渲染和输出校验：

   ```text
   <py> <渲染器>/deck_check.py deck.json
   <py> <渲染器>/render_pptx.py --input deck.json --style style.json --output <输出.pptx>
   # deck 含 formula block 或 $...$ 时，在这里运行：
   <py> scripts/formula.py <输出.pptx>
   <py> scripts/validate_pptx.py <输出.pptx> --deck deck.json
   <py> scripts/render_preview.py <输出.pptx> -o <临时目录>/preview --engine officecli --render html --sheet
   ```

   Windows 且已安装 PowerPoint 时改用 `--render native`；HTML/Pillow、公式可见性、目标应用打开与实际放映的限制统一按 [validation.md](validation.md) 记录。Pillow 只作近似检查，缺少真实渲染时如实说明。缺少真实数据时跳过图表，不用占位数字补齐。

## 有模板 fast-lite

1. 复用一次紧凑 inventory，记录页数、尺寸、版式、占位符、标题、页脚、图片和图表位置；半自主再记录母页候选、正文区域与 slide/layout/master 三层保留和清除清单。紧凑记录不能省略三层动画及 MC 分支，扫描职责见 [source-intake.md](source-intake.md)，策略、CLI 优先级与前置阻断见 [template-animation.md](template-animation.md)。
2. 完全遵循先运行 `scripts/template_fill.py --animation-plan-only` 审阅计划；半自主按 canvas 的固定清理行为审查三层动画，接口范围见 template-animation。
3. 完全遵循建立 map 并批量填充；半自主先用 template_canvas.py 做三层母页清空检查，再注入自主正文布局。复用同次 inventory，不重复分析每页。
4. 检查模板页数、占位符类型、映射覆盖率与输出页数。图表只按 [charts/EXT.md](../extensions/charts/EXT.md) 保留或同位替换；map 不支持任意坐标新增或自动 SVG 回退。失败不交付半成品。
5. 运行：

   ```text
   # strict/custom 加 --baseline-map map.json；半自主结构页来源在 deck 中
   <py> scripts/validate_pptx.py <输出.pptx> --deck deck.json --baseline <模板.pptx> --baseline-map map.json
   # template_fill 和 template_canvas 输出的计划均须复核：
   <py> scripts/template_animation.py <输出.pptx> --plan <输出.pptx>.animation-plan.json
   ```

   再做一次真实预览。模板自身的越界按 baseline 规则降级，新增的 Hard 问题仍必须修复。动画计划核对目标、时序、转场及版式/母版关系的归一化语义指纹；它不验证实际播放。新增动效必须获明确授权并重建计划，不能沿用编辑前指纹声称保留。

## 自动回退

遇到复杂研究、资料缺口、多个模板融合、复杂图片裁切、公式密集页、无法由前置动画计划安全处理的复杂动画、需要重写模板图表数据或用户要求精细视觉控制时，停止 fast-lite，回退到普通 Quick 或完整模式，并在材料摘要中记录原因。动效新建或重建另走授权后的计划，不把编辑前指纹当作新增效果的证明。

## fast-lite 仍必须阻断的错误

- JSON 无法解析、`slides` 为空、未知 block kind 或必填字段缺失。
- 行内公式 `$` 不成对、图片路径不存在、公式替换失败。
- block 或形状越界、零尺寸、明显文本溢出、严重重叠。
- 图表的分类/系列长度不一致、非有限数值、重复系列名、禁止双轴，或多系列缺少来源和可访问性说明。
- PPTX 无法重新打开、模板页数或映射不一致。

快速模式只减少候选、规划文档和多轮确认，不降低现有脚本门槛。完全自主仍按 design 与 schema 执行；模板模式只使用 `--template-mode` 已定义的豁免范围。不能跳过事实缺口门，也不能用随机内容填充缺失信息。
