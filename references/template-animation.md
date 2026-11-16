# 模板动画处置

模板动画必须在模板分析阶段确认。静态 XML 清单可以确认部件、目标 ID、构建范围、触发器、主序列/交互序列、动作路径、媒体、命令、转场、自动翻页、声音和 Morph；它不能证明 PowerPoint 放映时的节奏和效果。`analyze_template.py` 会把实际页、slideLayout 和 slideMaster 分开扫描，并检查所有 `mc:AlternateContent` 分支。页面对象的 `spid` 只在本页面解析，版式和母版的 ID 只在各自部件解析。

```text
<py> scripts/analyze_template.py 模板.pptx -o 模板摘要.md
<py> scripts/template_fill.py --template 模板.pptx --deck deck.json \
    --map map.json --output 输出.pptx --animation-plan-only
```

`map.json` 和 `mothers.json` 可以写顶层 `animation_policy`，单页可以写 `animation_policy` 覆盖它；命令行 `--animation-policy auto|preserve|strip` 覆盖顶层默认值。`--animation-plan-only` 只做编辑前计划，不启动 OfficeCLI，不写输出 PPTX。

| 策略 | 编辑前行为 | 输出行为 |
| --- | --- | --- |
| `preserve` | 只接受能证明目标、时序、构建范围和关系仍可对应的编辑；删除动画目标、改变段落构建、复制对象、复杂图表/媒体/动作路径、Morph 关系变化和共享继承动画都会提前报错 | 不改变对象动画和转场；最终做归一化语义指纹核对 |
| `auto` | 根据编辑风险逐页决定 | 无法定义一对一映射时，清除受影响页的对象时序和构建列表；不把旧动画盲目复制给克隆组。普通转场可保留，但改为手动翻页并取消自动翻页及转场声音；Morph 关系改变时清除该页转场 |
| `strip` | 记录显式静态化原因；共享母版/版式部件必须对所有引用页统一策略 | 清除对象时序、构建列表、转场及其自动翻页和声音；共享部件按统一策略清除 |

这里的“受影响页”指最终复制出的页，而不是模板原页。文字替换只在整形状动画、没有段落或字符范围构建且目标仍能定位时保留对象时序；文本段落数量、`pRg`/`charRg` 范围或图表构建发生变化时，`auto` 会静态化该页。删除或替换动画目标、删除图片并放入图表、复制组合和新增形状都会改变对象集合；当前工作流不猜测一对多的点击顺序，`auto` 会静态化，`preserve` 会在调用 OfficeCLI 前失败。未受影响页的合法动画可以保留。

母版或版式中的时序是共享状态。不能为了一个页面静默修改共享部件；编辑带继承时序的页面在 `auto`/`preserve` 下提前阻断，只有显式且相关页面一致的 `strip` 才清除。Morph 依赖相邻页面的对象身份和顺序；复制、删除、重排或改变对应内容时，`auto` 清除受影响转场，`preserve` 提前失败。转场与对象动画始终分开记录和处置。

动画策略在 OfficeCLI、公式替换、图表替换和其它最后一次形状编辑之前确定。输出完成后由 `template_fill.py`/`template_mother.py` 应用计划，再用 `repair_package(..., repair_animations=False)` 检查包；严格生成路径不会因为悬空 `spid` 悄悄删除整页动画。计划写入 `<输出>.animation-plan.json`，其中保留每页原因、源页和版式/母版关联。所有后处理结束后运行：

```text
<py> scripts/template_animation.py 输出.pptx \
    --plan 输出.pptx.animation-plan.json
```

核对包含页数、版式/母版关联、对象目标、时序节点依赖、段落范围、转场和关系引用的归一化指纹；对象 ID 重编本身不会误报，目标换成另一形状会失败。失败时保留既有成功输出，回到 map、母页或策略重做。`template_animation.py` 的通过信息只表示 XML 计划一致，不表示已经放映。

如果用户启用 `extensions/animation` 新增入场动效，必须在公式、图表和所有形状增删之后运行。它不是模板动画保留的证明：新增动效应生成自己的目标/点击计划并自检，不能拿模板的旧指纹比较。`--com-check` 或 PowerPoint 放映检查另行记录。

最终报告分开记录四项：包结构与引用、静态布局渲染、目标应用能否打开、实际放映播放。HTML/Pillow/原生 PDF 都不能替代动画播放检查；没有放映验证时必须明确写 `playback_verified=false`。
