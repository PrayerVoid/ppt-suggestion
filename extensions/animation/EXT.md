# 拓展：动效（转场 + 入场）

仅在用户选了"动效"时使用。对**成品 pptx 后处理**，输出新文件，不改原件。依赖 `power_pptx`（只用来设转场、读写文件）。

## 用法

```
python extensions/animation/apply_animation.py 成品.pptx --deck deck.json --out 成品_动效.pptx \
    [--transition fade|push|wipe|cover|cut|morph|none] [--entrance fade|wipe|appear|none] [--com-check]
```

默认：转场 `fade`（500 ms）、入场 `fade`、点击触发。

## 默认方案

- 转场：第 2 页起统一一种，克制为先；可选 `fade`、`push`、`wipe`、`cover`、`cut`、`morph`。`morph` 需要 PowerPoint 2019+，旧版本不支持，且只在上下页有同名/同位形状时才有意义。
- 入场：只加在 `catalog` 与 `content` 页；`cover`、`section`、`close` 不加。可选 `fade`、修正为标准 PowerPoint XML 的 `wipe`、最轻量的 `appear`。
- 分组：内容区（标题线之下、页脚之上）的顶层形状按垂直位置聚成"行"，行数向 deck.json 里该页 block 数靠拢。每行一次点击：第一个形状 `clickEffect`，其余形状 `withEffect`，所以一页有几个 block 就点几次。组合形状整体做动画，不对组合内子形状单独加。
- 跳过名字以 `DECO_` 开头的装饰形状；标题、页脚等内容区之外的形状始终可见。

## 实现与自检

- 入场动画按 PowerPoint 自己写出的结构生成：`mainSeq` 主序列 → 每次点击一个外层 `par` → 内层 `par` 里放该行所有效果，并写 `bldLst`。不使用 `power_pptx` 的 `slide.animations`：它不写 `mainSeq`，PowerPoint 读回主序列为空，播放时部分形状不出现。
- 当前入场预设只保留 `fade`、`wipe`、`appear`；其中 `wipe` 使用标准 `presetID=8`、`presetSubtype=2` 和 `wipe(dir=left)`。转场选项为 `fade`、`push`、`wipe`、`cover`、`cut`、`morph`。
- 保存后自动自检：目标 id 都存在于该页且无重复、没有形状被重复加入场、主序列唯一、每次点击由 1 个 `clickEffect` 加若干 `withEffect` 组成、效果数与点击数和计划一致，失败返回非 0。
- `--com-check`（仅 Windows、需 `pip install pywin32` 与已装 PowerPoint）：用 PowerPoint 打开成品，读回每页主序列的效果数、点击数和形状名，和计划逐页核对。自检通过不等于已放映验证，节奏仍建议在 PowerPoint 里放映一遍。

## 顺序与限制

- 模板输入和模板生成必须先运行 `analyze_template.py` 的动画 inventory；它会分别检查 slide、slideLayout、slideMaster 和所有 `mc:AlternateContent` 分支。模板已有动画按 `references/template-animation.md` 先形成 `auto|preserve|strip` 计划，不能等 PowerPoint 打开后再补救。
- 动画由形状 id 引用，必须在 `formula.py`（它会删除占位框、新增公式形状）和几何装饰**之后**运行；此后再用 officecli 或其他工具增删形状，会让动画目标失效，需要重新生成动效计划并重跑本脚本。
- 只做点击触发的入场；不支持自动播放、退出、强调和动作路径效果。模板已有动作路径、媒体、命令、交互触发或段落构建在编辑时由动画策略处理；`preserve` 无法证明时提前失败，`auto` 保守静态化，不能盲目复制旧目标。
- 模板模式的页面形状布局不由本仓库控制，内容区按页高比例判断，必要时调整 `CONTENT_TOP_FRAC` / `CONTENT_BOTTOM_FRAC`。
- 同一页里若形状外框纵向重叠，会被并入同一行（同一次点击）。
- 新增动效必须另做目标/点击计划；它不继承模板动画的语义指纹。`--com-check` 或 PowerPoint 放映检查另行记录。
- 动画结果检查分为包结构、静态渲染、目标应用打开和实际放映四层；静态出图和 `template_animation.py` 通过都不能声称播放顺序已经验证。
