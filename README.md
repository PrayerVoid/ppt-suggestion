# ppt-suggestion

基于用户提供的资料（文字 / Word / PDF / 图片 / 表格 / HTML / URL）与可选 PPT 模板，生成**更像人制作的、可在 PowerPoint/WPS 二次编辑**的正式汇报 PPT 的技能。

覆盖学术科研（答辩 / 课题申报 / 结题）、工程技术（方案评审 / 架构说明 / 复盘）、商务咨询（工作汇报 / 经营分析 / 方案提案）、政务公共（政策解读 / 工作总结）、教育培训（课件 / 科普）等场景。

## 核心目标

- 最终效果像该领域专业人士手工制作，**没有明显的 AI 生成痕迹**（视觉与文案双重去 AI 味）。
- 术语、名称、单位、数字与原始资料**完全一致**，不虚构、不夸大。
- 有模板时支持四种选择：完全遵循模板、半自主、完全自主（默认白底极简）、完全自主（用户指定其它风格）。前两种用 officecli 继承或扩展模板设计语言。
- 输出优先为**原生可编辑对象**，包括数学公式（PowerPoint 内可直接编辑的 OMML 公式）；图表在原生类型不支持或兼容性不稳定时允许用 SVG 兜底。
- 配色与版式由随机采样器给出候选，AI 保留最终裁量，避免千篇一律。

## 工作流（六阶段，严格串行）

```
① 阅读材料 ─► ② 询问用户 ─► ③ 设置基本版式 ─► ④ 生成内容与微调版式 ─► ⑤ 排版渲染 ─► ⑥ 结果校验
   素材转换      模板/快速/图片    随机采样配色       STORY + deck.json        渲染器副本/      机械校验 +
   模板分析      /拓展             复制渲染器副本     去AI味 + 大纲确认        模板填充/公式    officecli 出图
```

| 阶段 | 做什么 | 产物 | 门禁 |
| --- | --- | --- | --- |
| ① 阅读材料 | 转 Markdown、判断领域、模板分析、充分性检查；有模板时无论是否启用动效都做三层动画 inventory（含 MC 分支），分开记录对象动画与转场 | `材料摘要.md`、`模板摘要.md`、动画清单 | 资料缺口和动画风险先处理，绝不臆造或静默清理 |
| ② 询问用户 | 场合受众页数；四种模板风格模式、快速模式、复用资料图片、可选拓展（图表默认开启） | 记入 `材料摘要.md` | ⛔ 一次问完 |
| ③ 设置基本版式 | `sample_style.py` 采样配色与版式，复制渲染器三件套并改常量 | `渲染器/`、`style.json` | 配色确认（快速模式跳过） |
| ④ 生成内容与微调版式 | 意图→骨架→逐页；全量去 AI 味；`deck_check.py` | `STORY.md`、`deck.json` | 大纲确认（快速模式跳过） |
| ⑤ 排版渲染 | 渲染器副本 / `template_fill.py` → `formula.py` → `chart` / 可选拓展；模板动效先生成计划再编辑后核对 | `<中文名>.pptx`、`.animation-plan.json` | 校验失败禁止跳过 |
| ⑥ 结果校验 | `validate_pptx.py` + `render_preview.py` 出图逐页复核；另行记录结构、打开、静态渲染、实际放映 | `校验报告.md` | 修复闭环后交付 |

完整模式会消耗更多 token 和时间，但可控性更强；快速模式按 `references/fast-mode.md` 走 fast-lite 路径，减少候选、规划和确认，保留严重错误校验。启用动效、几何装饰或特殊排版会增加额外 token、时间和出错概率；几何底纹还需单独明确允许。

## 目录结构

```
ppt-suggestion/
├── README.md                # 本文件
├── SKILL.md                 # 工作流总控：六阶段、门禁、纪律、文件导航
├── requirements.txt         # Python 依赖（国内走清华源）
├── references/
│   ├── source-intake.md     # 阶段一：阅读材料
│   ├── story.md             # 阶段四：意图/骨架/逐页/版式选用
│   ├── humanize.md          # 阶段四：去 AI 味全量规则
│   ├── deck-schema.md       # deck.json 结构说明与质量检查清单
│   ├── design.md            # 视觉/版面/密度边界（采样先验 + 硬约束）
│   ├── template-inherit.md  # 完全遵循模板：OfficeCLI 编辑模板副本
│   ├── template-semi-autonomous.md # 半自主：内容页母页与组合扩展
│   ├── fast-mode.md         # 完整 / Quick / fast-lite 路径
│   └── validation.md        # 阶段六：结果校验（Hard/Soft 分级）
├── scripts/
│   ├── source_to_md.py      # HTML/URL/文本 → Markdown
│   ├── analyze_template.py  # 模板分析 → 模板摘要
│   ├── sample_style.py      # 配色 / 基本版式 / 内容版式随机采样器
│   ├── deck_check.py        # deck.json 结构校验（渲染前）
│   ├── render_pptx.py       # 渲染器原件（阶段三复制到项目里使用）
│   ├── text_measure.py      # 文本估算（渲染器/校验共用）
│   ├── template_fill.py     # 完全遵循/半自主：OfficeCLI 复制、填充、图表替换与组合扩展
│   ├── template_mother.py   # 半自主：从模板正文页生成清空内容的母页
│   ├── template_animation.py # 模板动效策略、计划、应用与生成后语义核对
│   ├── animation_inventory.py # 母版/版式/实际页及 MC 动画、转场静态清单
│   ├── style_showcase.py    # 随机风格候选的可视化展示 PPTX
│   ├── formula.py           # 把公式占位框换成原生可编辑公式
│   ├── ocli.py              # officecli 调用封装（定位二进制、batch、落盘）
│   ├── pptx_package.py      # 包引用检查与页数元数据修正（生成路径不自动清理动画）
│   ├── validate_pptx.py     # 包结构及递归布局校验（含组合内文字）
│   └── render_preview.py    # pptx 逐页转 PNG 供目视自查
└── extensions/              # 图表默认开启；其余拓展用户选了才读
    ├── README.md            # 拓展索引
    ├── charts/              # 图表：原生 power_pptx 优先，必要时 SVG 兜底
    ├── animation/           # 动效：转场 + 入场（power_pptx）
    ├── geometry-deco/       # 几何装饰：多种可复用几何母题
    └── special-layout/      # 特殊排版：核心数字、非对称、引文、斜切、关系、边注、层级与四象限
```

## 安装

```bash
python -m venv .venv
.venv/Scripts/python -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt   # macOS/Linux 用 .venv/bin/python
.venv/Scripts/python -m officecli install
```

- Python 包（`power-pptx`、`officecli-sdk`、`lxml`、`Pillow`、`matplotlib`、`cairosvg`）国内走清华源 `https://pypi.tuna.tsinghua.edu.cn/simple`。`power-pptx` 的导入名是 `power_pptx`；图表 SVG 兜底依赖 `matplotlib` 与 `cairosvg`。
- `python -m officecli install` 运行官方安装脚本，从 `d.officecli.ai` 或 GitHub 下载 officecli 二进制，国内可能超时。清华源、npm 镜像和 Gitee 镜像（`gitee.com/zhu_evan/OfficeCLI` 只有源码、没有发行版）都不提供二进制。
- 下载失败时手动安装：到 <https://github.com/iOfficeAI/OfficeCLI/releases> 下载本平台文件（Windows 为 `officecli-win-x64.exe`，其它平台见同页文件名），对照同页 `SHA256SUMS` 校验，放到 `%LOCALAPPDATA%\OfficeCLI\officecli.exe`（Windows）或 `~/.local/bin/officecli`（macOS/Linux）。`scripts/ocli.py` 会先查 PATH，再查这两个位置。

## 快速上手

```bash
# 1. 阅读材料（源文件转 Markdown）；有模板时分析模板
python scripts/source_to_md.py 参考资料.html -o 材料.md
python scripts/analyze_template.py 模板.pptx -o 模板摘要.md

# 阶段一：有模板时无论是否启用动效，都先扫描母版、版式、实际页及 MC 分支
python scripts/animation_inventory.py 模板.pptx > 模板.animation-inventory.json

# 半自主：按模板摘要选择正文页，清除明确列出的原始内容，先生成动效计划
python scripts/template_mother.py --template 模板.pptx --manifest mothers.json \
  --output mothers.pptx --animation-plan-only
python scripts/template_mother.py --template 模板.pptx --manifest mothers.json --output mothers.pptx
# 完全遵循模板：先审阅编辑前计划，再执行填充（计划模式不调用 OfficeCLI）
python scripts/template_fill.py --template 模板.pptx --deck deck.json --map map.json \
  --output 汇报.pptx --animation-plan-only

# 检查随机风格候选（无模板或完全自主模式）
python scripts/style_showcase.py --seed 20261116 --count 5 --scene 商务 --output style-showcase.pptx

# 2. 设置基本版式：采样配色/版式，复制渲染器三件套并套用常量
python scripts/sample_style.py --seed 7 --count 3 --scene 工程 --pages 12
python scripts/sample_style.py --seed 7 --emit-constants 1

# 3. 渲染前校验 deck.json（用渲染器副本里的 deck_check.py）
python 渲染器/deck_check.py deck.json

# 4. 渲染；deck 含公式时再替换为原生公式
python 渲染器/render_pptx.py --input deck.json --output 汇报.pptx
python scripts/formula.py 汇报.pptx

# 5. 校验与出图自查
python scripts/validate_pptx.py 汇报.pptx --deck deck.json
python scripts/render_preview.py 汇报.pptx -o <临时目录>/ppt_preview --sheet
```

### 四种模板风格模式

- **完全遵循模板**：按 `references/template-inherit.md` 选择原页，`template_fill.py` 复制模板页并替换已有对象。模板图片位可以换图；顶层图片或图表位也可以用 `fields` 的 `chart` 对象删除后在相同坐标添加 OfficeCLI 原生图表。
- **半自主**：按 `references/template-semi-autonomous.md` 选择正文母页。`template_mother.py` 只清除 manifest 中明确列出的文字、图片或内容组合，保留背景、装饰、字体、主题色和页眉页脚；最终 `map.json` 的 `clone_groups` 可以复制模板组合并修改组内文字，扩展新的卡片、步骤或说明块。
- **完全自主（默认白底极简）**：运行 `sample_style.py`，用渲染器生成白底、低饱和、原生可编辑对象。
- **完全自主（其它风格，用户指定）**：仍走无模板渲染器，但将用户指定的颜色、字体、密度和参考风格写入设计约束；随机采样只能在约束允许的范围内提供候选。

模板模式中的图表职责分工是：`power_pptx` 负责无模板渲染器的 native chart；OfficeCLI 负责模板页复制、组合复制，以及模板图片/图表位的 native chart 同位替换。当前 map 不支持任意坐标新增 chart 或自动 SVG 回退；这些需要明确的额外编辑流程。复杂嵌套图片位应改用顶层图片位或另写明确的图片回退操作。

模板生成先在编辑前对母版、版式和实际页做动画 inventory，再依据 `animation_policy`（全局默认及逐页覆盖 `auto`/`preserve`/`strip`）生成 `.animation-plan.json`。`--animation-policy` 可覆盖全局默认，但逐页策略仍优先；`--animation-plan-only` 只把计划写到 stdout，不调用 OfficeCLI。计划包含原因、对象动画与转场动作、共享部件关系和语义指纹。

默认 `auto` 对受影响页整树取消对象时序；不做一对多动画映射。普通转场在内容编辑时改为手动推进并取消转场声音；Morph 在前后页关系或内容改变时取消。段落构建遇到文本变化保守取消。含母版/版式继承时序的页，在 `auto`/`preserve` 编辑前拒绝；共享部件同时要求保留和清除也拒绝。`preserve` 无法安全支持时编辑前报错，不静默清理。`strip` 清除对象 timing、构建列表和转场，涉及的共享层也清除，但不承诺全面清除媒体或超链接等交互。

OfficeCLI 编辑后由计划应用清理动作，再核对目标、时序、转场和版式/母版关系的归一化语义指纹；新增动效必须另行授权并重建计划，不能拿旧指纹宣称保留。生成后可再次运行：

```bash
python scripts/template_animation.py 输出.pptx --plan 输出.pptx.animation-plan.json
```

动画错误由前置计划或最终校验阻断。生成路径调用 `repair_package(repair_animations=False)`，只更新页数元数据；不会自动发现悬空就清整页，其他动画错误保持阻断。失败保留既有成功输出。

### 公式

`formula` block 的 `latex` 字段直接写 LaTeX。渲染器画一个 `FORMULA_<序号>` 占位文本框，并写出 `<输出名>.formulas.json`；
`formula.py` 用 officecli 在原位置新增公式对象并删除占位框，结果是 PowerPoint 内可编辑的原生公式（OMML）。
写入失败整批回滚、非 0 退出，不会降级成图片。

### 预览渲染说明

- `render_preview.py` 默认 `--engine auto`，用 officecli 逐页截图；`--render auto|html|native` 选 officecli 的渲染路径，默认 `html`（跨平台，结果可复现）。
- `html` 渲染会多画母版里的日期/页码占位，是渲染假象，不是 pptx 里的内容。`native` 调用本机 PowerPoint（仅 Windows 且已安装），最贴近真实效果，也更慢。
- `--engine pillow` 是无 officecli 时的近似兜底，看不到公式和动画，用了它交付时要说明"未经真实渲染"。
- officecli 常驻进程会延迟落盘，其他程序读文件前需要先 close；`ocli.py` 与各脚本已处理。
- 每页必须目视复核。看到溢出必须修复，即使脚本为零；小修无效就退回内容规划或模板选页阶段。HTML 预览、机械校验、Mac/Windows Office 实际打开分别记录，不能互相替代。

回归检查：`python -m unittest discover -s tests -v` 和 `python -m unittest discover -s scripts -p test_pptx_package.py -v`，覆盖组合内文字、字符间距、不换行、动画引用、缺失关系与生成失败保留旧输出。

## 定制指南

- **改设计规范**：调整 `references/design.md`（视觉/密度）、`references/template-inherit.md`（模板模式）、`references/source-intake.md`（阅读材料）等。
- **改视觉**：无模板的配色/字号/间距集中在渲染器顶部常量，JSON 里不允许出现颜色字段。阶段三把渲染器复制到项目后改副本，不改原件。
- **改采样分布**：`sample_style.py` 顶部的权重倍数常量，每项注明对应的文档规则。
- **改校验强度**：`validate_pptx.py` 用 `--strict` 让警告也失败；调整阈值在脚本内常量。
- **新增版式**：在渲染器的 `BLOCK_RENDERERS` 里加一个渲染函数，并在 `deck-schema.md` 登记；渲染前校验用 `--extra-kinds` 放行。
- **新增拓展**：在 `extensions/` 下建子目录，写 `EXT.md`，并登记到 `extensions/README.md`。
