# ppt-suggestion

基于用户提供的资料（文字 / Word / PDF / 图片 / 表格 / HTML / URL）与可选 PPT 模板，生成**更像人制作的、可在 PowerPoint/WPS 二次编辑**的正式汇报 PPT 的技能。

覆盖学术科研（答辩 / 课题申报 / 结题）、工程技术（方案评审 / 架构说明 / 复盘）、商务咨询（工作汇报 / 经营分析 / 方案提案）、政务公共（政策解读 / 工作总结）、教育培训（课件 / 科普）等场景。

## 核心目标

- 最终效果像该领域专业人士手工制作，**没有明显的 AI 生成痕迹**（视觉与文案双重去 AI 味）。
- 术语、名称、单位、数字与原始资料**完全一致**，不虚构、不夸大。
- 用户要求完全遵循模板时，用 officecli 直接编辑模板副本，继承模板的设计语言。
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
| ① 阅读材料 | 转 Markdown、判断领域、模板分析、充分性检查 | `材料摘要.md`、`模板摘要.md` | 资料缺口先问用户，绝不臆造 |
| ② 询问用户 | 场合受众页数；是否完全遵循模板、快速模式、复用资料图片、可选拓展（图表默认开启） | 记入 `材料摘要.md` | ⛔ 一次问完 |
| ③ 设置基本版式 | `sample_style.py` 采样配色与版式，复制渲染器三件套并改常量 | `渲染器/`、`style.json` | 配色确认（快速模式跳过） |
| ④ 生成内容与微调版式 | 意图→骨架→逐页；全量去 AI 味；`deck_check.py` | `STORY.md`、`deck.json` | 大纲确认（快速模式跳过） |
| ⑤ 排版渲染 | 渲染器副本 / `template_fill.py` → `formula.py` → `chart` / 可选拓展 | `<中文名>.pptx` | 校验失败禁止跳过 |
| ⑥ 结果校验 | `validate_pptx.py` + `render_preview.py` 出图逐页复核 | `校验报告.md` | 修复闭环后交付 |

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
│   ├── template-inherit.md  # 模板模式：officecli 编辑模板副本
│   ├── fast-mode.md         # 完整 / Quick / fast-lite 路径
│   └── validation.md        # 阶段六：结果校验（Hard/Soft 分级）
├── scripts/
│   ├── source_to_md.py      # HTML/URL/文本 → Markdown
│   ├── analyze_template.py  # 模板分析 → 模板摘要
│   ├── sample_style.py      # 配色 / 基本版式 / 内容版式随机采样器
│   ├── deck_check.py        # deck.json 结构校验（渲染前）
│   ├── render_pptx.py       # 渲染器原件（阶段三复制到项目里使用）
│   ├── text_measure.py      # 文本估算（渲染器/校验共用）
│   ├── template_fill.py     # 模板模式：复制模板原页并填充
│   ├── formula.py           # 把公式占位框换成原生可编辑公式
│   ├── ocli.py              # officecli 调用封装（定位二进制、batch、落盘）
│   ├── validate_pptx.py     # 输出校验（重叠/越界/溢出/图片/版式）
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

### 公式

`formula` block 的 `latex` 字段直接写 LaTeX。渲染器画一个 `FORMULA_<序号>` 占位文本框，并写出 `<输出名>.formulas.json`；
`formula.py` 用 officecli 在原位置新增公式对象并删除占位框，结果是 PowerPoint 内可编辑的原生公式（OMML）。
写入失败整批回滚、非 0 退出，不会降级成图片。

### 预览渲染说明

- `render_preview.py` 默认 `--engine auto`，用 officecli 逐页截图；`--render auto|html|native` 选 officecli 的渲染路径，默认 `html`（跨平台，结果可复现）。
- `html` 渲染会多画母版里的日期/页码占位，是渲染假象，不是 pptx 里的内容。`native` 调用本机 PowerPoint（仅 Windows 且已安装），最贴近真实效果，也更慢。
- `--engine pillow` 是无 officecli 时的近似兜底，看不到公式和动画，用了它交付时要说明"未经真实渲染"。
- officecli 常驻进程会延迟落盘，其他程序读文件前需要先 close；`ocli.py` 与各脚本已处理。

## 定制指南

- **改设计规范**：调整 `references/design.md`（视觉/密度）、`references/template-inherit.md`（模板模式）、`references/source-intake.md`（阅读材料）等。
- **改视觉**：无模板的配色/字号/间距集中在渲染器顶部常量，JSON 里不允许出现颜色字段。阶段三把渲染器复制到项目后改副本，不改原件。
- **改采样分布**：`sample_style.py` 顶部的权重倍数常量，每项注明对应的文档规则。
- **改校验强度**：`validate_pptx.py` 用 `--strict` 让警告也失败；调整阈值在脚本内常量。
- **新增版式**：在渲染器的 `BLOCK_RENDERERS` 里加一个渲染函数，并在 `deck-schema.md` 登记；渲染前校验用 `--extra-kinds` 放行。
- **新增拓展**：在 `extensions/` 下建子目录，写 `EXT.md`，并登记到 `extensions/README.md`。
