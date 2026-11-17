# ppt-suggestion

基于用户提供的文字、Word、PDF、图片、表格、HTML/URL 与可选 PPT 模板，生成可在 PowerPoint/WPS 二次编辑的正式汇报 PPT。适用于学术答辩、课题申报、工程评审、商务汇报、政务总结和教学课件。

内容以资料为准，术语、数字与单位保持一致，不虚构或夸大。输出优先使用原生文字、形状、图表和可编辑数学公式；SVG/PNG 图表回退作为图片交付时须说明编辑限制。

## 工作方式

按“阅读材料 → 询问用户 → 设置基本版式 → 生成内容与微调版式 → 排版渲染 → 结果校验”六阶段执行。完整流程、确认门和产物位置以 [SKILL.md](SKILL.md) 为准；用户要求“快速”或“直接做”时按 [fast-mode.md](references/fast-mode.md) 减少候选与确认，仍保留必要校验。

支持四种风格模式：完全遵循模板、半自主、完全自主（默认白底极简）、完全自主（用户指定其它风格）。前两种继承模板主题与母版，后两种使用采样器提供候选并按内容调整。核心提供 26 个 kind；特殊排版扩展另提供 4 个 kind。图表能力默认可用，仅在有真实数据关系时生成；动效、几何装饰和特殊排版须启用，几何底纹还需单独明确允许。

## 文档职责与入口

| 文档 | 负责内容 |
|---|---|
| [SKILL.md](SKILL.md) | 工作流总控、阶段顺序、用户选择、门禁与执行导航 |
| [source-intake.md](references/source-intake.md) | 资料读取、事实摘录、领域判断、模板分析与缺口清单 |
| [story.md](references/story.md) | 信息层级、叙事骨架、逐页意图与 STORY 结构 |
| [humanize.md](references/humanize.md) | 全部文案的语言规则与人工自查 |
| [design.md](references/design.md) | 视觉、字体、密度、页面节奏与采样职责 |
| [layout-types.md](references/layout-types.md) | 核心/特殊版式选型、页面与图片构图、能力边界 |
| [deck-schema.md](references/deck-schema.md) | deck 字段、block 容量、渲染前机械检查与报错解释 |
| [template-inherit.md](references/template-inherit.md) | 完全遵循模板、map 格式与 OfficeCLI 编辑边界 |
| [template-semi-autonomous.md](references/template-semi-autonomous.md) | 半自主母页清空、manifest 与组合复制扩展 |
| [template-animation.md](references/template-animation.md) | 模板动画 inventory、策略、前置计划与语义核对 |
| [fast-mode.md](references/fast-mode.md) | 完整 / Quick / fast-lite 的执行差异 |
| [extensions/README.md](extensions/README.md) | 图表及可选扩展的入口和执行顺序 |
| [validation.md](references/validation.md) | 最终 Hard/Soft 检查、预览、修复闭环与验证限制 |
| [playback-report.md](references/playback-report.md) | 结构、静态渲染、目标应用打开和实际放映的证据记录 |

## 安装

在 skill 根目录创建虚拟环境；下文 `python` 均指该环境的解释器。

```bash
python -m venv .venv
# macOS/Linux
.venv/bin/python -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt
.venv/bin/python -m officecli install
# Windows 将 .venv/bin/python 换成 .venv/Scripts/python
```

| 依赖层 | 安装文件 | 用途与缺失处理 |
|---|---|---|
| 核心 | `requirements-core.txt` | `power-pptx`、`officecli-sdk`、`lxml`，用于核心渲染、模板和结构校验；导入名为 `power_pptx` |
| 可选 | `requirements-optional.txt` | Pillow 用于图片处理/近似预览，PyMuPDF 用于 PDF，matplotlib + cairosvg 用于图表 SVG/PNG 回退；缺少对应包时安装或明确报告该能力不可用，不影响无关核心入口 |
| 完整 | `requirements.txt` | 汇总以上两层 |

模板图片容量预检在 Pillow 缺失时仍能读取 PNG/JPEG 头部；无法识别尺寸会报错，不能把未知清晰度当作通过。近似预览、PDF 转换和图表回退不能在依赖缺失时宣称已完成。

OfficeCLI Python 包和二进制分别安装。`python -m officecli install` 从 `d.officecli.ai` 或 GitHub 下载二进制；清华源、npm/Gitee 源码镜像不提供该文件。下载失败时从 [OfficeCLI releases](https://github.com/iOfficeAI/OfficeCLI/releases) 获取本平台文件，对照 `SHA256SUMS` 校验后放入 `%LOCALAPPDATA%\OfficeCLI\officecli.exe`（Windows）或 `~/.local/bin/officecli`（macOS/Linux）。[ocli.py](scripts/ocli.py) 先查 PATH，再查这些位置。

创建图表或 KPI 时按 dataviz skill 检查色板。其完整检查优先用 Node；没有 Node 时可运行 `python scripts/palette_check.py "#1A1A1A,#C45A3C,#2D6A73" --mode light`，只验证基础明度、对比度与相邻颜色，不替代完整 CVD 检查。

## 快速上手

以下是完全自主路径的脚本顺序；执行前先按 SKILL 读取材料、确定场景与大纲。

```bash
python scripts/source_to_md.py 资料.html -o 材料.md
python scripts/sample_style.py --seed 7 --count 3 --scene 工程 --pages 12
python scripts/sample_style.py --seed 7 --emit-constants 0
```

将 `scripts/render_pptx.py`、`scripts/text_measure.py`、`scripts/deck_check.py` 复制到结果目录的 `渲染器/`，把选中候选常量写入副本，并保存完整 palette、base 与采样参数到 `style.json`。候选 ID 从 0 开始；按 schema 写好 `deck.json` 后运行：

```bash
python 渲染器/deck_check.py deck.json
python 渲染器/render_pptx.py --input deck.json --output 汇报.pptx
# 有 formula block 或 $...$ 行内公式时执行
python scripts/formula.py 汇报.pptx
python scripts/validate_pptx.py 汇报.pptx --deck deck.json
python scripts/render_preview.py 汇报.pptx -o /tmp/ppt_preview --sheet
```

模板路径先运行 `python scripts/analyze_template.py 模板.pptx -o 模板摘要.md --slides-json`，然后按模板继承或半自主文档生成 map/manifest；不运行配色采样。图表字段见 schema，模板图表操作边界见 [charts/EXT.md](extensions/charts/EXT.md)。模板输出的动画计划在所有后处理之后复核。公式、预览引擎与目标应用的限制统一见 validation。

默认工作目录为 `ppt-suggestion 的结果/`；中间文件保留在此，交付 `.pptx` 的绝对路径，并如实说明资料缺口、图片图表回退和未验证项。

## 维护与验证

文档规则在上表对应文件维护，跨文档用导航引用。无模板视觉参数集中在渲染器顶部常量；项目定制改渲染器副本，JSON 不携带视觉字段。调整采样分布改 `scripts/sample_style.py` 的权重；文案句式由 humanize 人工检查，不放入结构校验脚本。

添加版式时实现并注册 `BLOCK_RENDERERS`，在 schema 登记字段，更新选型、校验 kind 名单及适用测试；副本扩展通过 `--extra-kinds` 放行。添加扩展时提供 `extensions/<名>/EXT.md` 并登记扩展索引。稳定脚本入口保留在 `scripts/`。

```bash
python -m unittest discover -s tests -v
python -m unittest discover -s scripts -p test_pptx_package.py -v
```
