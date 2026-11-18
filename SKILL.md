---
name: ppt-suggestion
description: >
  基于用户提供的资料（文字/Word/PDF/图片/表格/HTML/URL）与可选 PPT 模板，生成专业、克制、白底极简、
  可在 PowerPoint/WPS 二次编辑的正式汇报 PPT。覆盖学术答辩/课题申报、工程技术评审、商务汇报、
  政务总结、教学课件等场景。用户提到做 PPT、做汇报、排 PPT、答辩材料、根据某份资料生成幻灯片、
  套用某个模板生成 PPT 时使用本 skill，即使没有逐字说"PPT"——只要意图是把资料做成可编辑的演示文稿。
metadata:
  version: "3.0.0"
---
# PPT Suggestion Skill

## 任务与环境

用户提供资料与可选模板，产出像专业人员手工制作的可编辑 `.pptx`。术语、名称、单位、数字与资料一致，不虚构数据、结论或案例；“显著、领先、国际先进”等判断必须有资料依据。

首次使用按 [README 安装说明](README.md#安装) 创建项目 `.venv` 并安装依赖及 OfficeCLI。下文 `<py>` 指 `.venv/bin/python`（macOS/Linux）或 `.venv/Scripts/python`（Windows）；`power-pptx` 的导入名为 `power_pptx`。可选依赖与降级边界也以 README 为准。

## 规则导航

| 职责 | 规则入口 |
|---|---|
| 资料读取、模板分析、事实与缺口 | [source-intake.md](references/source-intake.md) |
| 信息层级、叙事与 STORY | [story.md](references/story.md) |
| 全部文案的语言自查 | [humanize.md](references/humanize.md) |
| 视觉、密度、节奏与采样职责 | [design.md](references/design.md) |
| 核心/特殊版式及图片构图选型 | [layout-types.md](references/layout-types.md) |
| deck 字段、容量与渲染前机械检查 | [deck-schema.md](references/deck-schema.md) |
| 模板继承与 map | [template-inherit.md](references/template-inherit.md) |
| 半自主正文画布、三层清理 manifest 与自主布局 | [template-semi-autonomous.md](references/template-semi-autonomous.md) |
| 模板动画 inventory、计划与核对 | [template-animation.md](references/template-animation.md) |
| 快速路径差异 | [fast-mode.md](references/fast-mode.md) |
| 图表与可选扩展 | [extensions/README.md](extensions/README.md) |
| 结果校验、修复与证据记录 | [validation.md](references/validation.md)、[playback-report.md](references/playback-report.md) |

核心版式无需启用特殊排版。图表能力默认可用，仅在资料有真实数据关系时使用；使用图表时读 charts 的 EXT。动效、几何装饰、特殊排版默认关闭，启用后才读相应 EXT；几何底纹须单独明确允许。

## 六阶段工作流

严格串行：**阅读材料 → 询问用户 → 设置基本版式 → 生成内容与微调版式 → 排版渲染 → 结果校验**。不跨阶段提前准备产物。已有用户指令和选择持续有效，只询问缺失信息；完整与 Quick 在确认门等明确确认，fast-lite 的差异见 fast-mode。

### 一、阅读材料

按 source-intake 读取全部资料，核对转换后的术语、数字、单位与公式。HTML/URL 用 `scripts/source_to_md.py`。

有模板时运行 `<py> scripts/analyze_template.py 模板.pptx -o 模板摘要.md`，另运行 `<py> scripts/analyze_template.py 模板.pptx --slides-json > 模板清单.json`，人工补全模板分析。无论是否启用动效都检查三层动画 inventory；策略与风险处理按 template-animation，静态 XML 清单不能当作播放证明。

⛔ **信息缺口门**：列出缺失项、影响与需要补充的信息，请用户补齐或确认留白；不臆造。整理 `材料摘要.md` 与有模板时的 `模板摘要.md`。

### 二、询问用户

一次问完，只问缺的，能推断的不问：

1. 场合、受众、页数，合计最多 3 个问题。
2. 有模板时确定风格模式：**完全遵循模板**（复制原页、替换已有对象）、**半自主**（保留背景、边缘装饰、页眉页脚及字体配色，清除中央正文及原版式后自主排版）、**完全自主白底极简**、**完全自主用户指定风格**。后两种不继承模板页面；选择写入材料摘要。
3. 是否快速模式：说明完整模式耗时和 token 更多、可控性更强；快速模式减少候选、规划和确认，适合初稿。用户说“直接做”“快速”即采用 fast-lite；复杂任务按 fast-mode 回退。
4. 是否允许复用资料图片。
5. 是否启用动效、几何装饰、特殊排版，说明额外成本；几何底纹另行明确允许。
6. 是否在每页原生备注中加入讲稿，**默认加入**；用户明确不需要时写 `notes_enabled: false`。与前述问题一并询问，不单独增加确认轮次。

快速模式默认：有模板则完全遵循，允许复用资料图片，图表能力可用，其余扩展和底纹关闭。资料缺口仍按 fast-mode 处理。选择记入 `材料摘要.md`。

### 三、设置基本版式

完全遵循与半自主继承模板主题、字体和母页，不运行采样器。两个完全自主模式按 design 采样；用户指定风格写入设计约束并覆盖候选中的相应参数。

```text
<py> scripts/sample_style.py --seed <随机整数> --count 3 --scene <学术|工程|商务|政务|教育> --pages <页数或类型序列> [--force-color HEX]
```

比较候选后按内容取舍；fast-lite 使用 `--count 1`。⛔ **版式确认门**：告知选定配色与风格，等用户确认（fast-lite 跳过）。

完全自主模式运行 `<py> scripts/prepare_renderer.py <结果目录>/渲染器`，复制 render_pptx、text_measure、deck_check、layout_capacity、speaker_notes、pptx_package 六个依赖模块。ID 从 0 开始，保存 seed、候选 ID、参数、完整 palette/base、逐页 type 与用户覆盖项到 `style.json`，渲染时用 `--style style.json` 实际注入并核对 `.render.json` 中的实际 palette。后续项目定制改副本。完全自主从空白演示文稿生成，使用 deck 与 style.json 重现设计。

### 四、生成内容与微调版式

完整或 Quick 按 story 写 `STORY.md`（意图 → 骨架 → 逐页）；fast-lite 保留紧凑事实清单，直接写 deck。按 layout-types 选型、design 检查视觉与密度、schema 写 `deck.json`；采样候选仅作起点。每页讲稿、开关和备注保留行为按 [schema 的讲稿字段](references/deck-schema.md#讲稿)，语言按 humanize。

**必经文案检查**：全部文本按 humanize 全量规则人工自查，包含标题、body、note、insight、表格与图表文字；不以结构脚本替代。⛔ **大纲确认门**：列出每页一行结论句并等用户确认（fast-lite 跳过）。

渲染前必须通过校验，失败按 schema 报错解释修复：

```text
<py> 渲染器/deck_check.py deck.json
# 完全遵循或半自主使用：
<py> scripts/deck_check.py deck.json --template-mode
```

特殊排版接入按其 EXT；`--extra-kinds` 只放行名称，不安装渲染函数。模板模式跳过的检查范围见 schema；快速模式不额外放宽脚本门槛。

### 五、排版渲染

- **完全自主**：`<py> 渲染器/render_pptx.py --input deck.json --style style.json --output "<中文名>.pptx"`。几何装饰启用后同样传入 `--style style.json`，显式记录母题或 seed；装饰脚本核对实际渲染证据，拒绝混用默认 accent。
- **完全遵循**：按 template-inherit 写 map，再运行 `<py> scripts/template_fill.py --template 模板.pptx --deck deck.json --map map.json --output "<中文名>.pptx"`。
- **半自主**：按 template-semi-autonomous 写三层 manifest，运行 `template_canvas.py --template 模板.pptx --manifest mothers.json --output mothers.pptx`，预览并审阅 `.canvas-evidence.json`；再以相同输入加 `--deck deck.json --output 最终.pptx` 注入自主正文布局。禁止以复制模板卡片为主路径。custom 对象编辑接口见 template-inherit。

模板编辑前审阅动画计划；容量或源图清晰度不足就换页、拆页或换图，不缩字号硬塞。失败不覆盖已有成功输出。图表的各模式支持范围与回退操作按 charts EXT。

deck 含 `formula` 或 `$...$` 时，渲染后运行 `<py> scripts/formula.py "<中文名>.pptx"`。可选后处理按“公式 → 几何装饰 → 动效”顺序执行，输出新文件；特殊排版属于渲染器接入。所有形状编辑结束后按 template-animation 复核计划。渲染、校验均为前台任务。

### 六、结果校验与交付

按 validation 完成机械校验、逐页预览与目视复核；模板使用原模板作 baseline。**Hard 必须为零**，Soft 在 `--strict` 时同样阻断。脚本通过不能覆盖目视发现的溢出、遮挡或缺失。

小修限位置、字号（角色锚点 ±2pt）、间距、对齐与底衬；仍装不下就退回内容规划或模板选页。按 validation 的相应生成路径重跑检查。

验证证据和未实测项按 validation 分开记录。交付 `.pptx` 的**绝对路径**，简短列出资料缺口、假设、回退与未验证项，预览图不交付。

## 产物位置

中间文件与产物默认放工作区 `ppt-suggestion 的结果/`：材料/模板摘要、`STORY.md`（完整/Quick）、`style.json` 与 `渲染器/`（完全自主）、`deck.json`、map/manifest 与母页（相应模板模式）、`<中文名>.pptx`、模板动画计划与校验报告。页面管理优先依托母版和版式，有模板时继承其设计语言。
