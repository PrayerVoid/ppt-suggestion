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

## 任务

用户提供资料与可选模板，你产出像专业人员手工制作的、无明显 AI 痕迹的可编辑 `.pptx`。
先做对内容，再谈美观；不虚构、不夸大，术语与数字与原始资料完全一致。

## 环境（首次使用）

```
python -m venv .venv
.venv/Scripts/python -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt
.venv/Scripts/python -m officecli install
```

macOS/Linux 把 `.venv/Scripts/python` 换成 `.venv/bin/python`。下文 `<py>` 都指这个 venv 的 python（导入名是 `power_pptx`）。
officecli 二进制下载失败时的处理见 `README.md` 的"安装"一节。脚本通过 `scripts/ocli.py` 自动定位二进制。

## 文件导航


| 阶段                  | 必读                                                                           | 必跑 / 产物                                                                                         |
| ----------------------- | -------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| ① 阅读材料           | `references/source-intake.md` + `references/template-animation.md`             | `source_to_md.py`（HTML/URL）、`analyze_template.py`（有模板时，含三层动画 inventory）；产物 `材料摘要.md`、`模板摘要.md` |
| ② 询问用户           | 本文件                                                                         | 无脚本；结果记入 `材料摘要.md`                                                                      |
| ③ 设置基本版式       | `references/design.md`                                                         | 完全自主模式跑 `sample_style.py`；可用 `style_showcase.py` 检查候选；模板模式继承主题与母页             |
| ④ 生成内容与微调版式 | `references/story.md` + `references/humanize.md` + `references/deck-schema.md` | 渲染器副本里的 `deck_check.py`；产物 `STORY.md`、`deck.json`                                        |
| ⑤ 排版渲染           | 完全遵循读 `references/template-inherit.md`；半自主再读 `references/template-semi-autonomous.md`；模板动画读 `references/template-animation.md`；启用拓展读 `extensions/<名>/EXT.md` | `template_mother.py`、渲染器副本 / `template_fill.py`、`formula.py`、拓展脚本；产物 `<中文名>.pptx` + `.animation-plan.json` |
| ⑥ 结果校验           | `references/validation.md`                                                     | `validate_pptx.py`、`render_preview.py` 出图自查；产物 `校验报告.md`                                |

每个 references 文档本身就是完整的设计规范，按要求执行即可。统计图表能力默认可用，但只有资料中存在真实数据关系时才生成图表；无真实数据时不凭空添加。动效、几何装饰、特殊排版等可选拓展默认关闭，说明在 `extensions/README.md`；用户没选就不用读。

## 六阶段工作流（严格串行）

总流水线：**阅读材料 → 询问用户 → 设置基本版式 → 生成内容与微调版式 → 排版渲染 → 结果校验**。
每个「⛔ 门禁」处停下，向用户确认或列出问题，确认后才继续。不跨阶段提前准备产物。

### 阶段一：阅读材料（按 references/source-intake.md）

1. 读取全部资料：Markdown 直接读；Word/PDF/Excel 用 Read 工具；HTML/URL 用 `source_to_md.py` 转 Markdown。
   转换失败则如实列缺口问用户。
2. 有模板时跑 `analyze_template.py <模板.pptx> -o 模板摘要.md` 并人工核对模板分析清单。清单必须包含 slide、slideLayout、slideMaster 三层的对象时序、构建范围、触发器、段落范围、高级特性和转场；看到 `playback_verified=false` 时不能把 XML 清单当成放映验证。
3. 判断领域、汇报场景、目标听众，按领域通行惯例组织内容。
4. ⛔ **信息缺口门**：术语、名称、单位、数字必须与资料完全一致；资料不足的部分列缺口问用户补齐
   或确认留白；**绝不臆造数据、结论、案例**。
5. 产物：`材料摘要.md`（事实清单、术语与数据摘录、模板摘要、缺口清单）。

### 阶段二：询问用户

⛔ 一次问完，只问缺的，能推断的不问：

1. **场合、受众、页数**：最多 3 个问题；用户已说明的不再问。
2. **模板风格模式**（给了模板才问，四选一）：
   - **完全遵循模板**：复制最接近的模板原页，只替换已有文字、图片、表格；必要时可把顶层图片位替换成同尺寸图片或 OfficeCLI 原生图表。
   - **半自主**：从模板正文页生成内容页母页，清除明确的原始内容，保留母版、背景、装饰、字体、主题色和页眉页脚；再用母页的组合设计原子复制扩展内容，并按模板风格新增图表或少量文本框。
   - **完全自主（默认白底极简）**：不继承模板页面，使用 `sample_style.py` 和渲染器生成白底极简 PPT。
   - **完全自主（其它风格，用户指定）**：不继承模板页面，但按用户指定的色彩、字体、密度或参考风格生成，并记录设计约束。
3. **是否快速模式**：先提示：完整模式会消耗更多 token 和时间，但可控性更强；快速模式会减少规划、候选和确认，适合先出初稿，交付后可能需要调整。选择是 → 按 `references/fast-mode.md` 走 fast-lite 路径，并跳过本阶段之后的提问与确认；选择否 → 走完整或普通 Quick 流程。
4. **是否允许复用参考资料里的图片**：否 → `figure` 不使用资料原图。
5. **是否启用可选拓展**：统计图表默认开启（仅有真实数据关系时生成）；动效、几何装饰、特殊排版默认关闭。说明：拓展会增加额外 token、时间和出错概率；几何装饰中的“底纹”还需单独明确允许。

用户说"直接做""快速"即视为快速模式。快速模式的默认值：有模板则遵循模板；允许复用图片；统计图表能力保持开启；动效、几何装饰、特殊排版和几何底纹均关闭。
把选择记入 `材料摘要.md`。

### 阶段三：设置基本版式

完全遵循模板和半自主模式跳过本阶段的配色采样，直接从模板主题色、字体与版式母页继承。两个完全自主模式才运行 `sample_style.py`；其中默认白底极简使用采样器默认规则，用户指定风格使用用户给出的颜色、字体和版式边界覆盖采样结果。快速模式先读 `references/fast-mode.md`：无模板只生成一个 seed 和一个候选；有模板只做一次紧凑 inventory 并复用结果。

1. 跑采样器，得到若干套配色 / 基本版式常量 / 逐页版式候选：
   ```
   <py> scripts/sample_style.py --seed <随机整数> --count 3 --scene <学术|工程|商务|政务|教育> --pages <页数或类型序列> [--force-color HEX]
   ```

   采样只提供随机性，**最终取舍归你**：觉得不合适可以换 seed、改色或手调；用户指定了颜色就用 `--force-color` 强制采用。
   普通流程取第 1 套至第 3 套比较；快速模式用 `--count 1`，直接取唯一候选并保留 seed 以便复现。
2. ⛔ 把选定的配色（色值）与版式风格告诉用户，没反对再往下（快速模式不停）。
3. 复制渲染器三件套到项目：`scripts/render_pptx.py`、`text_measure.py`、`deck_check.py` → `ppt-suggestion 的结果/渲染器/`。
   用 `sample_style.py --seed <种子> --emit-constants <候选序号>` 输出的常量块替换副本顶部的常量；还要改风格就改副本，**不改 skill 原件**。
4. 记下所用 seed 与候选序号，存为 `ppt-suggestion 的结果/style.json`，便于复现。

### 阶段四：生成内容与微调版式（按 story.md + humanize.md + deck-schema.md）

1. 完整或普通 Quick 按 `story.md` 写 `STORY.md`：意图（五句）→ 骨架（目录与章节一一对应、相邻页版式不同）→ 逐页。fast-lite 只保留一次紧凑事实清单，不生成完整 `STORY.md` 或设计锁定文件。
2. **去 AI 味（必经步骤）**：完整或普通 Quick 对 `STORY.md` 与 `deck.json` 的全部文本按 `humanize.md` 的**全量规则**逐条自查；fast-lite 直接对 `deck.json` 做同样的文本检查。
3. ⛔ **大纲确认门**：向用户列出每页一行结论句，没反对再往下（快速模式跳过；资料缺口改为交付时如实列出）。
4. 按 `deck-schema.md` 写 `deck.json`；每页 block 以采样出的版式候选为起点，按内容微调。快速模式仍运行 `deck_check.py`，保留公式、图表、越界、溢出和严重重叠等 Hard 校验。
   **渲染前跑渲染器副本里的校验**：`<py> 渲染器/deck_check.py deck.json`（用了特殊排版拓展加 `--extra-kinds bignum,asym,quote,diagonal,radial,ring,marginalia,concentric,quadrant`；完全遵循模板和半自主模式用 `scripts/deck_check.py deck.json --template-mode`，它只放宽无模板版式多样性门槛，不放宽标题、图表数据、公式和对象完整性要求），失败按错误解读表修复，禁止跳过。
   正文里的行内数学写成 `$LaTeX$`（如 `$n_l$`），字面美元符号写 `\$`；语法见 `deck-schema.md`。
5. 产物：`STORY.md`、`deck.json`。

### 阶段五：排版渲染

1. 无模板：
   ```
   <py> 渲染器/render_pptx.py --input deck.json --output "<中文文件名>.pptx"
   ```
2. 有模板（完全遵循）：按 `references/template-inherit.md` 写 `map.json`，再用 OfficeCLI 批量编辑模板副本：
   ```
   <py> scripts/template_fill.py --template 模板.pptx --deck deck.json --map map.json --output "<中文文件名>.pptx"
   ```
   半自主先按 `references/template-semi-autonomous.md` 写母页清单，生成并检查内容页母页：
   ```
   <py> scripts/template_mother.py --template 模板.pptx --manifest mothers.json --output mothers.pptx
   ```
   再在 `map.json` 中用 `clone_groups` 复制模板组合原子、填充组合内文字，并用 `template_fill.py` 生成最终页。母页只是规划和检查产物，最终输出仍从原模板复制，避免母页中间文件造成继承漂移。
   ```
   <py> scripts/template_fill.py --template 模板.pptx --deck deck.json --map map.json --output "<中文文件名>.pptx"
   ```
3. deck 含 `formula` block 或正文含 `$...$` 行内公式时，渲染后跑：`<py> scripts/formula.py "<中文文件名>.pptx"`，把占位框与行内标记换成 PowerPoint 原生可编辑公式。
4. 图表能力默认可用：有真实数据关系时使用 `chart` block。完全自主模式优先由 `power_pptx` 生成原生图表；完全遵循模板和半自主模式由 OfficeCLI 在模板副本中保留已有图表，或通过 `fields.chart` 同位替换顶层图片/图表。当前 `template_fill.py` 不支持任意坐标新增图表，也不自动执行 SVG 回退；需要时单独走明确的 OfficeCLI 编辑或 SVG 图片替换流程并校验。完全自主渲染器可显式用 `fallback: "svg"`，依赖 `matplotlib` 与 `cairosvg`。启用了其他拓展就按各自 `extensions/<名>/EXT.md` 处理；后处理类拓展在公式之后、结果校验之前运行。几何装饰的底纹只在用户明确允许且通过安全门槛时使用。
5. 渲染器内置 deck 校验，失败会非 0 退出并打印原因：按报错改 `deck.json` 再跑。渲染与校验是前台短任务，不要放后台。
6. 产物：`.pptx`（文字、原生图表、形状和原生公式可编辑；SVG/PNG 回退图表作为图片，需明确说明）。模板输出先在临时副本中编辑；先按动画计划处理对象时序与转场，再用 `repair_package(..., repair_animations=False)` 检查包引用和布局，成功后才替换输出。严格生成路径不会因悬空目标悄悄删除整页动画；计划写入 `.animation-plan.json`，最后必须运行 `template_animation.py` 复核。不能只以文字落位相等作为生成成功条件。

### 阶段六：结果校验（按 references/validation.md）

1. 机械校验：`<py> scripts/validate_pptx.py "<输出的.pptx>" --deck deck.json`（完全遵循模板和半自主模式加 `--baseline 模板.pptx`，忽略模板自带的越界）。模板图表替换还必须确认输出中存在原生 chart；母页必须确认只清除了 manifest 明确列出的内容。**Hard 问题必须为零**才交付；Soft 可接受（`--strict` 除外）。
2. 模型自查：先逐页渲染成图片再看（支持视觉时必做）：
   ```
   <py> scripts/render_preview.py "<输出的.pptx>" -o <系统临时目录>/ppt_preview --sheet
   ```

   用 Read 看每一页（含公式的页加 `--render native` 才能看到真实公式），再对照 `validation.md` 的 Hard/Soft 清单复核。看见文字伸出标签、卡片或文本框即阻断交付，即使脚本报告为零。小修限位置、字号（±2pt 内）、间距、对齐、底衬；仍装不下就退回阶段四或五，重选模板页、拆分内容或调整结构，再重渲染、重校验。预览图不交付。
3. 结构检查、HTML/原生静态目视检查、目标应用打开和实际放映分别记录。`template_animation.py` 只核对 XML 计划，`officecli validate` 与 HTML 预览不证明 Mac Office 无修复提示或动画播放；未做目标应用打开或放映时必须分别明确说明。资料缺口如实指出，不编造填充。
4. 交付：只给 `.pptx` **绝对路径**；如需说明假设、模板继承回退或缺口，简短列出。

## 纪律

- **串行**：阶段顺序不可跳、不可并行合并。
- **门禁即停**：⛔ 处等用户明确确认；大纲门与版式门以"用户没反对再往下"为准；快速模式除外。
- **就地修复**：内容问题改规划产物，版式问题改 `deck.json` 重渲染，渲染器问题改项目里的渲染器副本。
  每次修复重跑闭环：`deck_check.py` → 渲染 → `validate_pptx.py`。不推倒重来。
- **不虚构**：资料没有的数据、结论、案例一律不写；"显著/领先/国际先进"必须有资料依据。
- **像人做的**：最终效果应像领域专业人士手工制作，不是 AI 模板味。有疑问时先自查
  `references/humanize.md` 与 `references/design.md` 的禁止项。

## 产物与位置

- 中间文件与产物默认放在工作区 `ppt-suggestion 的结果/`：`材料摘要.md`、`模板摘要.md`、`STORY.md`、
  `style.json`、`渲染器/`、`deck.json`、`<中文名>.pptx`、`校验报告.md`。
- 页面管理优先依托母版和版式：有模板时继承模板版式，而不是每页自由发挥。
