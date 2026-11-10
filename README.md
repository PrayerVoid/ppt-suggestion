# ppt-suggestion

基于用户提供的资料（文字 / Word / PDF / 图片 / 表格 / HTML / URL）与可选 PPT 模板，生成**专业、克制、白底极简、可在 PowerPoint/WPS 二次编辑**的正式汇报 PPT 的 Comate 项目技能。

覆盖学术科研（答辩 / 课题申报 / 结题）、工程技术（方案评审 / 架构说明 / 复盘）、商务咨询（工作汇报 / 经营分析 / 方案提案）、政务公共（政策解读 / 工作总结）、教育培训（课件 / 科普）等场景。

## 核心目标

- 最终效果像该领域专业人士手工制作，**没有明显的 AI 生成痕迹**（视觉与文案双重去 AI 味）。
- 术语、名称、单位、数字与原始资料**完全一致**，不虚构、不夸大。
- 有模板时**严格继承模板**的设计语言（页面比例 / 母版版式 / 配色 / 装饰），而不是另起炉灶。
- 输出全部为**原生可编辑对象**，可直接在 PowerPoint / WPS 中改字、改表、改图。

## 工作流（四阶段，严格串行）

```
① 阅读材料 ──► ② 生成内容规划 ──► ③ 排版渲染 ──► ④ 结果校验
   素材转换        叙事(去AI味)       白底极简/模板继承    机械+人工复核
   模板分析        大纲确认         deck_check       修复回路
```


| 阶段            | 做什么                                        | 产物                         | 门禁                         |
| ----------------- | ----------------------------------------------- | ------------------------------ | ------------------------------ |
| ① 阅读材料     | 转 Markdown、判断领域、模板分析、充分性检查   | `材料摘要.md`、`模板摘要.md` | 资料缺口先问用户，绝不臆造   |
| ② 生成内容规划 | 意图→骨架→逐页结论句；全量去 AI 味改写      | `STORY.md`、`deck.json`      | 大纲确认（用户没反对再往下） |
| ③ 排版渲染     | `deck_check.py` 校验 → `render_pptx.py` 渲染 | `<中文名>.pptx`              | 校验失败禁止跳过             |
| ④ 结果校验     | `validate_pptx.py` + 逐页复核清单             | `校验报告.md`                | 修复闭环后交付               |

## 目录结构

```
ppt-suggestion/
├── README.md                # 本文件：skill 使用说明
├── SKILL.md                 # 工作流总控：四阶段流程、门禁、纪律、文件导航
├── references/
│   ├── source-intake.md     # 阶段一：阅读材料
│   ├── story.md             # 阶段二-叙事：意图/骨架/逐页/版式选用
│   ├── humanize.md          # 阶段二-去 AI 味：全量规则（24 类模式 + 清单）
│   ├── deck-schema.md       # deck.json 结构说明与质量检查清单
│   ├── design.md            # 视觉/版面/密度边界
│   ├── template-inherit.md  # 模板继承规则与渲染器模板模式
│   └── validation.md        # 阶段四：结果校验（Hard/Soft 分级）
└── scripts/
    ├── source_to_md.py      # HTML/URL/文本 → Markdown
    ├── analyze_template.py  # 模板分析 → 模板摘要
    ├── deck_check.py        # deck.json 结构校验（渲染前）
    ├── render_pptx.py       # 渲染器（无模板白底 / 有模板继承）
    ├── validate_pptx.py     # 输出校验（重叠/越界/溢出/图片/版式）
    └── text_measure.py      # 文本估算（renderer/validator 共用）
```

## 快速上手

```bash
# 1. 阅读材料（HTML 源文件转 Markdown）
python <venv>/bin/python .comate/skills/ppt-suggestion/scripts/source_to_md.py 参考资料.html -o 材料.md

# 2. 有模板时先分析模板
python <venv>/bin/python .comate/skills/ppt-suggestion/scripts/analyze_template.py 模板.pptx -o 模板摘要.md

# 3. 渲染前校验 deck.json
python <venv>/bin/python .comate/skills/ppt-suggestion/scripts/deck_check.py deck.json

# 4. 渲染（无模板白底；有模板加 --template）
python <venv>/bin/python .comate/skills/ppt-suggestion/scripts/render_pptx.py --input deck.json --output 汇报.pptx
python <venv>/bin/python .comate/skills/ppt-suggestion/scripts/render_pptx.py --input deck.json --output 汇报.pptx --template 模板.pptx

# 5. 校验产出
python <venv>/bin/python .comate/skills/ppt-suggestion/scripts/validate_pptx.py 汇报.pptx --deck deck.json
```

本机渲染请用专用 venv 的 python（系统 python3 受 PEP 668 管理）：
`/Users/vow/.comate-engine/pptx-venv/bin/python`。渲染/校验是前台短任务，不要放后台。

## 定制指南

- **改设计规范**：按需调整 `references/design.md`（视觉/密度）、`references/template-inherit.md`（模板继承）、`references/source-intake.md`（阅读材料）等。
- **改视觉**：无模板的配色/字号/间距全部集中在 `scripts/render_pptx.py` 顶部常量，JSON 里不允许出现颜色字段——要改风格就改这里。
- **改校验强度**：`validate_pptx.py` 用 `--strict` 让警告也失败；调整阈值在脚本内常量。
- **新增版式**：在 `render_pptx.py` 的 `BLOCK_RENDERERS` 里加一个渲染函数，并在 `deck-schema.md` 登记。
