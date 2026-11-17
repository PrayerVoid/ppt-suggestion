# 扩展索引

本文件负责能力入口、启用条件与执行顺序。图表能力默认可用，使用图表时阅读其 EXT；其余扩展须用户启用后才读。核心版式与图片构图见 [layout-types.md](../references/layout-types.md)。

| 能力 | 说明入口 | 启用条件 |
|---|---|---|
| 图表 | [charts/EXT.md](charts/EXT.md)：完全自主原生/图片回退，模板 OfficeCLI 同位替换及限制 | 默认可用，仅有真实数据关系时使用 |
| 动效 | [animation/EXT.md](animation/EXT.md)：成品页间转场与内容入场 | 用户明确选择动效 |
| 几何装饰 | [geometry-deco/EXT.md](geometry-deco/EXT.md)：3 个基础母题与 41 个纹样模板，原生可编辑形状 | 用户启用几何装饰；底纹另行明确允许 |
| 特殊排版 | [special-layout/EXT.md](special-layout/EXT.md)：diagonal、radial、ring、concentric 的渲染器接入 | 用户启用特殊排版 |

## 执行约定

- 特殊排版在渲染器副本中接入，图表在渲染或模板填充阶段完成。
- 后处理顺序为“渲染/模板填充 → formula（如有）→ 几何装饰（如启用）→ 动效（如启用）→ 最终校验与逐页预览”。后处理输出新文件，不覆盖输入 PPTX。
- 使用项目 `.venv` 解释器。装饰对象以 `DECO_` 命名；公式占位对象以 `FORMULA_` 命名，后续脚本据此前缀识别。
- 模板动画保留、清除与计划核对归 [template-animation.md](../references/template-animation.md)；新增动效需自己的目标/点击计划。所有形状后处理完成后复核适用计划，不能以静态结构通过声称实际播放已验证。
- 最终 Hard/Soft 检查及验证范围按 [validation.md](../references/validation.md)。
