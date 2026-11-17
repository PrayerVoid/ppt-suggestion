# 拓展：几何装饰

仅当用户在“询问用户”阶段明确选了“几何装饰”，或明确说“炫酷”“要设计感”时才读本文件。默认外观仍是白底极简、无角落装饰；不启用就不要加。

## 处理方式

对已经渲染好的成品 PPTX 后处理，给整份 deck 增加一套可编辑的几何母题。脚本只使用 PowerPoint 原生形状，所有新增形状命名为 `DECO_<序号>`，放在每页底层；重复运行会先删掉已有 `DECO_` 形状。

一份 deck 只选一个母题。`--motif` 用于明确选型；`--seed` 用于在同一组母题中可复现地抽取一个。两者都不传时使用最克制的 `circles`。

```bash
python extensions/geometry-deco/add_deco.py 成品.pptx \
  --deck deck.json \
  --out 成品_装饰.pptx \
  --motif rosette \
  --accent 4F6082
```

`--accent` 必须与渲染器副本的 `ACCENT` 一致，不要在拓展里另起颜色。

## 模式与安全 profile

`--mode auto` 为默认值，使用角落装饰；`--mode corner` 显式选择角落装饰。只有明确传入 `--mode background` 才会尝试底纹，底纹不会由 `auto` 自动启用。

`--profile` 默认是 `formal`，只允许低密度 `circles` 或 `lines`，并限制形状复杂度、页面覆盖率和不透明度。`balanced` 放宽这些门槛，`expressive` 允许完整模板库，但仍受安全检查约束。示例：

```bash
python extensions/geometry-deco/add_deco.py 成品.pptx \
  --deck deck.json \
  --out 成品_安全底纹.pptx \
  --mode background \
  --profile balanced \
  --motif circles \
  --accent 4F6082
```

背景模式按页面类型生成候选边缘区域，并同时检查 deck 语义上的标题、正文、页脚保护区和 PPTX 中已有对象保护区。纹样生成后会重新读取新增形状的实际外接框，再复核越界、二次碰撞、覆盖率、复杂度和透明度；没有安全位置时跳过该页并打印原因。

## 可选母题

`--motif` 支持 41 个纹样模板与 3 个基础母题。模板函数统一使用 1280×720 坐标，并输出 PowerPoint 原生可编辑形状；下列 ID 可用于 `--motif`，最终效果仍须逐页检查。

| 家族 | 模板 ID |
|---|---|
| 基础母题 | `circles`、`blocks`、`lines` |
| 经典派生形 | `vesica`、`rosette`、`seed-of-life`、`moon`、`octagram`、`octagon`、`hexagram` |
| 弧与螺旋 | `seigaiha`、`pinwheel`、`scallop`、`shippo`、`truchet`、`reuleaux`、`laurel`、`cloud`、`ring-chain`、`guilloche`、`golden-spiral`、`rainbow`、`wave`、`square-whirl`、`moire` |
| 涌现场 | `phyllotaxis`、`halftone`、`tartan` |
| 东亚铺砌 | `kagome`、`asanoha`、`kikko`、`uroko` |
| 西亚/欧洲铺砌 | `tumbling-blocks`、`star-cross`、`ichimatsu`、`argyle`、`meander`、`egg-dart`、`herringbone`、`basket-weave` |
| 放射与中式 | `sunburst`、`fangsheng`、`bead-reel`、`qiulu` |

不传选型参数时使用 `circles`；`--seed` 从完整 ID 集合可复现选择，`--motif <ID>` 固定指定模板。母题选择不改变白底极简、单强调色与低干扰规则；背景模式还受 profile 限制。

## 页型规则

| 页面类型 | 装饰处理 |
|---|---|
| `cover` / `close` | 大构图，优先贴右上或右下；封面已有内容占据角落时自动缩小或跳过 |
| `section` | 同一母题的大构图，换到另一角落，作为章节锚点 |
| `catalog` / `content` | 同一母题的小角标，固定位置、大小和颜色 |
| 启用 `background` 模式 | 只在通过内容保护区、对象保护区和安全门槛的边缘区域放置底纹；无安全位置则跳过 |

角落模式先读取每页所有非装饰形状（包括组合内子形状）的位置，再为同类页面寻找共同的角落和缩放；装饰与已有内容至少保留 `0.12` 英寸间距。背景模式按页寻找安全区域，并在生成后按实际形状外接框重新检查。旋转形状按其外接框做保守避让。

## 视觉纪律

- 全稿只用一个母题；不要每页换图案。
- 大构图只放在一侧，内容页的小角标不能与标题、页码、页脚争夺焦点。
- 装饰色必须来自正文强调色；浅色块由脚本自动取强调色的浅色混合，不新增颜色。
- 装饰放在最底层，不遮挡文字、图片主体或模板页眉页脚。
- 复杂纹样只在内容确实需要时使用；商务、政务和工程汇报优先 `circles`、`blocks`、`lines`。
- 后处理顺序固定为：渲染 → `formula.py` → 几何装饰 → 动效 → `validate_pptx.py` → `render_preview.py`。

## 结果检查

`validate_pptx.py` 会忽略 `DECO_` 形状的重叠与文字溢出检查，但越界检查仍然生效。交付前仍要逐页渲染，重点看：

1. 眯眼时先看到标题和主体，而不是装饰。
2. 目录、正文和页脚的固定内容没有被装饰压住。
3. 同一母题在封面、章节和内容页的尺度变化有规律。
4. 复杂母题没有造成密集纹理或新的视觉焦点。
