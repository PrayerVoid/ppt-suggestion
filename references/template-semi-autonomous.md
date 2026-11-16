# 半自主模板模式

半自主模式把模板当作一组可复用的视觉原子。它保留模板的母版、背景、页眉页脚、主题色、字体和装饰，再从正文页中清除明确的原始内容，得到“内容页母页”；生成新页时复制母页对应的原始模板页，并复制其中的卡片、步骤或内容面板组合来扩展结构。

最终输出仍从原模板复制，`mothers.pptx` 只用于检查清空结果和帮助理解版式。这样可以避免先保存母页再二次保存造成主题、路径或字体继承漂移。

## 流程

1. 运行一次模板 inventory：

   ```text
   <py> scripts/analyze_template.py 模板.pptx -o 模板摘要.md --slides-json
   ```

2. 按内容签名选正文原页。先排除封面、目录、章节扉页和结束页，再考虑图片方向、文本块数量、条目数量和容量。模板“国潮敦煌”的推荐正文页如下：

   | 内容意图 | 原页 | 适用范围 |
   | --- | ---: | --- |
   | 一条判断 + 右侧横图 | 9 | 默认图文正文 |
   | 一段说明 + 右侧竖图 | 6 | 人物、文物、竖版海报 |
   | 左图右文 | 8 | 图片先行、解释跟随 |
   | 四项短模块 | 12 | 四个能力、观点或步骤 |
   | 六项短流程 | 17 | 4–6 个短环节 |
   | 双栏现状/目标 | 20 | 明确二分对比 |
   | 两段观点 + 高图 | 16 | 上下对照 |
   | 三图案例 | 23 | 三个同类案例各配一图 |

3. 写 `mothers.json`，只列出要清除的内容对象。不要用“删除所有文字”的启发式，因为章节标签、页眉、页码和装饰文字可能也是模板结构的一部分：

   ```json
   {
     "mothers": [
       {
         "name": "content-right-image",
         "template_slide": 9,
         "remove": ["文本框 35", "文本框 2", "图片 80"]
       },
       {
         "name": "content-four-cards",
         "template_slide": 12,
         "remove": ["组合 13", "组合 14", "组合 18", "组合 23"]
       },
       {
         "name": "content-two-columns",
         "template_slide": 20,
         "remove": ["组合 14", "组合 21"]
       }
     ]
   }
   ```

4. 生成母页并真实预览：

   ```text
   <py> scripts/template_mother.py --template 模板.pptx \
       --manifest mothers.json --output mothers.pptx
   officecli validate mothers.pptx
   <py> scripts/render_preview.py mothers.pptx -o <临时目录>/mothers --sheet --engine officecli --render html
   ```

   母页应保留模板背景、边框、章节标签、页脚和固定装饰；被列入 `remove` 的原始文字、图片或内容组合应全部消失。任何遗漏都回到 manifest 补充，不能在生成后随意删掉未确认的 logo、页码或页眉。

5. 用 `map.json` 填充母页对应的模板原页。半自主扩展优先使用 `clone_groups`，因为整组复制会保留模板的字体、颜色、圆角、边框、编号和段落样式：

   ```json
   {
     "slides": [
       {
         "template_slide": 12,
         "fields": {"组合 11/文本框 4": "新的四项能力"},
         "remove": ["组合 13", "组合 14", "组合 18", "组合 23"],
         "clone_groups": [
           {
             "source": "组合 13",
             "x": "2.6cm", "y": "4.2cm", "width": "6.4cm", "height": "8.1cm",
             "fields": {"组合 13/文本框 8": "内容一：保留模板卡片语言"}
           },
           {
             "source": "组合 13",
             "x": "9.6cm", "y": "4.2cm", "width": "6.4cm", "height": "8.1cm",
             "fields": {"组合 13/文本框 8": "内容二：复制后扩展结构"}
           }
         ]
       }
     ]
   }
   ```

   `source` 必须是模板页上的顶层组合。OfficeCLI 不能稳定地把嵌套 shape 直接 `add --from` 到页面根部；复制顶层 group 后，再通过 `fields` 修改组合内文字是当前最稳的路径。复制后的对象 ID 会重编，脚本内部使用复制后的同级位置处理，不要在 map 中写复制后的 ID。

6. 需要图表时，在顶层图片位或明确的图表位使用：

   ```json
   {
     "图片 80": {
       "chart": {
         "chart_type": "column_clustered",
         "categories": ["一", "二", "三"],
         "series": [{"name": "数量", "values": [12, 21, 17]}],
         "source": "资料中的统计表",
         "colors": "accent1,accent2"
       }
     }
   }
   ```

   `template_fill.py` 预先读取原位坐标，在一次 OfficeCLI batch 中添加同位 native chart 并删除原图片，再在输出校验中确认 chart 对象实际存在。数据必须来自资料，必须有来源或口径说明。复杂嵌套图片位、系列结构变化很大的既有图表，改用同尺寸 SVG/PNG 并记录回退原因。

7. 最终检查：

   `template_fill.py` 在 OfficeCLI 运行前生成替换文字容量与源图片清晰度门禁。章节标签、竖排文字和缩放后的组合按实际显示尺寸估算；固定装饰的 baseline 告警只可降低装饰自身的 Soft 噪声，不能豁免新增文字溢出、图片模糊或内容碰撞。母页生成后会写出 `<mothers.pptx>.text-evidence.json`，记录清空后实际仍存在的文字；母页 manifest 有遗漏时依据该文件回补，不凭旧模板摘要判断。

   ```text
   <py> scripts/template_fill.py --template 模板.pptx --deck deck.json \
       --map map.json --output 输出.pptx
   <py> scripts/validate_pptx.py 输出.pptx --deck deck.json --baseline 模板.pptx
   officecli validate 输出.pptx
   <py> scripts/render_preview.py 输出.pptx -o <临时目录>/preview --sheet --engine officecli --render html
   ```

## 选页硬约束

- 图片方向和比例必须匹配；没有合适图片时选第 17 页或第 20 页等无强制图片位母页。
- 四卡片、六行流程和三图案例的每个槽位都要控制字数，装不下就换母页或拆页，不靠压缩字号硬塞。
- 连续页面避免机械重复同一个母页；内容签名相同但证据方向不同，可以在 6、8、9、16 之间轮换。
- 只保留模板装饰和主题语言；新增文本框使用主题色，chart 配色先核对模板可见颜色与主题是否一致，不重新采样另一套配色。
- 清空母页时也清除原始章节文字，保留其底衬和装饰。生成母页前必须按 `references/template-animation.md` 检查三层动画；manifest 可写 `animation_policy`，删除对象、克隆组合和改变段落构建按计划处理。默认 `auto` 不猜测克隆对象的一对多点击顺序，会静态化受影响页；`preserve` 无法证明时在 OfficeCLI 前失败；共享母版/版式时序只有统一显式 `strip` 才能清除。输出的 `.animation-plan.json` 必须在公式、图表和其它后处理后复核。
- 章节标签要按容量使用短标题，完整结论放在足够宽的正文标题区；嵌套组合内文字必须参与越界检查。
- 半自主必须逐页看图；发现文字超出容器即退回修改，不能用机械检查为零覆盖肉眼问题。HTML 预览可能显示日期/页码占位，且不能验证 Office 的修复机制。Windows 可用 `--render native`；Mac 应用 PowerPoint 实际打开生成副本并导出或查看，未实测时说明范围。
