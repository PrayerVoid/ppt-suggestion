"""把 sample_style.py 的每个候选渲染成一页可比较的风格展示 PPTX。

用法:
  python style_showcase.py --seed 20261116 --count 5 --scene 商务 \
      --pages 12 --output style-showcase.pptx

每页展示同一候选的色板、文字层级、基础参数和一个代表性内容块，
用于检查随机采样结果；它不替代正式 deck 渲染。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from sample_style import MIN_TEXT, contrast, make_candidate, page_types  # noqa: E402

try:
    from power_pptx import Presentation
    from power_pptx.dml.color import RGBColor
    from power_pptx.enum.shapes import MSO_SHAPE
    from power_pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from power_pptx.util import Inches, Pt
except ImportError as exc:  # pragma: no cover - dependency guard
    raise SystemExit(f"缺少 power-pptx，请先安装 requirements.txt: {exc}")

PAGE_W, PAGE_H = 13.333333, 7.5
FONT = "微软雅黑"
PALETTE_KEYS = ("ACCENT", "ACCENT_DK", "ACCENT_LT", "PANEL", "LINE", "TEXT", "SUBTLE", "FAINT")


def rgb(hex_value):
    value = str(hex_value).lstrip("#")
    return RGBColor.from_hex(value.upper())


def set_run(run, size, color, bold=False, italic=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def add_text(slide, text, x, y, w, h, size, color, bold=False, align=PP_ALIGN.LEFT):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    shape.shadow.clear()
    shape.text_frame.clear()
    shape.text_frame.margin_left = shape.text_frame.margin_right = 0
    shape.text_frame.margin_top = shape.text_frame.margin_bottom = 0
    shape.text_frame.word_wrap = True
    shape.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    paragraph = shape.text_frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = str(text)
    set_run(run, size, color, bold)
    return shape


def add_box(slide, x, y, w, h, fill, line=None, radius=False):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    # 默认形状样式的 effectRef 会继承主题阴影；只清空局部 effectLst 不够。
    shape.shadow.clear()
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = line or fill
    shape.line.width = Pt(1)
    return shape


def draw_content_sample(slide, candidate, palette, x=0.85, y=4.55, w=11.65, h=2.15):
    layout = next((item for item in candidate["layout"] if item["type"] == "content"), None)
    kinds = [block["kind"] for block in layout.get("blocks", [])] if layout else ["lead", "points"]
    accent = rgb(palette["ACCENT"])
    dark = rgb(palette["ACCENT_DK"])
    panel = rgb(palette["PANEL"])
    line = rgb(palette["LINE"])
    text = rgb(palette["TEXT"])
    subtle = rgb(palette["SUBTLE"])
    add_text(slide, "代表性内容块", x, y - 0.25, 2.2, 0.22, 10, subtle, True)
    if any(kind in {"cards", "columns", "stats"} for kind in kinds):
        gap = 0.18
        card_w = (w - gap * 2) / 3
        for i, label in enumerate(("结论", "证据", "行动")):
            xx = x + i * (card_w + gap)
            add_box(slide, xx, y, card_w, h, panel, line)
            add_box(slide, xx, y, 0.08, h, accent, accent, False)
            add_text(slide, f"0{i + 1}", xx + 0.25, y + 0.22, 0.45, 0.3, 12, dark, True)
            add_text(slide, label, xx + 0.25, y + 0.68, card_w - 0.45, 0.34, 16, text, True)
            add_text(slide, "用一条短句说明这个内容块的作用。", xx + 0.25, y + 1.18, card_w - 0.45, 0.5, 10, subtle)
    elif any(kind in {"compare", "split", "rows", "table"} for kind in kinds):
        add_box(slide, x, y, w * 0.48, h, panel, line)
        add_box(slide, x + w * 0.52, y, w * 0.48, h, panel, line)
        add_text(slide, "现状", x + 0.28, y + 0.25, 2, 0.35, 16, dark, True)
        add_text(slide, "目标", x + w * 0.52 + 0.28, y + 0.25, 2, 0.35, 16, dark, True)
        add_text(slide, "信息密度、节奏和风险被放在同一条阅读路径中。", x + 0.28, y + 0.85, 4.4, 0.65, 11, text)
        add_text(slide, "用清晰的动作项把结论转成下一步安排。", x + w * 0.52 + 0.28, y + 0.85, 4.4, 0.65, 11, text)
    else:
        add_box(slide, x, y, w * 0.67, h, panel, line)
        add_box(slide, x + w * 0.72, y, w * 0.28, h, accent, accent)
        add_text(slide, "一条结论句放在视觉起点", x + 0.28, y + 0.28, 5.8, 0.4, 18, text, True)
        add_text(slide, "下面用两行解释说明边界、依据和下一步，保持段落短而可读。", x + 0.28, y + 0.93, 5.9, 0.55, 11, subtle)
        ink = palette["WHITE"] if contrast(palette["WHITE"], palette["ACCENT"]) >= MIN_TEXT else "000000"
        add_text(slide, "强调\n色块", x + w * 0.72, y + 0.52, w * 0.28, 0.85, 18, rgb(ink), True, PP_ALIGN.CENTER)


def render_candidate(prs, candidate, scene):
    palette = candidate["palette"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    white = rgb(palette["WHITE"])
    text = rgb(palette["TEXT"])
    subtle = rgb(palette["SUBTLE"])
    accent = rgb(palette["ACCENT"])
    dark = rgb(palette["ACCENT_DK"])
    background = slide.background.fill
    background.solid()
    background.fore_color.rgb = white

    add_box(slide, 0, 0, PAGE_W, 0.12, accent, accent, False)
    add_text(slide, f"STYLE SHOWCASE  /  {scene or '默认场景'}  /  CANDIDATE {candidate['id'] + 1:02d}",
             0.85, 0.34, 6.5, 0.25, 9, subtle, True)
    add_text(slide, f"风格候选 {candidate['id'] + 1:02d}", 0.85, 0.72, 6.4, 0.55, 27, text, True)
    add_text(slide, "先看系统参数，再看一个内容块样例；页面只用于比较采样结果。",
             0.85, 1.32, 6.6, 0.28, 11, subtle)

    # 色板
    add_text(slide, "PALETTE", 0.85, 1.88, 1.4, 0.25, 10, subtle, True)
    swatch_h, swatch_gap = 0.55, 0.12
    swatch_w = (11.65 - (len(PALETTE_KEYS) - 1) * swatch_gap) / len(PALETTE_KEYS)
    for i, key in enumerate(PALETTE_KEYS):
        x = 0.85 + i * (swatch_w + swatch_gap)
        add_box(slide, x, 2.18, swatch_w, swatch_h, rgb(palette[key]), rgb(palette[key]), False)
        add_text(slide, key, x, 2.82, swatch_w, 0.2, 8, subtle, False, PP_ALIGN.CENTER)
        add_text(slide, f"#{palette[key]}", x, 3.06, swatch_w, 0.2, 8, text, False, PP_ALIGN.CENTER)

    # 参数独占色板下方一行，避免遮住最后一枚色板及其标签。
    base = candidate["base"]
    summary = candidate["summary"]
    metrics = [
        ("margin", f"{base['MARGIN_L']:.2f} in"),
        ("gap", f"{base['BLOCK_GAP']:.2f} in"),
        ("title", f"{base['TITLE_SIZE']} pt"),
        ("asym", f"{summary['asymmetric_ratio']:.0%}" if summary['asymmetric_ratio'] is not None else "—"),
        ("insight", str(summary['insight_count'])),
    ]
    for i, (label, value) in enumerate(metrics):
        xx = 0.85 + i * 1.65
        add_text(slide, label, xx, 3.55, 1.4, 0.2, 9, subtle)
        add_text(slide, value, xx, 3.84, 1.4, 0.22, 10, text, True)
    notices = summary["unmet_soft_targets"] + summary["warnings"]
    warning = "；".join(notices) if notices else "通过软目标"
    add_text(slide, warning, 9.3, 3.55, 3.2, 0.55, 9, dark, True)
    draw_content_sample(slide, candidate, palette)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--seed", default="20261116")
    ap.add_argument("--count", type=int, default=5)
    ap.add_argument("--scene", default="")
    ap.add_argument("--force-color", default=None)
    ap.add_argument("--pages", default="12")
    ap.add_argument("--figure", action="store_true")
    ap.add_argument("--output", required=True)
    args = ap.parse_args()
    if args.count < 1:
        raise SystemExit("--count 必须大于 0")
    forced = args.force_color.lstrip("#").upper() if args.force_color else None
    if forced and (len(forced) != 6 or any(char not in "0123456789ABCDEF" for char in forced)):
        raise SystemExit("--force-color 需要 6 位十六进制")
    types = page_types(args.pages)
    candidates = [make_candidate(args.seed, i, args.scene, forced, types, args.figure)
                  for i in range(args.count)]
    prs = Presentation()
    prs.slide_width = Inches(PAGE_W)
    prs.slide_height = Inches(PAGE_H)
    for candidate in candidates:
        render_candidate(prs, candidate, args.scene)
    output = Path(args.output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output))
    from pptx_package import repair_package
    repair_package(output)
    print(output)


if __name__ == "__main__":
    main()
