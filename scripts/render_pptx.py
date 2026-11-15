#!/usr/bin/env python3
"""Render a formal, restrained, white-based PPTX from deck.json.

Visual rules live HERE, not in deck.json: white background, low-saturation
single accent, native editable shapes. 本脚本是"可被复制修改的模板"：
AI 在设置基本版式阶段把它复制到项目里，只改下方常量块（配色/字体/边距）。

Usage:
  python render_pptx.py --input deck.json --output out.pptx

若 deck 含 formula block，会在 out.pptx 旁写 out.pptx.formulas.json（公式占位符坐标），
由 formula.py 替换为原生可编辑公式。

Exit codes:
  0  success (prints absolute output path)
  1  validation error (deck.json violates design rules)
  2  deck.json unreadable / malformed
  3  power-pptx missing
"""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from text_measure import estimate_height as text_height  # noqa: E402
from deck_check import validate_deck  # noqa: E402

try:
    from power_pptx import Presentation
    from power_pptx.dml.color import RGBColor
    from power_pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
    from power_pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from power_pptx.oxml.ns import qn
    from power_pptx.util import Emu, Inches, Pt
except ImportError:
    sys.stderr.write("缺少 power-pptx，请先运行: pip install -r requirements.txt\n")
    raise SystemExit(3)

# ---------------------------------------------------------------------------
# 风格常量块：复制本脚本后，只改这里（配色 / 字体 / 页边距）
# ---------------------------------------------------------------------------
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT = RGBColor(0x26, 0x26, 0x26)          # near-black for body/titles
SUBTLE = RGBColor(0x6E, 0x6E, 0x6E)        # secondary text
FAINT = RGBColor(0x9A, 0x9A, 0x9A)         # footers / captions
ACCENT = RGBColor(0x4F, 0x60, 0x82)        # slate blue: desaturated, has G
ACCENT_DK = RGBColor(0x3A, 0x47, 0x5F)
ACCENT_LT = RGBColor(0xE9, 0xEC, 0xF2)     # light tint for panels
LINE = RGBColor(0xD9, 0xDD, 0xE3)          # separators
PANEL = RGBColor(0xF3, 0xF5, 0xF8)         # light panels

FONT = "微软雅黑"
PAGE_W, PAGE_H = 13.333333, 7.5             # 16:9 canvas
MARGIN_L, MARGIN_R = 0.9, 0.9
CONTENT_TOP = 1.55
CONTENT_BOTTOM = 7.05
BLOCK_GAP = 0.22
TITLE_SIZE = 25                             # 内容页标题字号

def _set_font(run, size=14, bold=False, color=TEXT, italic=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = FONT
    from lxml import etree
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", FONT)


def no_shadow(shape) -> None:
    from lxml import etree
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst"
    for node in shape._element.findall(f".//{ns}"):
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
    # 形状样式的 effectRef 会引用主题阴影；clear() 会同时改写 effectRef，才能真正关掉
    try:
        shape.shadow.clear()
    except Exception:
        pass


def add_rect(slide, l, t, w, h, fill, line=None, line_w=0.75, shape=MSO_SHAPE.RECTANGLE):
    sh = slide.shapes.add_shape(shape, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(line_w)
    no_shadow(sh)
    return sh


def add_hline(slide, l, r, y, color=LINE, w=0.75):
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(l), Inches(y), Inches(r), Inches(y))
    conn.line.color.rgb = color
    conn.line.width = Pt(w)
    no_shadow(conn)
    return conn


def add_text(slide, l, t, w, h, text, size=14, bold=False, color=TEXT,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=1.15,
             wrap=True, auto_size=None):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = line
        _set_font(r, size=size, bold=bold, color=color)
    return tb


# text_height 来自 text_measure.py（renderer 与 validator 共用同一把尺子）


# ---------------------------------------------------------------------------
# deck.json structural validation is delegated to deck_check.validate_deck()
# (imported above) and runs as a hard gate in main().
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# No-template rendering
# ---------------------------------------------------------------------------
def new_presentation():
    prs = Presentation()
    prs.slide_width = Emu(int(Inches(PAGE_W)))
    prs.slide_height = Emu(int(Inches(PAGE_H)))
    blank = prs.slide_layouts[6]
    return prs, blank


def render_cover(slide, s):
    kicker = s.get("kicker", "")
    title = s.get("title", "")
    subtitle = s.get("subtitle", "")
    footer = s.get("footer", "")
    if kicker:
        add_text(slide, MARGIN_L, 1.7, 11.5, 0.4, kicker, size=13, color=ACCENT_DK)
    size = 40
    if len(title) > 18:
        size = 34
    if len(title) > 26:
        size = 30
    add_text(slide, MARGIN_L, 2.1, 11.5, 1.6, title, size=size, bold=True, color=TEXT,
             line_spacing=1.05)
    add_hline(slide, MARGIN_L, MARGIN_L + 1.4, 3.75, color=ACCENT, w=1.5)
    if subtitle:
        add_text(slide, MARGIN_L, 4.0, 11.0, 0.6, subtitle, size=17, color=SUBTLE)
    if footer:
        add_text(slide, MARGIN_L, 6.55, 11.5, 0.4, footer, size=12, color=FAINT)
    add_hline(slide, MARGIN_L, PAGE_W - MARGIN_R, 6.95, color=LINE, w=0.75)


def render_catalog(slide, s):
    title = s.get("title", "目录")
    add_text(slide, MARGIN_L, 0.55, 10, 0.55, title, size=26, bold=True)
    add_hline(slide, MARGIN_L, MARGIN_L + 1.2, 1.25, color=ACCENT, w=1.5)
    y = 1.85
    items = []
    for b in s.get("blocks", []):
        if b.get("kind") == "agenda":
            items = b.get("items", [])
    for i, it in enumerate(items):
        num = it.get("num", f"{i + 1:02d}")
        t = it.get("title", "")
        note = it.get("note", "")
        add_text(slide, MARGIN_L + 0.05, y, 0.85, 0.6, num, size=24, bold=True, color=ACCENT)
        add_text(slide, MARGIN_L + 1.0, y + 0.06, 6.4, 0.6, t, size=17, bold=True)
        if note:
            add_text(slide, MARGIN_L + 7.6, y + 0.12, 4.0, 0.5, note, size=13, color=SUBTLE,
                     align=PP_ALIGN.RIGHT)
        add_hline(slide, MARGIN_L, PAGE_W - MARGIN_R, y + 0.82, color=LINE, w=0.5)
        y += 1.02


def render_section(slide, s, num, title, subtitle):
    if num:
        add_text(slide, MARGIN_L, 1.6, 2.4, 1.55, num, size=72, bold=True, color=ACCENT_LT)
    add_text(slide, MARGIN_L, 2.9, 11.0, 0.9, title, size=34, bold=True)
    if subtitle:
        add_text(slide, MARGIN_L, 4.0, 10.5, 0.5, subtitle, size=16, color=SUBTLE)
    add_hline(slide, MARGIN_L, MARGIN_L + 1.6, 4.75, color=ACCENT, w=1.5)


def content_area():
    return MARGIN_L, CONTENT_TOP, PAGE_W - MARGIN_L - MARGIN_R, CONTENT_BOTTOM - CONTENT_TOP


def render_content(slide, s, style):
    title = s.get("title", "")
    chapter = s.get("chapter", "")
    add_text(slide, MARGIN_L, 0.5, 10.8, 0.6, title, size=TITLE_SIZE, bold=True)
    if chapter:
        add_text(slide, PAGE_W - MARGIN_R - 3.4, 0.2, 3.4, 0.3, chapter, size=11,
                 color=FAINT, align=PP_ALIGN.RIGHT)
    add_hline(slide, MARGIN_L, PAGE_W - MARGIN_R, 1.32, color=LINE, w=0.75)
    x, y0, w, h = content_area()
    y = y0
    first_idx = len(slide.shapes)
    for b in s.get("blocks", []):
        if not isinstance(b, dict):
            continue
        bh = render_block(slide, x, y, w, b, style)
        if y + bh > CONTENT_BOTTOM + 0.03:
            raise RuntimeError(f"内容超出页面底部: 「{title}」blocks 排布过高")
        y += bh + BLOCK_GAP
    # 内容偏少时整体下移，让上下留白更均匀（最多下移 1.0in）
    total = y - BLOCK_GAP - y0
    avail = CONTENT_BOTTOM - y0
    if 0 < total < avail * 0.85:
        dy = Inches(min((avail - total) / 2, 0.8))
        for sh in list(slide.shapes)[first_idx:]:
            sh.top = sh.top + dy
    # footer: source + page number
    src = s.get("source") or style.get("source", "")
    if src:
        add_text(slide, MARGIN_L, 7.08, 7.0, 0.3, src, size=10, color=FAINT)


def render_block(slide, x, y, w, b, style):
    kind = b.get("kind")
    fn = BLOCK_RENDERERS.get(kind)
    if fn is None:
        raise RuntimeError(f"未知 block kind: {kind}")
    return fn(slide, x, y, w, b, style)


def _panel(slide, x, y, w, h, fill=PANEL):
    add_rect(slide, x, y, w, h, fill)


def _insight(slide, x, y, w, b, style):
    h = 0.62
    add_rect(slide, x, y, w, h, PANEL)
    add_rect(slide, x, y, 0.07, h, ACCENT)
    kicker = b.get("kicker", "结论")
    text = b.get("text", "")
    tx = x + 0.25
    add_text(slide, tx, y + 0.08, 1.2, 0.4, kicker, size=12, bold=True, color=ACCENT_DK)
    add_text(slide, tx + 1.2, y + 0.1, w - 1.6, h - 0.2, text, size=14, bold=True, color=TEXT)
    return h


def _lead(slide, x, y, w, b, style):
    h = 0.62
    add_text(slide, x, y + 0.05, w, h, b.get("text", ""), size=15, color=SUBTLE)
    return h


def _stats(slide, x, y, w, b, style):
    items = b.get("items", [])
    n = len(items)
    if n < 2 or n > 4:
        n = 3
    gap = 0.25
    cw = (w - gap * (n - 1)) / n
    tw = cw - 0.44
    # 标签、说明会换行：按估算高度逐段排布，卡片取各卡所需高度的最大值，文字才不会压到下一个块
    lhs = [text_height(it.get("label", ""), 13, tw) for it in items]
    nhs = [text_height(it.get("note", ""), 11, tw, 1.1) for it in items]
    h = max([1.75] + [0.95 + lh + 0.08 + nh + 0.2 for lh, nh in zip(lhs, nhs)])
    for i, it in enumerate(items):
        cx = x + i * (cw + gap)
        add_rect(slide, cx, y, cw, h, PANEL)
        add_rect(slide, cx, y, cw, 0.05, ACCENT)
        add_text(slide, cx + 0.22, y + 0.28, tw, 0.6, str(it.get("value", "")),
                 size=30, bold=True, color=ACCENT_DK, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, cx + 0.22, y + 0.95, tw, lhs[i] + 0.05, it.get("label", ""),
                 size=13, bold=True, color=TEXT)
        add_text(slide, cx + 0.22, y + 0.95 + lhs[i] + 0.08, tw, nhs[i] + 0.05,
                 it.get("note", ""), size=11, color=SUBTLE, line_spacing=1.1)
    return h


def _columns(slide, x, y, w, b, style):
    items = b.get("items", [])
    n = min(len(items), 3)
    if n == 0:
        return 0
    gap = 0.3
    cw = (w - gap * (n - 1)) / n
    maxh = 0.0
    for i, it in enumerate(items):
        cx = x + i * (cw + gap)
        icon = it.get("icon") or f"{i + 1:02d}"
        add_text(slide, cx, y, 0.8, 0.4, icon, size=13, bold=True, color=ACCENT)
        kicker = it.get("kicker", "")
        if kicker:
            add_text(slide, cx, y + 0.32, cw, 0.3, kicker, size=11, bold=True, color=FAINT)
        t = it.get("title", "")
        if t:
            add_text(slide, cx, y + 0.62, cw, 0.4, t, size=17, bold=True)
        body = it.get("body", "")
        if body:
            bh = text_height(body, 15, cw)
            add_text(slide, cx, y + 1.15, cw, bh + 0.35, body, size=15, color=SUBTLE, line_spacing=1.25)
            maxh = max(maxh, 1.15 + bh + 0.1)
    return max(maxh, 1.35)


def _steps(slide, x, y, w, b, style):
    items = b.get("items", [])
    n = len(items)
    if n == 0:
        return 0
    n = min(n, 5)
    gap = 0.55
    cw = (w - gap * (n - 1)) / n
    h = 1.15
    for it in items:
        bb = it.get("body", "")
        if bb:
            h = max(h, 0.75 + text_height(bb, 11, cw - 0.3) + 0.15)
    for i, it in enumerate(items):
        cx = x + i * (cw + gap)
        add_rect(slide, cx, y, cw, h, PANEL)
        num = it.get("num", f"{i + 1:02d}")
        add_text(slide, cx + 0.15, y + 0.1, 0.7, 0.3, num, size=12, bold=True, color=ACCENT)
        add_text(slide, cx + 0.15, y + 0.4, cw - 0.3, 0.35, it.get("title", ""), size=13, bold=True)
        body = it.get("body", "")
        if body:
            add_text(slide, cx + 0.15, y + 0.72, cw - 0.3, h - 0.8, body, size=11, color=SUBTLE,
                     line_spacing=1.1)
        if i < n - 1:
            add_rect(slide, cx + cw + 0.08, y + h / 2 - 0.11, gap - 0.16, 0.22, LINE,
                     shape=MSO_SHAPE.RIGHT_ARROW)
    return h


def _chips(slide, x, y, w, b, style):
    items = b.get("items", [])
    if not items:
        return 0
    h = 0.42
    cx = x
    maxy = y
    for it in items:
        tw = min(1.9, 0.5 + len(str(it)) * 0.16)
        if cx + tw > x + w:
            cx = x
            maxy = y + h + 0.12
            y = maxy
        add_rect(slide, cx, y, tw, h, PANEL)
        add_text(slide, cx, y + 0.07, tw, h - 0.1, str(it), size=12, color=ACCENT_DK,
                 align=PP_ALIGN.CENTER)
        cx += tw + 0.18
    return (maxy - y) + h


def _layers(slide, x, y, w, b, style):
    items = b.get("items", [])
    n = len(items)
    if n == 0:
        return 0
    arrow_gap = 0.24
    available = CONTENT_BOTTOM - y
    layer_h = (available - arrow_gap * (n - 1)) / n
    layer_h = min(1.35, max(0.8, layer_h))
    h = 0.0
    for i, it in enumerate(items):
        ly = y + h
        add_rect(slide, x, ly, w, layer_h, PANEL)
        add_text(slide, x + 0.25, ly + 0.1, w - 0.5, 0.35, it.get("kicker", ""), size=12,
                 bold=True, color=ACCENT_DK)
        t = it.get("title", "")
        if t:
            add_text(slide, x + 0.25, ly + 0.38, w - 0.5, 0.3, t, size=13, bold=True)
        modules = it.get("modules", [])
        if modules:
            mw = 1.9
            mgap = 0.18
            mh = 0.5
            mx = x + 0.25
            my = ly + layer_h - mh - 0.15
            for m in modules:
                if mx + mw > x + w - 0.2:
                    break
                add_rect(slide, mx, my, mw, mh, WHITE, line=LINE, line_w=0.75)
                add_text(slide, mx + 0.1, my + 0.04, mw - 0.2, 0.3, m.get("title", ""), size=11,
                         bold=True, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
                body = m.get("body", "")
                if body:
                    add_text(slide, mx + 0.1, my + 0.28, mw - 0.2, 0.2, body, size=8, color=SUBTLE)
                mx += mw + mgap
        if i < n - 1:
            add_rect(slide, x + w / 2 - 0.09, ly + layer_h + 0.02, 0.18, 0.2, FAINT,
                     shape=MSO_SHAPE.DOWN_ARROW)
        h += layer_h + (arrow_gap if i < n - 1 else 0)
    return h


def _table(slide, x, y, w, b, style):
    headers = b.get("headers", [])
    rows = b.get("rows", [])
    ncols = len(headers)
    nrows = len(rows)
    if ncols == 0 or nrows == 0:
        return 0.5
    cw = w / ncols
    rhs = [0.5] + [
        max(0.5, max(text_height(str(c), 12, cw - 0.25) for c in row) + 0.22) for row in rows
    ]
    h = sum(rhs)
    tbl = slide.shapes.add_table(nrows + 1, ncols, Inches(x), Inches(y), Inches(w), Inches(h)).table
    tbl.first_row = True
    tbl.horz_banding = False
    for ri_, rh_ in enumerate(rhs):
        tbl.rows[ri_].height = Inches(rh_)
    for ci, hd in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = PANEL
        cell.margin_top = cell.margin_bottom = Inches(0.04)
        p = cell.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = str(hd)
        _set_font(r, size=13, bold=True, color=TEXT)
    for ri, row in enumerate(rows):
        for ci in range(ncols):
            cell = tbl.cell(ri + 1, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if ri % 2 == 0 else PANEL
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            p = cell.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = str(row[ci]) if ci < len(row) else ""
            _set_font(r, size=12, color=TEXT)
    return h


def _split(slide, x, y, w, b, style):
    left = b.get("left", {})
    right = b.get("right", {})
    lw = w * 0.58
    rw = w - lw - 0.3
    h = 2.1
    if left.get("kicker"):
        add_text(slide, x, y, lw, 0.3, left.get("kicker", ""), size=11, bold=True, color=FAINT)
    t = left.get("title", "")
    if t:
        th = sum(text_height(seg, 24, lw) for seg in t.split("\n")) + 0.1 * (t.count("\n"))
        add_text(slide, x, y + 0.3, lw, th + 0.2, t, size=24, bold=True, line_spacing=1.1)
    body = left.get("body", "")
    if body:
        bh = text_height(body, 14, lw)
        by = y + 0.3 + (th if t else 0) + 0.3
        add_text(slide, x, by, lw, bh + 0.1, body, size=14, color=SUBTLE, line_spacing=1.25)
    rx = x + lw + 0.3
    rk = right.get("kicker", "")
    items = right.get("items", [])
    is_panel = right.get("kind") == "panel"
    if is_panel:
        ihs = [text_height(str(it), 13, rw - 0.4) + 0.16 for it in items]
        ihs = [max(0.5, v) for v in ihs]
        ph = 0.5 + sum(ihs) + 0.3
        add_rect(slide, rx, y, rw, ph, ACCENT_LT)
        if rk:
            add_text(slide, rx + 0.2, y + 0.12, rw - 0.4, 0.3, rk, size=11, bold=True, color=ACCENT_DK)
        ry = y + 0.45
        for it, ih in zip(items, ihs):
            add_text(slide, rx + 0.2, ry, rw - 0.4, ih, str(it), size=13, color=TEXT)
            ry += ih
        h = max(h, ph)
    else:
        ry = y
        if rk:
            add_text(slide, rx, ry, rw, 0.3, rk, size=11, bold=True, color=FAINT)
            ry += 0.35
        for it in items:
            add_hline(slide, rx, rx + rw, ry + 0.3, color=LINE, w=0.5)
            add_text(slide, rx, ry, rw, 0.4, str(it), size=12, color=TEXT)
            ry += 0.5
        h = max(h, ry - y + 0.15)
    return h


def _callouts(slide, x, y, w, b, style):
    items = b.get("items", [])
    h = 0.0
    for it in items:
        ch = 0.85
        body = it.get("body", "")
        if body:
            ch = max(ch, 0.6 + text_height(body, 13, w - 0.5) + 0.15)
        add_rect(slide, x, y + h, w, ch, PANEL)
        add_rect(slide, x, y + h, 0.06, ch, ACCENT)
        add_text(slide, x + 0.22, y + h + 0.12, w - 0.5, 0.3, it.get("title", ""), size=15, bold=True)
        if body:
            add_text(slide, x + 0.22, y + h + 0.46, w - 0.5, ch - 0.5, body, size=13, color=SUBTLE,
                     line_spacing=1.15)
        h += ch + 0.18
    return h


def _cards(slide, x, y, w, b, style):
    items = b.get("items", [])
    cols = int(b.get("columns", 3))
    n = len(items)
    rows = math.ceil(n / cols)
    gap = 0.22
    cw = (w - gap * (cols - 1)) / cols
    ch = 0.95
    h = rows * ch + (rows - 1) * gap
    for i, it in enumerate(items):
        r, c = divmod(i, cols)
        cx = x + c * (cw + gap)
        cy = y + r * (ch + gap)
        add_rect(slide, cx, cy, cw, ch, PANEL)
        add_text(slide, cx + 0.18, cy + 0.12, cw - 0.36, 0.35, it.get("title", ""), size=13, bold=True)
        body = it.get("body", "")
        if body:
            add_text(slide, cx + 0.18, cy + 0.45, cw - 0.36, ch - 0.55, body, size=11, color=SUBTLE,
                     line_spacing=1.1)
    return h


def _rows(slide, x, y, w, b, style):
    items = b.get("items", [])
    h = 0.0
    for it in items:
        rh = 0.56
        tag = it.get("tag", "")
        if tag:
            tw = 0.1 + len(tag) * 0.18
            rh = max(0.7, text_height(it.get("text", ""), 15, w - tw - 0.3) + 0.3)
            add_rect(slide, x, y + h, tw, rh, ACCENT_LT)
            add_text(slide, x, y + h + 0.06, tw, 0.3, tag, size=11, bold=True, color=ACCENT_DK,
                     align=PP_ALIGN.CENTER)
            add_text(slide, x + tw + 0.22, y + h + 0.12, w - tw - 0.3, rh - 0.1, it.get("text", ""),
                     size=15, color=TEXT)
        else:
            add_text(slide, x, y + h + 0.08, w, rh - 0.1, it.get("text", ""), size=13, color=TEXT)
        if h > 0:
            add_hline(slide, x, x + w, y + h, color=LINE, w=0.5)
        h += rh
    return h


def _compare(slide, x, y, w, b, style):
    left = b.get("left", {})
    right = b.get("right", {})
    cw = (w - 0.28) / 2
    h = 2.3
    add_rect(slide, x, y, cw, h, PANEL)
    add_rect(slide, x, y, cw, 0.05, FAINT)
    add_text(slide, x + 0.22, y + 0.2, cw - 0.44, 0.3, left.get("kicker", ""), size=11,
             bold=True, color=FAINT)
    lt = left.get("title", "")
    if lt:
        add_text(slide, x + 0.22, y + 0.5, cw - 0.44, 0.5, lt, size=15, bold=True)
    ly = y + 1.05
    for it in left.get("items", []):
        text = it if isinstance(it, str) else it.get("text", "")
        add_text(slide, x + 0.22, ly, cw - 0.44, 0.4, "×  " + text, size=12, color=SUBTLE)
        ly += 0.42
    rx = x + cw + 0.28
    add_rect(slide, rx, y, cw, h, ACCENT)
    add_rect(slide, rx, y, cw, 0.05, ACCENT_DK)
    add_text(slide, rx + 0.22, y + 0.2, cw - 0.44, 0.3, right.get("kicker", ""), size=11,
             bold=True, color=WHITE)
    rt = right.get("title", "")
    if rt:
        add_text(slide, rx + 0.22, y + 0.5, cw - 0.44, 0.5, rt, size=15, bold=True, color=WHITE)
    ry = y + 1.05
    for it in right.get("items", []):
        text = it if isinstance(it, str) else it.get("text", "")
        add_text(slide, rx + 0.22, ry, cw - 0.44, 0.4, "✓  " + text, size=12, color=WHITE)
        ry += 0.42
    return h


def _pipeline(slide, x, y, w, b, style):
    items = b.get("items", [])
    n = len(items)
    if n == 0:
        return 0
    n = min(n, 5)
    gap = 0.7
    cw = (w - gap * (n - 1)) / n
    body_h = max([text_height(it.get("body", ""), 10, cw - 0.3) for it in items] or [0.2])
    h = 0.7 + body_h + 0.34
    h = max(h, 1.15)
    for i, it in enumerate(items):
        cx = x + i * (cw + gap)
        add_rect(slide, cx, y, cw, h, PANEL)
        add_rect(slide, cx, y, cw, 0.05, ACCENT)
        add_text(slide, cx + 0.15, y + 0.12, cw - 0.3, 0.2, it.get("kicker", ""), size=10,
                 bold=True, color=FAINT)
        add_text(slide, cx + 0.15, y + 0.36, cw - 0.3, 0.3, it.get("title", ""), size=13, bold=True)
        body = it.get("body", "")
        if body:
            add_text(slide, cx + 0.15, y + 0.7, cw - 0.3, body_h + 0.05, body, size=10,
                     color=SUBTLE, line_spacing=1.1)
        metric = it.get("metric", "")
        if metric:
            add_text(slide, cx + 0.15, y + h - 0.26, cw - 0.3, 0.2, metric, size=9, color=ACCENT_DK)
        if i < n - 1:
            add_rect(slide, cx + cw + 0.1, y + h / 2 - 0.11, gap - 0.2, 0.22, FAINT,
                     shape=MSO_SHAPE.RIGHT_ARROW)
    return h


def _hub(slide, x, y, w, b, style):
    center = b.get("center", {})
    items = b.get("items", [])
    cw = 3.4
    ch = 1.1
    cx = x + w / 2 - cw / 2
    cy = y + 0.55
    add_rect(slide, cx, cy, cw, ch, ACCENT)
    add_text(slide, cx + 0.2, cy + 0.12, cw - 0.4, 0.25, center.get("kicker", ""), size=10,
             bold=True, color=WHITE)
    add_text(slide, cx + 0.2, cy + 0.36, cw - 0.4, 0.35, center.get("title", ""), size=15,
             bold=True, color=WHITE)
    body = center.get("body", "")
    if body:
        add_text(slide, cx + 0.2, cy + 0.7, cw - 0.4, 0.3, body, size=10, color=WHITE)
    h = ch + 0.55
    n = len(items)
    if n:
        box_w = (w - 0.3 * (n - 1)) / n if n <= 4 else (w - 0.3 * 3) / 4
        box_w = min(box_w, 2.6)
        gap_x = 0.3
        total = n * box_w + (n - 1) * gap_x
        start = x + (w - total) / 2
        for i, it in enumerate(items):
            bx = start + i * (box_w + gap_x)
            by = cy + ch + 0.35
            add_rect(slide, bx, by, box_w, 1.0, PANEL)
            add_text(slide, bx + 0.15, by + 0.1, box_w - 0.3, 0.3, it.get("title", ""), size=12, bold=True)
            ibody = it.get("body", "")
            if ibody:
                add_text(slide, bx + 0.15, by + 0.38, box_w - 0.3, 0.55, ibody, size=10,
                         color=SUBTLE, line_spacing=1.1)
            add_hline(slide, bx + box_w / 2, cx + cw / 2, by - 0.1, color=FAINT, w=0.75)
        h = ch + 0.35 + 1.0 + 0.2
    return h


def _timeline(slide, x, y, w, b, style):
    items = b.get("items", [])
    n = len(items)
    if n == 0:
        return 0
    n = min(n, 5)
    add_hline(slide, x, x + w, y + 0.5, color=LINE, w=1.2)
    cw = w / n
    h = 2.15
    for i, it in enumerate(items):
        cx = x + i * cw
        add_rect(slide, cx, y + 0.44, 0.12, 0.12, ACCENT if i == 0 else FAINT, shape=MSO_SHAPE.OVAL)
        add_text(slide, cx - 0.2, y + 0.66, cw + 0.4, 0.3, it.get("phase", ""), size=11,
                 bold=True, color=ACCENT_DK)
        add_text(slide, cx - 0.2, y + 0.98, cw + 0.4, 0.35, it.get("title", ""), size=13, bold=True)
        body = it.get("body", "")
        if body:
            add_text(slide, cx - 0.2, y + 1.32, cw + 0.4, 0.5, body, size=10, color=SUBTLE,
                     line_spacing=1.1)
        metric = it.get("metric", "")
        if metric:
            add_text(slide, cx - 0.2, y + h - 0.3, cw + 0.4, 0.2, metric, size=9, color=ACCENT_DK)
    return h


def _points(slide, x, y, w, b, style):
    items = b.get("items", [])
    h = 0.0
    for i, it in enumerate(items):
        ph = 0.72
        add_text(slide, x, y + h + 0.08, 0.5, 0.4, f"{i + 1:02d}", size=14, bold=True, color=ACCENT)
        t = it.get("title", "")
        add_text(slide, x + 0.55, y + h + 0.05, w - 0.6, 0.4, t, size=17, bold=True)
        body = it.get("body", "")
        if body:
            bh = text_height(body, 14, w - 0.6)
            add_text(slide, x + 0.55, y + h + 0.45, w - 0.6, bh + 0.05, body, size=14, color=SUBTLE,
                     line_spacing=1.2)
            ph = 0.5 + bh + 0.22
        add_hline(slide, x, x + w, y + h + ph - 0.05, color=LINE, w=0.5)
        h += ph
    return h


FORMULA_SRC: dict = {}  # 形状名 -> (latex, display)，供 main() 写 .formulas.json


def _formula(slide, x, y, w, b, style):
    """公式占位框，文字为 LaTeX 源码；formula.py 随后把它替换成原生公式。"""
    display = bool(b.get("display", True))
    h = 1.1 if display else 0.6
    name = f"FORMULA_{len(FORMULA_SRC) + 1}"
    tb = add_text(slide, x, y, w, h, b["latex"], size=18, color=TEXT,
                  align=PP_ALIGN.CENTER if display else PP_ALIGN.LEFT,
                  anchor=MSO_ANCHOR.MIDDLE)
    tb.name = name
    FORMULA_SRC[name] = (b["latex"], display)
    caption = b.get("caption", "")
    if caption:
        add_text(slide, x, y + h + 0.05, w, 0.25, caption, size=10, color=FAINT,
                 align=PP_ALIGN.CENTER if display else PP_ALIGN.LEFT)
        h += 0.3
    return h


BASE_DIR = Path.cwd()  # 图片 src 相对 deck.json 所在目录，main() 里覆盖


def _figure(slide, x, y, w, b, style):
    """资料原图 + 可选侧栏文字。图片等比缩放放入框内，不拉伸。"""
    from PIL import Image
    src = Path(b.get("src", ""))
    if not src.is_absolute():
        src = BASE_DIR / src
    if not src.exists():
        raise RuntimeError(f"figure 图片不存在: {src}")
    with Image.open(src) as im:
        pw, ph = im.size
    aspect = pw / ph
    caption = b.get("caption", "")
    cap_h = 0.3 if caption else 0.0
    box_h = float(b.get("height", CONTENT_BOTTOM - y - cap_h))
    side = b.get("side")
    gap = 0.35
    iw = w * float(b.get("image_ratio", 0.62)) if side else w
    tw = w - iw - gap
    img_left = b.get("image_side", "right") == "left"
    ix0 = x if (not side or img_left) else x + tw + gap
    tx = x + iw + gap if img_left else x
    dw = iw
    dh = dw / aspect
    if dh > box_h:
        dh = box_h
        dw = dh * aspect
    ix = ix0 + (iw - dw) / 2 if not side else ix0
    if side and img_left is False:
        ix = ix0 + iw - dw
    slide.shapes.add_picture(str(src), Inches(ix), Inches(y), Inches(dw), Inches(dh))
    add_rect(slide, ix, y, dw, dh, None, line=LINE, line_w=0.5)
    if caption:
        add_text(slide, ix, y + dh + 0.07, dw, 0.25, caption, size=10, color=FAINT)
    if side:
        ty = y
        if side.get("kicker"):
            add_text(slide, tx, ty, tw, 0.3, side["kicker"], size=11, bold=True, color=FAINT)
            ty += 0.35
        if side.get("title"):
            th = text_height(side["title"], 18, tw)
            add_text(slide, tx, ty, tw, th + 0.1, side["title"], size=18, bold=True, line_spacing=1.15)
            ty += th + 0.25
        for it in side.get("items", []):
            ih = text_height(str(it), 12, tw)
            add_hline(slide, tx, tx + tw, ty - 0.06, color=LINE, w=0.5)
            add_text(slide, tx, ty, tw, ih + 0.15, str(it), size=12, color=TEXT, line_spacing=1.25)
            ty += ih + 0.22
    return dh + cap_h


BLOCK_RENDERERS = {
    "figure": _figure,
    "lead": _lead,
    "insight": _insight,
    "stats": _stats,
    "columns": _columns,
    "steps": _steps,
    "chips": _chips,
    "layers": _layers,
    "table": _table,
    "split": _split,
    "callouts": _callouts,
    "cards": _cards,
    "rows": _rows,
    "compare": _compare,
    "pipeline": _pipeline,
    "hub": _hub,
    "timeline": _timeline,
    "points": _points,
    "formula": _formula,
}


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", required=True)
    args = ap.parse_args()

    deck_path = Path(args.input)
    try:
        deck = json.loads(deck_path.read_text(encoding="utf-8"))
    except Exception as e:
        sys.stderr.write(f"deck.json 读取失败: {e}\n")
        return 2

    errs = validate_deck(deck, extra_kinds=set(BLOCK_RENDERERS))
    if errs:
        for e in errs:
            sys.stderr.write(f"✗ {e}\n")
        sys.stderr.write("请先跑 deck_check.py 修好 deck.json 再渲染，禁止跳过。\n")
        return 1

    global BASE_DIR
    BASE_DIR = deck_path.resolve().parent
    source = deck.get("source", "")
    style = {"text": TEXT, "sub": SUBTLE, "accent": ACCENT, "panel": PANEL, "source": source}

    prs, blank = new_presentation()
    formulas = []
    FORMULA_SRC.clear()
    for idx, s in enumerate(deck["slides"], start=1):
        slide = prs.slides.add_slide(blank)
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = WHITE
        s_type = s.get("type")
        if s_type == "cover":
            render_cover(slide, s)
        elif s_type == "catalog":
            render_catalog(slide, s)
        elif s_type == "section":
            render_section(slide, s, s.get("num", ""), s.get("title", ""),
                           s.get("subtitle", ""))
        elif s_type == "close":
            render_cover(slide, s)
        else:
            render_content(slide, s, style)
            for sh in slide.shapes:
                if sh.name in FORMULA_SRC:
                    latex, display = FORMULA_SRC[sh.name]
                    formulas.append({"slide": idx, "shape": sh.name, "latex": latex,
                                     "display": display, "left": int(sh.left),
                                     "top": int(sh.top), "width": int(sh.width),
                                     "height": int(sh.height)})

    out = Path(args.output).expanduser()
    if not out.is_absolute():
        out = Path.cwd() / out
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    if formulas:
        sidecar = out.with_name(out.name + ".formulas.json")
        sidecar.write_text(json.dumps(formulas, ensure_ascii=False, indent=2), encoding="utf-8")
    print(str(out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
