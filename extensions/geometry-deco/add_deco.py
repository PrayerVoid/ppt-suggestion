#!/usr/bin/env python3
"""对成品 pptx 后处理，加整套风格一致的几何装饰，并自动避开页面上的文字与图。

Usage:
  python add_deco.py in.pptx --deck deck.json --out out.pptx [--motif NAME] [--seed N] [--accent HEX]
  python add_deco.py in.pptx --deck deck.json --out out.pptx --mode background [--profile formal|balanced|expressive]

一份 deck 只用一种图案族（由 --motif 固定，或由 --seed 可复现地选择）：
  封面/结束页用大构图，章节页用同一构图的上下翻转版；目录页与内容页用完全相同的小角标。
装饰统一命名 DECO_<序号>，放在每页最底层。生成前读取每页所有非装饰形状（含组合内子形状）的位置，
整套选一个对所有同类页都不冲突的角落与缩放；没有这样的方案时，冲突的页省略装饰并打印页号。
重复运行会先删除已有 DECO_ 形状。
"""

from __future__ import annotations

import argparse
import json
import math
import random
import sys
from pathlib import Path

from lxml import etree

try:
    from power_pptx import Presentation
    from power_pptx.dml.color import RGBColor
    from power_pptx.enum.shapes import MSO_SHAPE
    from power_pptx.oxml.ns import qn
    from power_pptx.util import Inches, Pt
    from decor_templates import TEMPLATES
except ImportError:
    sys.stderr.write("缺少 power-pptx 或几何模板依赖，请先运行: pip install -r requirements.txt\n")
    raise SystemExit(3)

TEMPLATE_FUNCS = {tid: fn for tid, _, _, fn, _, _, _ in TEMPLATES}
TEMPLATE_IDS = tuple(TEMPLATE_FUNCS)

EMU = 914400.0
PAD = 0.12                      # 装饰与文字之间至少留的空隙（英寸）
BIG_SIDE, SMALL_SIDE = 3.6, 1.2  # 构图所在方形框的边长（英寸）
BIG_SCALES = (1.0, 0.8, 0.65, 0.5)
SMALL_SCALES = (1.0, 0.8, 0.65, 0.5)
MIN_LINE = 0.02                 # 细线最小厚度（英寸），再细在投影时会消失

# 图案族：(形状, 色, x, y, w, h)，坐标是 1×1 方框内的比例，原点在框左上角，框贴页面右侧
FAMILIES = {
    "circles": [
        ("oval", "tint", 0.38, 0.00, 0.62, 0.62),
        ("ring", "accent", 0.00, 0.40, 0.42, 0.42),
        ("oval", "accent", 0.86, 0.80, 0.07, 0.07),
    ],
    "blocks": [
        ("rect", "tint", 0.55, 0.00, 0.45, 0.45),
        ("rect", "accent", 0.80, 0.45, 0.20, 0.05),
        ("rect", "accent", 0.55, 0.52, 0.10, 0.10),
    ],
    "lines": [
        ("rect", "tint", 0.40, 0.00, 0.60, 0.40),
        ("rect", "accent", 0.00, 0.55, 1.00, 0.012),
        ("rect", "accent", 0.35, 0.62, 0.65, 0.012),
        ("rect", "accent", 0.65, 0.69, 0.35, 0.012),
    ],
}

# 纹样模板库作为可选扩展；默认使用 circles。
MOTIF_NAMES = tuple(FAMILIES) + TEMPLATE_IDS

# 背景模式的安全上限；角落模式使用内容避让检查。
PROFILES = {
    "formal": {
        "motifs": frozenset(("circles", "lines")),
        "max_shapes": 8,
        "max_coverage": 0.08,
        "max_opacity": 0.12,
        "opacity": 0.10,
    },
    "balanced": {
        "motifs": frozenset(MOTIF_NAMES),
        "max_shapes": 64,
        "max_coverage": 0.16,
        "max_opacity": 0.20,
        "opacity": 0.14,
    },
    "expressive": {
        "motifs": frozenset(MOTIF_NAMES),
        "max_shapes": 220,
        "max_coverage": 0.24,
        "max_opacity": 0.28,
        "opacity": 0.18,
    },
}


def _frame(side, scale, anchor, pw, ph):
    s = side * scale
    left = pw - s
    top = 0.0 if anchor == "TR" else ph - s
    return left, top, s


# 各类页的角落偏好：TR 右上，BR 右下（右下构图用上下翻转）
ROLES = {
    "big": {"types": ("cover", "close"), "anchors": ("TR", "BR"), "side": BIG_SIDE, "scales": BIG_SCALES},
    "section": {"types": ("section",), "anchors": ("BR", "TR"), "side": BIG_SIDE, "scales": BIG_SCALES},
    "small": {"types": ("catalog", "content"), "anchors": ("BR", "TR"), "side": SMALL_SIDE, "scales": SMALL_SCALES},
}


def place(family, side, scale, anchor, pw, ph):
    """返回图元 (形状, 色, l, t, w, h[, 旋转])，英寸；贴页面右侧并留在页内。"""
    if family in TEMPLATE_FUNCS:
        left, top, s = _frame(side, scale, anchor, pw, ph)
        return [("template", family, left, top, s, s)]
    s = side * scale
    out = []
    for kind, color, x, y, w, h in FAMILIES[family]:
        ww, hh = w * s, h * s
        if kind == "rect" and min(ww, hh) < MIN_LINE:
            ww, hh = (ww, MIN_LINE) if hh < ww else (MIN_LINE, hh)
        left = pw - s + x * s
        top = y * s if anchor == "TR" else ph - s + (1 - y - h) * s
        out.append((kind, color, min(left, pw - ww), min(max(top, 0.0), ph - hh), ww, hh))
    return out


def _group_boxes(grp):
    """组合内子形状换算到页面坐标；换算失败退回组合外框。"""
    boxes = []
    try:
        x = grp._element.grpSpPr.xfrm
        sx = x.ext.cx / x.chExt.cx if x.chExt.cx else 1.0
        sy = x.ext.cy / x.chExt.cy if x.chExt.cy else 1.0
        for ch in grp.shapes:
            l = x.off.x + (ch.left - x.chOff.x) * sx
            t = x.off.y + (ch.top - x.chOff.y) * sy
            boxes.append((l / EMU, t / EMU, (l + ch.width * sx) / EMU, (t + ch.height * sy) / EMU))
    except Exception:
        boxes = []
    return boxes or [(grp.left / EMU, grp.top / EMU, (grp.left + grp.width) / EMU, (grp.top + grp.height) / EMU)]


def _has_explicit_fill(sh):
    try:
        sp_pr = sh._element.spPr
        return any(sp_pr.find(qn(tag)) is not None for tag in (
            "a:solidFill", "a:gradFill", "a:pattFill", "a:blipFill"))
    except Exception:
        return True


def _text_occupied_box(sh):
    """返回无填充文本的实际占用框；None 表示按完整形状框保护，() 表示空文本。"""
    try:
        if not sh.has_text_frame or _has_explicit_fill(sh):
            return None
        paragraphs = list(sh.text_frame.paragraphs)
    except Exception:
        return None
    if not paragraphs:
        return ()

    shape_width = sh.width / EMU
    shape_height = sh.height / EMU
    metrics = []
    for paragraph in paragraphs:
        runs = list(paragraph.runs)
        text = "".join(run.text or "" for run in runs)
        if not text:
            text = getattr(paragraph, "text", "") or ""
        if not text.strip():
            continue
        for line in text.splitlines() or [text]:
            sizes = []
            for run in runs:
                try:
                    if run.font.size is not None:
                        sizes.append(run.font.size.pt)
                except Exception:
                    pass
            size = max(sizes, default=14.0)
            width = sum((1.0 if ord(ch) > 0x2E80 else 0.55) * size / 72.0
                        for ch in line)
            metrics.append((width, size))
    if not metrics:
        return ()

    used_width = min(shape_width, max(0.25, max(width for width, _ in metrics)))
    line_height = max(size for _, size in metrics) * 1.22 / 72.0
    line_count = sum(max(1, math.ceil(width / max(shape_width, 0.12)))
                     for width, _ in metrics)
    used_height = min(shape_height, max(0.25, line_count * line_height))
    left = sh.left / EMU
    top = sh.top / EMU
    alignments = [str(getattr(p, "alignment", "")).upper() for p in paragraphs]
    if any("RIGHT" in alignment for alignment in alignments):
        left += shape_width - used_width
    elif any("CENTER" in alignment for alignment in alignments):
        left += (shape_width - used_width) / 2
    return left, top, left + used_width, top + used_height


def obstacles(slide, pw, ph, include_large=False, occupied_text=False):
    out = []
    for sh in slide.shapes:
        if (sh.name or "").startswith("DECO_") or sh.width is None or sh.left is None:
            continue
        text_box = _text_occupied_box(sh) if occupied_text else None
        if text_box == ():
            continue
        if text_box is not None:
            boxes = [text_box]
        else:
            boxes = _group_boxes(sh) if sh._element.tag == qn("p:grpSp") else [
                (sh.left / EMU, sh.top / EMU,
                 (sh.left + sh.width) / EMU, (sh.top + sh.height) / EMU)]
        if include_large:
            out += boxes
        else:
            out += [b for b in boxes if (b[2] - b[0]) * (b[3] - b[1]) <= 0.6 * pw * ph]
    return out


def collides(parts, obs):
    for part in parts:
        _, _, l, t, w, h = part[:6]
        for ol, ot, orr, ob in obs:
            if l < orr + PAD and l + w > ol - PAD and t < ob + PAD and t + h > ot - PAD:
                return True
    return False


def _shape_bbox(sh):
    """读取形状的实际外接框，旋转形状按旋转后的矩形保守计算。"""
    left = sh.left / EMU
    top = sh.top / EMU
    width = sh.width / EMU
    height = sh.height / EMU
    angle = float(getattr(sh, "rotation", 0) or 0) % 180
    if angle:
        rad = math.radians(angle)
        bw = abs(width * math.cos(rad)) + abs(height * math.sin(rad))
        bh = abs(width * math.sin(rad)) + abs(height * math.cos(rad))
        cx, cy = left + width / 2, top + height / 2
        return cx - bw / 2, cy - bh / 2, cx + bw / 2, cy + bh / 2
    return left, top, left + width, top + height


def _boxes_collide(boxes, obstacles_, pad=PAD):
    for left, top, right, bottom in boxes:
        for ol, ot, oright, obottom in obstacles_:
            if (left < oright + pad and right > ol - pad and
                    top < obottom + pad and bottom > ot - pad):
                return True
    return False


def _union_area(boxes):
    """求矩形并集面积，用于背景覆盖率的保守门槛。"""
    if not boxes:
        return 0.0
    xs = sorted({x for box in boxes for x in (box[0], box[2])})
    area = 0.0
    for x0, x1 in zip(xs, xs[1:]):
        if x1 <= x0:
            continue
        intervals = sorted((box[1], box[3]) for box in boxes
                           if box[0] < x1 and box[2] > x0)
        covered = 0.0
        end = None
        for y0, y1 in intervals:
            if end is None:
                covered, end = max(0.0, y1 - y0), y1
            elif y0 > end:
                area += (x1 - x0) * covered
                covered, end = max(0.0, y1 - y0), y1
            else:
                if y1 > end:
                    covered += y1 - end
                    end = y1
        area += (x1 - x0) * covered
    return area


def _part_boxes(parts):
    return [(part[2], part[3], part[2] + part[4], part[3] + part[5])
            for part in parts]


def _content_protection(ds, pw, ph):
    """按 deck 页面语义保留标题、正文和页脚的不可侵入区域。"""
    slide_type = ds.get("type")
    boxes = [
        (0.55, 0.15, max(0.0, pw - 1.10), 1.20),
        (0.55, max(0.0, ph - 0.95), max(0.0, pw - 1.10), 0.80),
    ]
    if slide_type in ("cover", "close"):
        boxes.append((0.55, 1.35, min(8.75, pw - 1.10), 4.85))
    elif slide_type == "section":
        boxes.append((0.55, 1.25, min(8.75, pw - 1.10), 3.75))
    elif slide_type == "catalog":
        boxes.append((0.55, 1.50, max(0.0, pw - 1.10), 5.25))
    elif slide_type == "content":
        # 右侧窄安全带仍需经过实际对象检查，正文预留区只保护主体排版。
        boxes.append((0.55, 1.45, max(0.0, pw - 2.10), 5.45))
    return [(l, t, l + w, t + h) for l, t, w, h in boxes if w > 0 and h > 0]


def _background_regions(ds, pw, ph):
    """返回由页面边缘向内的候选底纹区域，按优先级排列。"""
    slide_type = ds.get("type")
    if slide_type in ("cover", "close", "section"):
        side_w = min(3.60, max(1.10, pw * 0.28))
        # 标题保护区结束后再开始，页脚上方保留一段中部安全带。
        side_top = 1.55
    else:
        side_w = min(1.15, max(0.85, pw * 0.10))
        side_top = 1.50
    side_bottom = max(side_top + 0.5, ph - 1.10)
    return [
        ("右侧", (pw - side_w - 0.15, side_top, side_w, side_bottom - side_top)),
        ("左侧", (0.15, side_top, side_w, side_bottom - side_top)),
        ("顶部", (0.25, 0.15, max(0.5, pw - 0.50), 0.72)),
        ("底部", (0.25, max(0.15, ph - 0.90), max(0.5, pw - 0.50), 0.65)),
    ]


def _background_parts(family, region):
    """在一个候选矩形内放置单个低密度母题实例。"""
    x, y, w, h = region
    side = min(w, h)
    if family in TEMPLATE_FUNCS:
        return [("template", family, x + (w - side) / 2, y + (h - side) / 2, side, side)]
    local = place(family, side, 1.0, "TR", side, side)
    return [(kind, color, left + x, top + y + (h - side) / 2, ww, hh, *part[6:])
            for part in local
            for kind, color, left, top, ww, hh in [part[:6]]]


def _set_shape_opacity(sh, opacity):
    """同时限制填充和线条透明度，避免背景母题成为新的视觉焦点。"""
    sp_pr = sh._element.spPr
    fills = []
    fill = sp_pr.find(qn("a:solidFill"))
    if fill is not None:
        fills.append(fill)
    line = sp_pr.find(qn("a:ln"))
    if line is not None:
        line_fill = line.find(qn("a:solidFill"))
        if line_fill is not None:
            fills.append(line_fill)
    for solid in fills:
        color = next((solid.find(qn(tag)) for tag in ("a:srgbClr", "a:schemeClr")
                      if solid.find(qn(tag)) is not None), None)
        if color is None:
            continue
        alpha = color.find(qn("a:alpha"))
        if alpha is None:
            alpha = etree.SubElement(color, qn("a:alpha"))
        alpha.set("val", str(round(opacity * 100000)))


def _shape_opacity(sh):
    sp_pr = sh._element.spPr
    alphas = []
    for solid in (sp_pr.find(qn("a:solidFill")),):
        if solid is None:
            continue
        for tag in ("a:srgbClr", "a:schemeClr"):
            color = solid.find(qn(tag))
            if color is not None:
                alpha = color.find(qn("a:alpha"))
                alphas.append(float(alpha.get("val")) / 100000 if alpha is not None else 1.0)
    line = sp_pr.find(qn("a:ln"))
    if line is not None:
        solid = line.find(qn("a:solidFill"))
        if solid is not None:
            for tag in ("a:srgbClr", "a:schemeClr"):
                color = solid.find(qn(tag))
                if color is not None:
                    alpha = color.find(qn("a:alpha"))
                    alphas.append(float(alpha.get("val")) / 100000 if alpha is not None else 1.0)
    return max(alphas, default=0.0)


def _remove_shapes(shapes):
    for sh in shapes:
        parent = sh._element.getparent()
        if parent is not None:
            parent.remove(sh._element)


def choose(roles, pages_by_role, family, pw, ph):
    """为一组角色选共同的缩放与各自的角落；返回 {角色: (锚点, 缩放, 冲突页索引)}。"""
    scales = sorted({s for r in roles for s in ROLES[r]["scales"]}, reverse=True)
    last = {}
    for scale in scales:
        picked = {}
        for r in roles:
            cfg = ROLES[r]
            cands = []
            for anchor in cfg["anchors"]:
                parts = place(family, cfg["side"], scale, anchor, pw, ph)
                bad = [i for i, obs in pages_by_role[r] if collides(parts, obs)]
                cands.append((len(bad), anchor, bad))
            n, anchor, bad = min(cands, key=lambda c: c[0])
            picked[r] = (anchor, scale, bad)
        last = picked
        if all(not v[2] for v in picked.values()):
            return picked
    return last


def _rgb_hex(color):
    return "".join(f"{int(v):02X}" for v in color)


def _template_palette(accent, tint):
    accent_hex = _rgb_hex(accent)
    tint_hex = _rgb_hex(tint)
    return {
        "main": accent_hex,
        "sub": accent_hex,
        "accent": accent_hex,
        "ink": "1A1A1A",
        "gray": "666660",
        "card": tint_hex,
        "hair": "D9D9D4",
    }


def _fit_template(added, pw, ph):
    """把模板形状整体移回页面内，处理旋转后外接框变大的情况。"""
    if not added:
        return
    min_left = min(sh.left / EMU for sh in added)
    min_top = min(sh.top / EMU for sh in added)
    max_right = max((sh.left + sh.width) / EMU for sh in added)
    max_bottom = max((sh.top + sh.height) / EMU for sh in added)
    dx = 0.01 - min_left if min_left < 0.01 else 0.0
    dy = 0.01 - min_top if min_top < 0.01 else 0.0
    if max_right + dx > pw - 0.01:
        dx = pw - 0.01 - max_right
    if max_bottom + dy > ph - 0.01:
        dy = ph - 0.01 - max_bottom
    if dx or dy:
        for sh in added:
            sh.left = int(sh.left + dx * EMU)
            sh.top = int(sh.top + dy * EMU)


def _draw_template(slide, part, accent, tint, start, pw, ph, opacity=None):
    _, template_id, left, top, width, height = part
    before = len(slide.shapes)
    fn = TEMPLATE_FUNCS[template_id]
    fn(slide, left * 96, top * 96, width * 96, height * 96,
       pal=_template_palette(accent, tint))
    added = list(slide.shapes)[before:]
    _fit_template(added, pw, ph)
    tree = slide.shapes._spTree
    for k, sh in enumerate(added):
        if opacity is not None:
            _set_shape_opacity(sh, opacity)
        sh.name = f"DECO_{start + k}"
        el = sh._element
        tree.remove(el)
        tree.insert(2 + k, el)
    return start + len(added)


def draw(slide, parts, accent, tint, start, pw, ph, opacity=None):
    if parts and parts[0][0] == "template":
        return _draw_template(slide, parts[0], accent, tint, start, pw, ph, opacity=opacity)
    tree = slide.shapes._spTree
    shapes = {
        "rect": MSO_SHAPE.RECTANGLE,
        "oval": MSO_SHAPE.OVAL,
        "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
        "diamond": MSO_SHAPE.DIAMOND,
    }
    for k, part in enumerate(parts):
        kind, color, l, t, w, h = part[:6]
        rotation = part[6] if len(part) > 6 else None
        sh = slide.shapes.add_shape(shapes.get(kind, MSO_SHAPE.OVAL),
                                    Inches(l), Inches(t), Inches(w), Inches(h))
        if kind == "ring":
            sh.fill.background()
            sh.line.color.rgb = accent if color == "accent" else tint
            sh.line.width = Pt(1)
        else:
            sh.fill.solid()
            sh.fill.fore_color.rgb = tint if color == "tint" else accent
            sh.line.fill.background()
        if rotation is not None:
            sh.rotation = rotation
        if opacity is not None:
            _set_shape_opacity(sh, opacity)
        sh.shadow.clear()
        sh.name = f"DECO_{start + k}"
        el = sh._element
        tree.remove(el)
        tree.insert(2 + k, el)
    return start + len(parts)


def _background_attempt(slide, parts, accent, tint, start, pw, ph,
                        protected, object_boxes, limits):
    planned = _part_boxes(parts)
    if _boxes_collide(planned, protected + object_boxes):
        return start, "候选纹样会进入内容保护区或对象保护区"

    before_names = {sh.name for sh in slide.shapes}
    next_start = draw(slide, parts, accent, tint, start, pw, ph,
                      opacity=limits["opacity"])
    added = [sh for sh in slide.shapes
             if sh.name.startswith("DECO_") and sh.name not in before_names]
    actual = [_shape_bbox(sh) for sh in added]
    reasons = []
    if any(l < 0.01 or t < 0.01 or r > pw - 0.01 or b > ph - 0.01
           for l, t, r, b in actual):
        reasons.append("实际形状越出页面边界")
    if _boxes_collide(actual, protected):
        reasons.append("生成后实际形状碰撞内容保护区")
    if _boxes_collide(actual, object_boxes):
        reasons.append("生成后实际形状碰撞对象保护区")
    if len(added) > limits["max_shapes"]:
        reasons.append(f"复杂度 {len(added)} 个形状超过上限 {limits['max_shapes']}")
    coverage = _union_area(actual) / (pw * ph) if pw and ph else 1.0
    if coverage > limits["max_coverage"]:
        reasons.append(f"覆盖率 {coverage:.1%} 超过上限 {limits['max_coverage']:.1%}")
    opacity = max((_shape_opacity(sh) for sh in added), default=0.0)
    if opacity > limits["max_opacity"] + 1e-6:
        reasons.append(f"透明度 {opacity:.1%} 超过上限 {limits['max_opacity']:.1%}")
    if reasons:
        _remove_shapes(added)
        return start, "；".join(reasons)
    return next_start, None


def _add_background(slide, ds, family, accent, tint, start, pw, ph, profile):
    limits = PROFILES[profile]
    if family not in limits["motifs"]:
        return start, f"profile={profile} 只允许低密度 circles/lines，拒绝母题 {family}"
    protected = _content_protection(ds, pw, ph)
    object_boxes = obstacles(slide, pw, ph, include_large=True, occupied_text=True)
    failures = []
    for region_name, region in _background_regions(ds, pw, ph):
        parts = _background_parts(family, region)
        next_start, reason = _background_attempt(
            slide, parts, accent, tint, start, pw, ph, protected, object_boxes, limits)
        if reason is None:
            print(f"background: 第{ds['_index'] + 1}页使用{region_name}安全区域，"
                  f"母题 {family}，profile={profile}")
            return next_start, None
        failures.append(f"{region_name}：{reason}")
    return start, "无安全位置（" + "；".join(failures) + "）"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--deck", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--seed", type=int)
    ap.add_argument("--motif", choices=MOTIF_NAMES,
                    help="固定几何母题；不传时由 seed 选择，未传 seed 使用 circles")
    ap.add_argument("--mode", choices=("corner", "background", "auto"), default="auto",
                    help="corner 使用角落装饰；background 必须显式启用；auto 默认走 corner")
    ap.add_argument("--profile", choices=tuple(PROFILES), default="formal",
                    help="背景安全门槛：formal、balanced 或 expressive")
    ap.add_argument("--accent", default="4F6082")
    args = ap.parse_args()

    deck = json.loads(Path(args.deck).read_text(encoding="utf-8"))
    prs = Presentation(args.pptx)
    if len(deck["slides"]) != len(prs.slides):
        sys.stderr.write("页数与 deck.json 不一致，请确认 deck 对应这份 pptx\n")
        return 1

    hexv = args.accent.lstrip("#")
    rgb = tuple(int(hexv[i:i + 2], 16) for i in (0, 2, 4))
    accent = RGBColor(*rgb)
    tint = RGBColor(*(round(255 - (255 - c) * 0.12) for c in rgb))
    mode = "corner" if args.mode == "auto" else args.mode
    if args.motif:
        family = args.motif
    elif args.seed is not None:
        choices = (MOTIF_NAMES if mode == "corner"
                   else tuple(PROFILES[args.profile]["motifs"]))
        family = random.Random(args.seed).choice(choices)
    else:
        family = "circles"
    pw, ph = prs.slide_width / EMU, prs.slide_height / EMU

    for slide in prs.slides:
        for sh in [s for s in slide.shapes if (s.name or "").startswith("DECO_")]:
            sh._element.getparent().remove(sh._element)

    n = 1
    if mode == "corner":
        role_of = {t: r for r, c in ROLES.items() for t in c["types"]}
        pages = {r: [] for r in ROLES}
        for i, (slide, ds) in enumerate(zip(prs.slides, deck["slides"])):
            r = role_of.get(ds.get("type"))
            if r:
                pages[r].append((i, obstacles(slide, pw, ph)))

        plan = {**choose(["big", "section"], pages, family, pw, ph),
                **choose(["small"], pages, family, pw, ph)}
        for r, (anchor, scale, bad) in plan.items():
            print(f"{r}: 图案族 {family}，角落 {anchor}，缩放 {scale}" +
                  (f"，省略第 {[i + 1 for i in bad]} 页" if bad else ""))
            parts = place(family, ROLES[r]["side"], scale, anchor, pw, ph)
            for i, _ in pages[r]:
                if i not in bad:
                    n = draw(prs.slides[i], parts, accent, tint, n, pw, ph)
    else:
        for i, (slide, ds) in enumerate(zip(prs.slides, deck["slides"])):
            slide_data = dict(ds)
            slide_data["_index"] = i
            n, reason = _add_background(slide, slide_data, family, accent, tint,
                                        n, pw, ph, args.profile)
            if reason:
                print(f"background: 第{i + 1}页跳过（{reason}）")

    out = Path(args.out).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print(str(out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
