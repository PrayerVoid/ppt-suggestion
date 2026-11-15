#!/usr/bin/env python3
"""对成品 pptx 后处理，加整套风格一致的几何装饰，并自动避开页面上的文字与图。

Usage:
  python add_deco.py in.pptx --deck deck.json --out out.pptx [--seed N] [--accent HEX]

一份 deck 只用一种图案族（由 --seed 决定，不传则固定第一种）：
  封面/结束页用大构图，章节页用同一构图的上下翻转版；目录页与内容页用完全相同的小角标。
装饰统一命名 DECO_<序号>，放在每页最底层。生成前读取每页所有非装饰形状（含组合内子形状）的位置，
整套选一个对所有同类页都不冲突的角落与缩放；没有这样的方案时，冲突的页省略装饰并打印页号。
重复运行会先删除已有 DECO_ 形状。
"""

from __future__ import annotations

import argparse
import json
import random
import sys
from pathlib import Path

try:
    from power_pptx import Presentation
    from power_pptx.dml.color import RGBColor
    from power_pptx.enum.shapes import MSO_SHAPE
    from power_pptx.oxml.ns import qn
    from power_pptx.util import Inches, Pt
except ImportError:
    sys.stderr.write("缺少 power-pptx，请先运行: pip install -r requirements.txt\n")
    raise SystemExit(3)

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
# 各类页的角落偏好：TR 右上，BR 右下（右下构图用上下翻转）
ROLES = {
    "big": {"types": ("cover", "close"), "anchors": ("TR", "BR"), "side": BIG_SIDE, "scales": BIG_SCALES},
    "section": {"types": ("section",), "anchors": ("BR", "TR"), "side": BIG_SIDE, "scales": BIG_SCALES},
    "small": {"types": ("catalog", "content"), "anchors": ("BR", "TR"), "side": SMALL_SIDE, "scales": SMALL_SCALES},
}


def place(family, side, scale, anchor, pw, ph):
    """返回构图里每个图元在页面上的 (形状, 色, l, t, w, h)，英寸；贴页面右侧并留在页内。"""
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


def obstacles(slide, pw, ph):
    out = []
    for sh in slide.shapes:
        if (sh.name or "").startswith("DECO_") or sh.width is None or sh.left is None:
            continue
        boxes = _group_boxes(sh) if sh._element.tag == qn("p:grpSp") else [
            (sh.left / EMU, sh.top / EMU, (sh.left + sh.width) / EMU, (sh.top + sh.height) / EMU)]
        out += [b for b in boxes if (b[2] - b[0]) * (b[3] - b[1]) <= 0.6 * pw * ph]
    return out


def collides(parts, obs):
    for _, _, l, t, w, h in parts:
        for ol, ot, orr, ob in obs:
            if l < orr + PAD and l + w > ol - PAD and t < ob + PAD and t + h > ot - PAD:
                return True
    return False


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


def draw(slide, parts, accent, tint, start):
    tree = slide.shapes._spTree
    for k, (kind, color, l, t, w, h) in enumerate(parts):
        sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE if kind == "rect" else MSO_SHAPE.OVAL,
                                    Inches(l), Inches(t), Inches(w), Inches(h))
        if kind == "ring":
            sh.fill.background()
            sh.line.color.rgb = accent
            sh.line.width = Pt(1)
        else:
            sh.fill.solid()
            sh.fill.fore_color.rgb = tint if color == "tint" else accent
            sh.line.fill.background()
        sh.shadow.clear()
        sh.name = f"DECO_{start + k}"
        el = sh._element
        tree.remove(el)
        tree.insert(2 + k, el)
    return start + len(parts)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--deck", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--seed", type=int)
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
    family = list(FAMILIES)[random.Random(args.seed).randrange(len(FAMILIES))] if args.seed is not None else "circles"
    pw, ph = prs.slide_width / EMU, prs.slide_height / EMU

    for slide in prs.slides:
        for sh in [s for s in slide.shapes if (s.name or "").startswith("DECO_")]:
            sh._element.getparent().remove(sh._element)

    role_of = {t: r for r, c in ROLES.items() for t in c["types"]}
    pages = {r: [] for r in ROLES}
    for i, (slide, ds) in enumerate(zip(prs.slides, deck["slides"])):
        r = role_of.get(ds.get("type"))
        if r:
            pages[r].append((i, obstacles(slide, pw, ph)))

    plan = {**choose(["big", "section"], pages, family, pw, ph), **choose(["small"], pages, family, pw, ph)}
    n = 1
    for r, (anchor, scale, bad) in plan.items():
        print(f"{r}: 图案族 {family}，角落 {anchor}，缩放 {scale}" + (f"，省略第 {[i + 1 for i in bad]} 页" if bad else ""))
        parts = place(family, ROLES[r]["side"], scale, anchor, pw, ph)
        for i, _ in pages[r]:
            if i not in bad:
                n = draw(prs.slides[i], parts, accent, tint, n)

    out = Path(args.out).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print(str(out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
