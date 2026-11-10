#!/usr/bin/env python3
"""Mechanical validation for a rendered PPTX (ppt-suggestion phase 4).

Checks per slide: shape overlap, out-of-bounds, estimated text overflow,
empty shapes, image distortion / low resolution, layout repetition, and
(optionally, with --deck) slide count and per-slide block signature.

Usage:
  python validate_pptx.py <out.pptx> [--deck deck.json] [--strict]

Exit codes:
  0  no errors (warnings allowed unless --strict)
  1  at least one error
  2  file open / parse failure
"""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path

try:
    from pptx import Presentation
except ImportError:
    sys.stderr.write("python-pptx 未安装，请用 venv python 运行本脚本\n")
    raise SystemExit(2)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from text_measure import estimate_height as measure_height  # noqa: E402
from text_measure import estimate_width  # noqa: E402

EMU_IN = 914400.0
PX_PER_IN = 96.0


def _inches(v) -> float:
    return (v or 0) / EMU_IN


def bbox(sh):
    return (_inches(getattr(sh, "left", None)), _inches(getattr(sh, "top", None)),
            _inches(getattr(sh, "width", None)), _inches(getattr(sh, "height", None)))


def has_visible_fill(sh) -> bool:
    """True when the shape carries an explicit solidFill (panel/rect)."""
    try:
        sp = sh._element.spPr
        if sp is None:
            return False
        from lxml import etree
        ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
        if sp.find(f"{ns}noFill") is not None:
            return False
        return sp.find(f"{ns}solidFill") is not None or sp.find(f"{ns}gradFill") is not None
    except Exception:
        return False


def has_text(sh) -> bool:
    try:
        if not sh.has_text_frame:
            return False
        return bool(sh.text_frame.text.strip())
    except Exception:
        return False


def estimate_text_height(tf, width_in: float) -> float:
    total = 0.0
    for p in tf.paragraphs:
        explicit = [r.font.size.pt for r in p.runs if r.font.size is not None]
        size = max(explicit) if explicit else 14.0
        text = "".join(r.text for r in p.runs)
        if not text:
            total += size * 1.22 / 72.0
            continue
        ls = p.line_spacing if p.line_spacing else 1.15
        total += measure_height(text, size, width_in, ls if isinstance(ls, float) else 1.15)
    return total


def occupied_bbox(sh):
    """(x, y, w, h) of the region actually occupied by visible content.

    Filled shapes occupy their whole box; transparent text boxes occupy only
    the estimated text footprint, so two side-by-side boxes whose text doesn't
    collide are not reported as overlapping.
    """
    x, y, w, h = bbox(sh)
    if has_visible_fill(sh):
        return x, y, w, h
    if not has_text(sh):
        return 0.0, 0.0, 0.0, 0.0
    try:
        tf = sh.text_frame
        max_w = 0.0
        total_h = 0.0
        for p in tf.paragraphs:
            size = 14.0
            text = ""
            for r in p.runs:
                if r.font.size is not None:
                    size = r.font.size.pt
                text += r.text
            eff = sum(1.0 if ord(c) > 0x2E80 else 0.55 for c in text)
            line_w = min(w, max(0.0, eff * size / 72.0))
            max_w = max(max_w, line_w)
            cpl = max(1.0, (w * 72.0) / size)
            lines = max(1, math.ceil(eff / cpl))
            ls = p.line_spacing if p.line_spacing else 1.15
            mult = ls if isinstance(ls, float) else 1.15
            total_h += lines * size * 1.22 * mult / 72.0
        return x, y, min(w, max_w), min(h, total_h)
    except Exception:
        return x, y, w, h


def _contains(a, b):
    """True when b's occupied box lies fully inside a's."""
    return (a[0] <= b[0] and a[1] <= b[1]
            and a[0] + a[2] >= b[0] + b[2] and a[1] + a[3] >= b[1] + b[3])


def check_slide(slide, sidx, deck_slide, issues, warnings, page_w, page_h):
    shapes = list(slide.shapes)
    visible = [sh for sh in shapes if has_visible_fill(sh) or has_text(sh)]

    # 1) overlap among non-placeholder shapes (on occupied regions)
    for i in range(len(shapes)):
        for j in range(i + 1, len(shapes)):
            a, b = shapes[i], shapes[j]
            if getattr(a, "is_placeholder", False) or getattr(b, "is_placeholder", False):
                continue
            oa, ob = occupied_bbox(a), occupied_bbox(b)
            if oa[2] <= 0 or oa[3] <= 0 or ob[2] <= 0 or ob[3] <= 0:
                continue
            ix = max(0.0, min(oa[0] + oa[2], ob[0] + ob[2]) - max(oa[0], ob[0]))
            iy = max(0.0, min(oa[1] + oa[3], ob[1] + ob[3]) - max(oa[1], ob[1]))
            inter = ix * iy
            small = min(oa[2] * oa[3], ob[2] * ob[3])
            if small <= 0 or inter <= 0:
                continue
            # containment: transparent text on a panel, or a small decoration strip
            if _contains(oa, ob) and (not has_visible_fill(b) or
                                      (ob[2] * ob[3]) / (oa[2] * oa[3]) < 0.18):
                continue
            if _contains(ob, oa) and (not has_visible_fill(a) or
                                      (oa[2] * oa[3]) / (ob[2] * ob[3]) < 0.18):
                continue
            if inter / small > 0.28:
                issues.append(f"第 {sidx} 页: 形状重叠 ({_n(inter)} in²) 于 "
                              f"({_n(oa[0])}, {_n(oa[1])}) 与 ({_n(ob[0])}, {_n(ob[1])})")

    # 2) out of bounds
    for sh in shapes:
        x, y, w, h = bbox(sh)
        if w <= 0 and h <= 0:
            continue
        if x < -0.05 or y < -0.05 or x + w > page_w + 0.05 or y + h > page_h + 0.05:
            txt = (sh.text_frame.text[:14] if has_text(sh) else sh.shape_type)
            issues.append(f"第 {sidx} 页: 元素越界「{txt}」({_n(x + w)}, {_n(y + h)}) 超出页面 "
                          f"({_n(page_w)} × {_n(page_h)})")

    # 3) text overflow estimate (Hard beyond 25%+0.15in, Soft otherwise)
    for sh in shapes:
        try:
            if not sh.has_text_frame:
                continue
            tf = sh.text_frame
            if not tf.text.strip():
                continue
            x, y, w, h = bbox(sh)
            if w <= 0 or not tf.word_wrap:
                continue
            need = estimate_text_height(tf, w)
            if need > h * 1.25 + 0.15:
                issues.append(f"第 {sidx} 页: 文本框溢出「{tf.text[:16]}…」"
                              f"估算 {_n(need)} in > 框高 {_n(h)} in")
            elif need > h:
                warnings.append(f"第 {sidx} 页: 文本框接近溢出「{tf.text[:16]}…」"
                                f"估算 {_n(need)} in / 框高 {_n(h)} in")
        except Exception:
            continue

    # 4) empty shapes
    for sh in visible:
        if has_visible_fill(sh) and not has_text(sh):
            # decorative shapes are allowed; only flag boxes with zero size
            x, y, w, h = bbox(sh)
            if w <= 0.01 or h <= 0.01:
                issues.append(f"第 {sidx} 页: 零尺寸形状 ({_n(x)}, {_n(y)})")

    # 5) pictures
    for sh in shapes:
        if sh.shape_type != 13:  # PICTURE
            continue
        try:
            x, y, w, h = bbox(sh)
            nw = _inches(sh.native_width)
            nh = _inches(sh.native_height)
        except Exception:
            continue
        if w <= 0 or h <= 0 or nw <= 0 or nh <= 0:
            continue
        disp_ratio = w / h
        nat_ratio = nw / nh
        drift = abs(disp_ratio - nat_ratio) / nat_ratio
        if drift > 0.10:
            issues.append(f"第 {sidx} 页: 图片纵横比失真（显示 {_n(disp_ratio)} vs 原始 {_n(nat_ratio)}）")
        elif drift > 0.05:
            warnings.append(f"第 {sidx} 页: 图片纵横比略有偏差（{_n(drift * 100)}%）")
        native_px_w = nw * PX_PER_IN
        res_ratio = native_px_w / (w * PX_PER_IN)
        if res_ratio < 0.55:
            issues.append(f"第 {sidx} 页: 图片分辨率不足（原生宽约 {int(native_px_w)}px，"
                          f"显示宽 {int(w * PX_PER_IN)}px）")
        elif res_ratio < 0.7:
            warnings.append(f"第 {sidx} 页: 图片分辨率偏低（原生宽约 {int(native_px_w)}px）")

    # 6) deck signature (layout repetition)
    if deck_slide is not None:
        kinds = [b.get("kind") for b in deck_slide.get("blocks", []) if isinstance(b, dict)]
        return (sidx, kinds)
    return (sidx, None)


def _n(v) -> str:
    return f"{v:.2f}"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--deck")
    ap.add_argument("--strict", action="store_true")
    args = ap.parse_args()

    out_path = Path(args.pptx).expanduser().resolve()
    try:
        prs = Presentation(str(out_path))
    except Exception as e:
        sys.stderr.write(f"打开 {out_path} 失败: {e}\n")
        return 2

    deck = None
    if args.deck:
        try:
            deck = json.loads(Path(args.deck).read_text(encoding="utf-8"))
        except Exception as e:
            sys.stderr.write(f"读取 deck.json 失败: {e}\n")
            return 2

    deck_slides = deck["slides"] if deck else None
    if deck_slides is not None and len(deck_slides) != len(prs.slides):
        sys.stderr.write(f"✗ 页数不一致: deck {len(deck_slides)} 页 vs pptx {len(prs.slides)} 页\n")
        return 1

    issues, warnings = [], []
    signatures = []
    page_w, page_h = _inches(prs.slide_width), _inches(prs.slide_height)
    for i, slide in enumerate(prs.slides, start=1):
        ds = deck_slides[i - 1] if deck_slides else None
        sig = check_slide(slide, i, ds, issues, warnings, page_w, page_h)
        if sig and sig[1]:
            signatures.append(sig)

    # layout repetition between consecutive content slides (no-template rendering)
    for i in range(1, len(signatures)):
        if signatures[i][1] == signatures[i - 1][1] and signatures[i][1]:
            issues.append(f"第 {signatures[i][0]} 页与上一页 block 版式完全相同")

    # report (Hard / Soft grading per references/validation.md)
    lines = []
    lines.append(f"# 校验报告 — {out_path.name}")
    lines.append(f"\n- 总页数: {len(prs.slides)}")
    lines.append(f"- Hard（必须修复）: {len(issues)} | Soft（建议检查）: {len(warnings)}")
    if issues:
        lines.append("\n## Hard（必须修复）")
        for e in issues:
            lines.append(f"- {e}")
    if warnings:
        lines.append("\n## Soft（建议检查；--strict 时也失败）")
        for e in warnings:
            lines.append(f"- {e}")
    if not issues and not warnings:
        lines.append("\n- 未发现机械问题。")
    report_txt = "\n".join(lines)
    report_path = out_path.with_name(f"{out_path.stem}.校验报告.md")
    report_path.write_text(report_txt, encoding="utf-8")

    print(report_txt)
    if issues:
        sys.stderr.write(f"\n共 {len(issues)} 个 Hard 问题，已写入 {report_path}。"
                         "请按 references/validation.md 修复后重新渲染。\n")
        return 1
    if args.strict and warnings:
        sys.stderr.write(f"\n--strict 模式: 共 {len(warnings)} 个 Soft 警告。\n")
        return 1
    print(f"\n校验通过（{len(warnings)} 个 Soft 警告可忽略）: {report_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
