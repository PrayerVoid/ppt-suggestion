#!/usr/bin/env python3
"""Mechanical validation for a rendered PPTX (ppt-suggestion phase 4).

Checks per slide: shape overlap, out-of-bounds, estimated text overflow,
empty shapes, image distortion / low resolution, unconverted inline `$...$`
formulas, layout repetition, and (optionally, with --deck) slide count and
per-slide block signature. With --baseline (template mode) out-of-bounds shapes
that the template itself already has are reported as Info, not Hard.

Usage:
  python validate_pptx.py <out.pptx> [--deck deck.json] [--baseline template.pptx] [--strict]

Exit codes:
  0  no errors (warnings allowed unless --strict)
  1  at least one error
  2  file open / parse failure
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path

try:
    from power_pptx import Presentation
except ImportError:
    sys.stderr.write("power-pptx 未安装，请用 .venv 的 python 运行本脚本（pip install -r requirements.txt）\n")
    raise SystemExit(2)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from text_measure import estimate_width  # noqa: E402

EMU_IN = 914400.0
PX_PER_IN = 96.0


def _inches(v) -> float:
    return (v or 0) / EMU_IN


def is_deco(sh) -> bool:
    return (getattr(sh, "name", "") or "").startswith("DECO_")


def is_formula(sh) -> bool:
    """公式占位框（FORMULA_ 前缀）或已替换的原生公式（含 m:oMath / a14:m）。"""
    if (getattr(sh, "name", "") or "").startswith("FORMULA_"):
        return True
    try:
        return bool(sh._element.xpath('.//*[local-name()="oMath" or local-name()="m"]'))
    except Exception:
        return False


def is_picture(sh) -> bool:
    """原生公式等无预设几何的形状取 shape_type 会抛 NotImplementedError。"""
    try:
        return sh.shape_type == 13
    except NotImplementedError:
        return False


def is_chart(sh) -> bool:
    """GraphicFrame 中的原生 Office 图表。"""
    try:
        return bool(getattr(sh, "has_chart", False)) and getattr(sh, "chart", None) is not None
    except Exception:
        return False


def chart_series_name(series) -> str:
    """读取 power-pptx 未暴露为 .name 的 OfficeCLI c:tx/c:v 系列名。"""
    value = str(getattr(series, "name", "") or "").strip()
    if value:
        return value
    try:
        values = series._element.xpath('.//*[local-name()="tx"]//*[local-name()="v"]/text()')
        return str(values[0]).strip() if values else ""
    except Exception:
        return ""


def bbox(sh):
    return (_inches(getattr(sh, "left", None)), _inches(getattr(sh, "top", None)),
            _inches(getattr(sh, "width", None)), _inches(getattr(sh, "height", None)))


def has_visible_fill(sh) -> bool:
    """True when the shape carries an explicit solidFill (panel/rect)."""
    try:
        sp = sh._element.spPr
        if sp is None:
            return False
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


class PositionedShape:
    """A leaf shape with group-local coordinates mapped to slide EMU."""
    def __init__(self, shape, transform, path):
        self.shape = shape
        self.path = path
        sx, sy, dx, dy = transform
        self.left = dx + sx * (shape.left or 0)
        self.top = dy + sy * (shape.top or 0)
        self.width = sx * (shape.width or 0)
        self.height = sy * (shape.height or 0)

    def __getattr__(self, name):
        return getattr(self.shape, name)


def positioned_shapes(shapes, transform=(1, 1, 0, 0), parent=""):
    """Recurse through groups, including off/ext/chOff/chExt at every level."""
    sx, sy, dx, dy = transform
    for index, shape in enumerate(shapes, 1):
        path = f"{parent}/{shape.name}[{index}]"
        if shape._element.tag.endswith("}grpSp"):
            xfrm = shape._element.grpSpPr.find(f"{A_NS}xfrm")
            if xfrm is None:
                yield from positioned_shapes(shape.shapes, transform, path)
                continue
            def point(tag, first, second, defaults):
                node = xfrm.find(f"{A_NS}{tag}")
                return tuple(float(node.get(key, default)) if node is not None else default
                             for key, default in zip((first, second), defaults))
            ox, oy = point("off", "x", "y", (0, 0))
            ex, ey = point("ext", "cx", "cy", (1, 1))
            cx, cy = point("chOff", "x", "y", (0, 0))
            cex, cey = point("chExt", "cx", "cy", (ex, ey))
            gx, gy = ex / cex if cex else 1, ey / cey if cey else 1
            child = (sx * gx, sy * gy, dx + sx * (ox - cx * gx), dy + sy * (oy - cy * gy))
            yield from positioned_shapes(shape.shapes, child, path)
        else:
            yield PositionedShape(shape, transform, path)


def paragraph_size(paragraph):
    explicit = [run.font.size.pt for run in paragraph.runs if run.font.size is not None]
    if explicit:
        return max(explicit)
    sizes = paragraph._p.xpath('./a:pPr/a:defRPr/@sz | ./a:endParaRPr/@sz')
    return max(float(value) / 100 for value in sizes) if sizes else 14.0


def paragraph_width(paragraph, text):
    tracking = [float(value) / 100 for value in paragraph._p.xpath(
        './a:pPr/a:defRPr/@spc | .//a:rPr/@spc')]
    spacing = max(tracking, default=0)
    return estimate_width(text, paragraph_size(paragraph)) + max(0, len(text) - 1) * spacing / 72


def estimate_text_height(tf, width_in: float) -> float:
    # Font sizes and text insets are points/EMU, even when a group uses a tiny
    # local coordinate system. Only the shape bounds receive the group transform.
    inner_width = max(0.01, width_in - _inches(tf.margin_left) - _inches(tf.margin_right))
    total = _inches(tf.margin_top) + _inches(tf.margin_bottom)
    for paragraph in tf.paragraphs:
        size = paragraph_size(paragraph)
        text = paragraph.text.replace("\v", "\n")
        lines = sum(max(1, math.ceil(paragraph_width(paragraph, line) / inner_width))
                    for line in text.split("\n")) if tf.word_wrap is not False else len(text.split("\n"))
        ls = paragraph.line_spacing
        line_height = ls.pt / 72 if hasattr(ls, "pt") else size * 1.22 * (ls or 1.0) / 72
        total += lines * line_height
        total += _inches(paragraph.space_before) + _inches(paragraph.space_after)
    return total


def occupied_bbox(sh, clamp_h=True):
    """(x, y, w, h) of the region actually occupied by visible content.

    Filled shapes occupy their whole box; transparent text boxes occupy only
    the estimated text footprint, so two side-by-side boxes whose text doesn't
    collide are not reported as overlapping. clamp_h=False keeps the footprint
    of text that spills past the bottom of its box, which is where it really
    collides with whatever sits below.
    """
    x, y, w, h = bbox(sh)
    if is_chart(sh) or has_visible_fill(sh):
        return x, y, w, h
    if not has_text(sh):
        return 0.0, 0.0, 0.0, 0.0
    try:
        tf = sh.text_frame
        left, right = _inches(tf.margin_left), _inches(tf.margin_right)
        top, bottom = _inches(tf.margin_top), _inches(tf.margin_bottom)
        inner = max(0.01, w - left - right)
        max_w = max((paragraph_width(p, line) for p in tf.paragraphs
                     for line in p.text.replace("\v", "\n").split("\n")), default=0)
        text_w = min(inner, max_w) if tf.word_wrap is not False else max_w
        text_h = max(0, estimate_text_height(tf, w) - top - bottom)
        # Respect centered/right-aligned labels and vertically anchored text.
        alignments = {str(p.alignment).split(" (")[0] for p in tf.paragraphs}
        if alignments == {"CENTER"}:
            left += max(0, inner - text_w) / 2
        elif alignments == {"RIGHT"}:
            left += max(0, inner - text_w)
        anchor = str(tf.vertical_anchor).split(" (")[0]
        slack = max(0, h - top - bottom - text_h)
        top += slack / 2 if anchor == "MIDDLE" else slack if anchor == "BOTTOM" else 0
        return x + left, y + top, text_w, min(max(0, h - top), text_h) if clamp_h else text_h
    except Exception:
        return x, y, w, h


def _contains(a, b):
    """True when b's occupied box lies fully inside a's."""
    return (a[0] <= b[0] and a[1] <= b[1]
            and a[0] + a[2] >= b[0] + b[2] and a[1] + a[3] >= b[1] + b[3])


INLINE_FORMULA_RE = re.compile(r"(?<!\\)\$[^$\n\v]+?(?<!\\)\$")
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def find_unconverted_inline(slide):
    """段落文字里仍带成对 `$...$` 的片段（含组合形状、表格单元格）。"""
    found = []
    for sh in slide.shapes:
        if is_formula(sh):
            continue
        for p in sh._element.iter(f"{A_NS}p"):
            text = "".join(t.text or "" for r in p.findall(f"{A_NS}r") for t in r.findall(f"{A_NS}t"))
            m = INLINE_FORMULA_RE.search(text)
            if m:
                found.append(m.group(0)[:20])
    return found


def oob_key(sh):
    """比对键：形状名 + 取整 bbox（模板页复制后页序会变，不含页号）。"""
    x, y, w, h = bbox(sh)
    return (getattr(sh, "name", "") or "", round(x, 1), round(y, 1), round(w, 1), round(h, 1))


def overlap_key(a, b):
    # Recognize the same decoration pattern after a template group is translated.
    # Include actual text for text/text pairs: new text must never inherit an
    # exemption for an old template collision.
    ordered = sorted((a, b), key=lambda sh: (sh.name, bbox(sh)))
    first, second = ordered
    fa, fb = bbox(first), bbox(second)
    pattern = tuple((sh.name, round(box[2], 1), round(box[3], 1))
                    for sh, box in zip(ordered, (fa, fb)))
    texts = tuple(sh.text_frame.text for sh in ordered) if all(has_text(sh) for sh in ordered) else ()
    return ("overlap", pattern, round(round(fb[0] - fa[0], 4), 2), round(round(fb[1] - fa[1], 4), 2), texts)


def is_out_of_bounds(sh, page_w, page_h):
    x, y, w, h = bbox(sh)
    if w <= 0 and h <= 0:
        return False
    return x < -0.05 or y < -0.05 or x + w > page_w + 0.05 or y + h > page_h + 0.05


def collect_baseline(prs):
    """模板自带的越界 / 重叠问题键集合，供 --baseline 把它们降为说明项。"""
    seen = set()
    pw, ph = _inches(prs.slide_width), _inches(prs.slide_height)
    for i, slide in enumerate(prs.slides, start=1):
        check_slide(slide, i, None, [], [], pw, ph, seen=seen)
    return frozenset(seen)


def check_slide(slide, sidx, deck_slide, issues, warnings, page_w, page_h,
                baseline=frozenset(), infos=None, seen=None, check_formulas=True):
    shapes = list(positioned_shapes(slide.shapes))
    visible = [sh for sh in shapes if has_visible_fill(sh) or has_text(sh)]

    def report(key, msg):
        """越界 / 重叠：键在 baseline（模板自带）里就降为说明项。"""
        if seen is not None:
            seen.add(key)
        if infos is not None and key in baseline:
            infos.append(msg)
        else:
            issues.append(msg)

    # 1) overlap among non-placeholder shapes (on occupied regions)
    for i in range(len(shapes)):
        for j in range(i + 1, len(shapes)):
            a, b = shapes[i], shapes[j]
            if getattr(a, "is_placeholder", False) or getattr(b, "is_placeholder", False):
                continue
            if is_deco(a) or is_deco(b):
                continue
            oa, ob = occupied_bbox(a), occupied_bbox(b)
            if (has_text(a) and has_text(b)
                    and not has_visible_fill(a) and not has_visible_fill(b)):
                oa, ob = occupied_bbox(a, clamp_h=False), occupied_bbox(b, clamp_h=False)
            if oa[2] <= 0 or oa[3] <= 0 or ob[2] <= 0 or ob[3] <= 0:
                continue
            ix = max(0.0, min(oa[0] + oa[2], ob[0] + ob[2]) - max(oa[0], ob[0]))
            iy = max(0.0, min(oa[1] + oa[3], ob[1] + ob[3]) - max(oa[1], ob[1]))
            inter = ix * iy
            small = min(oa[2] * oa[3], ob[2] * ob[3])
            if small <= 0 or inter <= 0:
                continue
            # containment: transparent text on a panel, or a small decoration strip
            if (_contains(oa, ob) and (has_visible_fill(a) or not has_text(a))
                    and (not has_visible_fill(b) or (ob[2] * ob[3]) / (oa[2] * oa[3]) < 0.18)):
                continue
            if (_contains(ob, oa) and (has_visible_fill(b) or not has_text(b))
                    and (not has_visible_fill(a) or (oa[2] * oa[3]) / (ob[2] * ob[3]) < 0.18)):
                continue
            if inter / small > 0.28:
                report(overlap_key(a, b),
                       f"第 {sidx} 页: 形状重叠 ({_n(inter)} in²) 于 "
                       f"({_n(oa[0])}, {_n(oa[1])}) 与 ({_n(ob[0])}, {_n(ob[1])})")

    # 2) out of bounds（模板自带的越界形状只作说明，不算 Hard）
    for sh in shapes:
        if not is_out_of_bounds(sh, page_w, page_h):
            continue
        x, y, w, h = bbox(sh)
        txt = sh.text_frame.text[:14] if has_text(sh) else sh.name
        report(("oob", oob_key(sh)),
               f"第 {sidx} 页: 元素越界「{txt}」({_n(x + w)}, {_n(y + h)}) 超出页面 "
               f"({_n(page_w)} × {_n(page_h)})")

    # 2b) native chart GraphicFrame
    for sh in shapes:
        if not is_chart(sh):
            continue
        x, y, w, h = bbox(sh)
        if w <= 0 or h <= 0:
            issues.append(f"第 {sidx} 页: 图表零尺寸 ({_n(x)}, {_n(y)})")
            continue
        chart = sh.chart
        series = list(getattr(chart, "series", ()) or ())
        if not series:
            issues.append(f"第 {sidx} 页: 原生图表没有数据系列")
        else:
            for idx, item in enumerate(series, start=1):
                if not chart_series_name(item):
                    issues.append(f"第 {sidx} 页: 原生图表第 {idx} 个系列缺少名称")
                values = tuple(getattr(item, "values", ()) or ())
                if not values:
                    issues.append(f"第 {sidx} 页: 原生图表第 {idx} 个系列没有数值")
                else:
                    try:
                        finite = all(math.isfinite(float(value)) for value in values)
                    except (TypeError, ValueError):
                        finite = False
                    if not finite:
                        issues.append(f"第 {sidx} 页: 原生图表第 {idx} 个系列含非有限数值")
            if len(series) > 1 and not bool(getattr(chart, "has_legend", False)):
                issues.append(f"第 {sidx} 页: 多系列原生图表缺少图例")
        chart_type = str(getattr(chart, "chart_type", "") or "").strip()
        if not chart_type:
            issues.append(f"第 {sidx} 页: 原生图表缺少图表类型")
        if bool(getattr(chart, "has_title", False)):
            try:
                title = chart.chart_title
                if not title.has_text_frame or not title.text_frame.text.strip():
                    issues.append(f"第 {sidx} 页: 原生图表标题为空")
            except Exception:
                issues.append(f"第 {sidx} 页: 原生图表标题无法读取")

    # deck 有对应 chart block 时，再核对原生图表的类型和系列数。
    if deck_slide is not None:
        blocks = [b for b in deck_slide.get("blocks", [])
                  if isinstance(b, dict) and b.get("kind") == "chart"]
        native_blocks = [b for b in blocks if str(b.get("fallback", "auto")).lower() != "svg"]
        native_shapes = [sh for sh in shapes if is_chart(sh)]
        if any(str(b.get("fallback", "auto")).lower() == "native" for b in native_blocks):
            if len(native_shapes) < len([b for b in native_blocks
                                         if str(b.get("fallback", "auto")).lower() == "native"]):
                issues.append(f"第 {sidx} 页: deck 要求原生图表，但 PPTX 中缺少 GraphicFrame")
        if len(native_blocks) == 1 and len(native_shapes) == 1:
            block = native_blocks[0]
            actual = str(getattr(native_shapes[0].chart, "chart_type", "")).split(" (", 1)[0].strip().lower()
            expected = str(block.get("chart_type", "")).strip().lower()
            if actual and expected and actual != expected:
                issues.append(f"第 {sidx} 页: 图表类型为「{actual}」，与 deck 的「{expected}」不一致")
            expected_count = len(block.get("series", []))
            actual_count = len(list(native_shapes[0].chart.series))
            if expected_count != actual_count:
                issues.append(f"第 {sidx} 页: 图表系列数为 {actual_count}，与 deck 的 {expected_count} 不一致")

    # 2c) inline formulas that formula.py has not converted yet
    for frag in find_unconverted_inline(slide) if check_formulas else ():
        issues.append(f"第 {sidx} 页: 公式未转换，请先跑 formula.py（残留「{frag}」）")

    # 3) Text inside groups is checked in slide coordinates. Overflow is never
    # baseline-exempt: changing only the text can break an unchanged template box.
    for sh in shapes:
        if not sh.has_text_frame or is_formula(sh) or is_deco(sh):
            continue
        tf = sh.text_frame
        if not tf.text.strip():
            continue
        x, y, w, h = bbox(sh)
        if w <= 0 or h <= 0:
            issues.append(f"第 {sidx} 页: 文字容器零尺寸「{tf.text[:16]}」")
            continue
        try:
            need = estimate_text_height(tf, w)
            inner_width = max(0.01, w - _inches(tf.margin_left) - _inches(tf.margin_right))
            if tf.word_wrap is False:
                widest = max((paragraph_width(p, line) for p in tf.paragraphs
                              for line in p.text.replace("\v", "\n").split("\n")), default=0)
                if widest > inner_width + 0.04:
                    issues.append(f"第 {sidx} 页: 不换行文字横向溢出「{tf.text[:16]}」"
                                  f"文字宽 {_n(widest)} in > 可用宽 {_n(inner_width)} in")
            if need > h + max(0.04, h * 0.10):
                issues.append(f"第 {sidx} 页: 文本框溢出「{tf.text[:16]}…」"
                              f"估算 {_n(need)} in > 框高 {_n(h)} in（{sh.path}）")
            elif need > h + 0.01:
                warnings.append(f"第 {sidx} 页: 文本框接近溢出「{tf.text[:16]}…」"
                                f"估算 {_n(need)} in / 框高 {_n(h)} in")
        except Exception as exc:
            issues.append(f"第 {sidx} 页: 无法检查文字容器 {sh.path}: {exc}")

    # 4) empty shapes
    for sh in visible:
        if has_visible_fill(sh) and not has_text(sh):
            # decorative shapes are allowed; only flag boxes with zero size
            x, y, w, h = bbox(sh)
            if w <= 0.01 or h <= 0.01:
                issues.append(f"第 {sidx} 页: 零尺寸形状 ({_n(x)}, {_n(y)})")
    for sh in shapes:
        if is_chart(sh):
            x, y, w, h = bbox(sh)
            if w <= 0.01 or h <= 0.01:
                issues.append(f"第 {sidx} 页: 图表 GraphicFrame 零尺寸「{sh.name}」")

    # 5) pictures
    for sh in shapes:
        if not is_picture(sh):
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
    ap.add_argument("--baseline", help="模板 .pptx；与之相同的越界形状不计入 Hard")
    ap.add_argument("--strict", action="store_true")
    args = ap.parse_args()

    out_path = Path(args.pptx).expanduser().resolve()
    try:
        prs = Presentation(str(out_path))
    except Exception as e:
        sys.stderr.write(f"打开 {out_path} 失败: {e}\n")
        return 2

    baseline = frozenset()
    if args.baseline:
        try:
            baseline = collect_baseline(Presentation(str(Path(args.baseline).expanduser().resolve())))
        except Exception as e:
            sys.stderr.write(f"打开 baseline {args.baseline} 失败: {e}\n")
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

    from pptx_package import package_problems
    package_issues = package_problems(out_path)
    # Extended-property counts are advisory metadata, not a broken slide or
    # relationship. Some producers leave them stale even in readable files.
    warnings = [problem for problem in package_issues if problem.startswith("docProps/app.xml:")]
    issues = [problem for problem in package_issues if problem not in warnings]
    infos = []
    signatures = []
    page_w, page_h = _inches(prs.slide_width), _inches(prs.slide_height)
    for i, slide in enumerate(prs.slides, start=1):
        ds = deck_slides[i - 1] if deck_slides else None
        sig = check_slide(slide, i, ds, issues, warnings, page_w, page_h, baseline, infos)
        if sig and sig[1]:
            signatures.append(sig)

    # layout repetition between consecutive content slides (no-template rendering)
    for i in range(1, len(signatures)):
        same = signatures[i][1] == signatures[i - 1][1] and signatures[i][1]
        chart_only = same == ["chart"]
        if same and not chart_only:
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
    if infos:
        lines.append(f"\n## 模板自带，已忽略 {len(infos)} 项（越界 / 重叠）")
        for e in infos:
            lines.append(f"- {e}")
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
