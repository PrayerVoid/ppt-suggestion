#!/usr/bin/env python3
"""Structural validation of deck.json for ppt-suggestion (render-time gate).

Checks the quality checklist in references/deck-schema.md §4. Run before
render_pptx.py; the renderer also imports validate_deck() as a hard gate.

Usage:
  python deck_check.py deck.json [--extra-kinds kind1,kind2] [--template-mode]

--extra-kinds 放行渲染器副本里新增的 block kind（如特殊版式拓展）。
--template-mode 模板模式：版式由模板决定，跳过字数、版式种类、相邻页等版式类检查，
                仅保留结构合法性、字段与行内公式 `$` 配对检查。

Exit codes:
  0 valid
  1 validation errors (printed, one per line)
  2 deck.json unreadable / malformed
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path

BUILTIN_KINDS = {
    "agenda", "chart", "figure", "lead", "insight", "stats", "columns", "steps", "chips",
    "layers", "table", "split", "callouts", "cards", "rows", "compare",
    "pipeline", "hub", "timeline", "points", "formula",
}
NON_STRUCTURE_KINDS = {"lead", "insight", "chips"}
MIN_KINDS = 5
PAGE_TYPES = {"cover", "catalog", "section", "content", "close"}
PAGE_REQUIRED = {
    "cover": ("title",),
    "catalog": ("title", "blocks"),
    "section": ("title",),
    "content": ("title", "blocks"),
    "close": ("title",),
}
NON_TEXT_KEYS = {"latex", "kind", "type", "src", "image", "path", "file", "layout"}
BARE_DOLLAR_RE = re.compile(r"(?<!\\)\$")
CHART_TYPES = {
    "column_clustered", "column_stacked", "column_stacked_100", "bar_clustered",
    "bar_stacked", "line", "line_markers", "area", "area_stacked", "pie",
    "doughnut", "radar", "radar_filled", "xy_scatter", "xy_scatter_lines", "bubble",
}
XY_CHART_TYPES = {"xy_scatter", "xy_scatter_lines", "bubble"}


def _finite_number(value):
    try:
        return math.isfinite(float(value))
    except (TypeError, ValueError):
        return False


def _check_chart(block, slide_no):
    errs = []
    chart_type = str(block.get("chart_type", "")).strip().lower()
    fallback = str(block.get("fallback", "auto")).strip().lower()
    series = block.get("series")
    if not chart_type:
        errs.append(f"第 {slide_no} 页 chart 缺少 chart_type")
    elif chart_type not in CHART_TYPES and fallback != "svg":
        errs.append(f"第 {slide_no} 页 chart 类型「{chart_type}」不在原生白名单，需明确 fallback=svg")
    if fallback not in {"auto", "native", "svg"}:
        errs.append(f"第 {slide_no} 页 chart fallback 必须是 auto/native/svg")
    if not isinstance(series, list) or not series:
        return errs + [f"第 {slide_no} 页 chart 至少需要一个 series"]
    names = []
    if chart_type in XY_CHART_TYPES:
        for si, item in enumerate(series, start=1):
            if not isinstance(item, dict) or not str(item.get("name", "")).strip():
                errs.append(f"第 {slide_no} 页 chart 第 {si} 个 series 缺少 name")
            else:
                names.append(item["name"])
            points = item.get("points") if isinstance(item, dict) else None
            if not isinstance(points, list) or not points:
                errs.append(f"第 {slide_no} 页 chart 第 {si} 个 series 缺少 points")
                continue
            for pi, point in enumerate(points, start=1):
                needed = ("x", "y", "size") if chart_type == "bubble" else ("x", "y")
                if not isinstance(point, dict) or any(k not in point for k in needed):
                    errs.append(f"第 {slide_no} 页 chart 第 {si} 个点 {pi} 字段不完整")
                    continue
                if any(not _finite_number(point[k]) for k in needed):
                    errs.append(f"第 {slide_no} 页 chart 第 {si} 个点 {pi} 含非有限数值")
                if chart_type == "bubble" and float(point["size"]) <= 0:
                    errs.append(f"第 {slide_no} 页 bubble 第 {si} 个点 {pi} 的 size 必须为正")
    else:
        categories = block.get("categories")
        if not isinstance(categories, list) or not categories:
            errs.append(f"第 {slide_no} 页 chart 缺少 categories")
            categories = []
        if len(categories) == 1:
            errs.append(f"第 {slide_no} 页 chart 只有一个数据点，不能表达比较关系；改用 stats/lead，或补齐可比类别")
        for si, item in enumerate(series, start=1):
            if not isinstance(item, dict) or not str(item.get("name", "")).strip():
                errs.append(f"第 {slide_no} 页 chart 第 {si} 个 series 缺少 name")
                continue
            names.append(item["name"])
            values = item.get("values")
            if not isinstance(values, list) or len(values) != len(categories):
                errs.append(f"第 {slide_no} 页 chart 第 {si} 个 series 长度必须等于 categories")
            elif any(not _finite_number(v) for v in values):
                errs.append(f"第 {slide_no} 页 chart 第 {si} 个 series 含非有限数值")
            elif chart_type in {"pie", "doughnut"}:
                numeric = [float(v) for v in values]
                if any(v < 0 for v in numeric):
                    errs.append(f"第 {slide_no} 页 {chart_type} 的数值不能为负")
                elif not any(v > 0 for v in numeric):
                    errs.append(f"第 {slide_no} 页 {chart_type} 至少需要一个正数扇区")
    if len(names) != len(set(names)):
        errs.append(f"第 {slide_no} 页 chart 的 series name 必须唯一")
    has_context = (str(block.get("source", "")).strip()
                   or bool(block.get("table_view"))
                   or bool(str(block.get("accessibility_note", "")).strip()))
    if not has_context:
        errs.append(f"第 {slide_no} 页 chart 需要 source、table_view 或 accessibility_note")
    if block.get("secondary_axis"):
        errs.append(f"第 {slide_no} 页 chart 禁止使用 secondary_axis")
    return errs


def _collect_side(d: dict) -> list:
    out = []
    for k in ("title", "body", "kicker", "phase", "metric"):
        if isinstance(d.get(k), str):
            out.append(d[k])
    for itm in d.get("items", []):
        if isinstance(itm, str):
            out.append(itm)
        elif isinstance(itm, dict) and isinstance(itm.get("text"), str):
            out.append(itm["text"])
    return out


def collect_text(s: dict) -> list:
    out = []
    for b in s.get("blocks", []):
        if not isinstance(b, dict):
            continue
        for k in ("text", "title", "subtitle", "body", "note", "source", "kicker", "phase", "metric", "tag", "unit", "table_view", "accessibility_note"):
            v = b.get(k)
            if isinstance(v, str):
                out.append(v)
        if isinstance(b.get("caption"), str):
            out.append(b["caption"])
        for side in ("left", "right", "side", "center", "main"):
            if isinstance(b.get(side), dict):
                out.extend(_collect_side(b[side]))
        for key in ("notes", "rings"):
            for it in b.get(key, []):
                if isinstance(it, dict):
                    for kk in ("label", "title", "body", "text", "key"):
                        if isinstance(it.get(kk), str):
                            out.append(it[kk])
        for it in b.get("items", []):
            if isinstance(it, str):
                out.append(it)
            elif isinstance(it, dict):
                for kk in ("title", "body", "note", "text", "num", "label", "value",
                           "phase", "metric", "kicker"):
                    if isinstance(it.get(kk), str):
                        out.append(it[kk])
                for m in it.get("modules", []):
                    if isinstance(m, dict):
                        for kk in ("title", "body"):
                            if isinstance(m.get(kk), str):
                                out.append(m[kk])
        if b.get("kind") == "table":
            for row in b.get("rows", []):
                out.extend(row)
            out.extend(b.get("headers", []))
    return out


def _iter_strings(obj, path=""):
    """递归取出 JSON 里所有文本字段及其路径；latex/kind/路径类字段不是正文，跳过。"""
    if isinstance(obj, str):
        yield path, obj
    elif isinstance(obj, dict):
        for k, v in obj.items():
            if k not in NON_TEXT_KEYS:
                yield from _iter_strings(v, f"{path}.{k}" if path else k)
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            yield from _iter_strings(v, f"{path}[{i}]")


def _check_inline_math(slides) -> list:
    """行内公式 `$...$` 的 `$` 必须成对；字面美元符号写成 `\\$`。"""
    errs = []
    for idx, s in enumerate(slides):
        for path, text in _iter_strings(s):
            n = len(BARE_DOLLAR_RE.findall(text))
            if n % 2:
                errs.append(f"第 {idx + 1} 页 {path} 的 $ 数量为奇数（{n}），行内公式未配对；字面美元符号写成 \\$：「{text[:30]}」")
            elif "$$" in text.replace("\\$", ""):
                errs.append(f"第 {idx + 1} 页 {path} 含空的行内公式 $$：「{text[:30]}」")
    return errs


def validate_deck(deck: dict, extra_kinds=(), template_mode=False) -> list:
    """Return a list of error strings; empty means valid."""
    errs = []
    known_kinds = BUILTIN_KINDS | set(extra_kinds)
    slides = deck.get("slides")
    if not isinstance(slides, list) or not slides:
        return ["slides 为空或缺失"]
    errs += _check_inline_math(slides)
    types = []
    for idx, slide in enumerate(slides, start=1):
        if not isinstance(slide, dict):
            errs.append(f"第 {idx} 页必须是对象")
            continue
        page_type = str(slide.get("type", "")).strip().lower()
        types.append(page_type)
        if page_type not in PAGE_TYPES:
            errs.append(f"第 {idx} 页 type「{page_type or '<空>'}」不是支持的页面类型")
            continue
        for field in PAGE_REQUIRED[page_type]:
            value = slide.get(field)
            if field == "blocks":
                valid = isinstance(value, list)
            else:
                valid = isinstance(value, str) and bool(value.strip())
            if not valid:
                errs.append(f"第 {idx} 页 {page_type} 缺少必填字段 {field}")
    has_section = "section" in types
    has_catalog = "catalog" in types
    if has_section and not has_catalog:
        errs.append("存在 section 但没有 catalog 目录页")
    if has_catalog and not has_section:
        errs.append("存在 catalog 目录页但没有对应 section 章节页（目录声明了几章，就必须有几章）")
    catalog_items = []
    for s in slides:
        if not isinstance(s, dict) or s.get("type") != "catalog":
            continue
        blocks = s.get("blocks", [])
        if not isinstance(blocks, list):
            continue
        for b in blocks:
            if isinstance(b, dict) and b.get("kind") == "agenda":
                catalog_items = [i.get("title", "").strip()
                                 for i in b.get("items", [])
                                 if isinstance(i, dict)]
    if has_section:
        sec_items = [s.get("title", "").strip() for s in slides
                     if isinstance(s, dict) and s.get("type") == "section"]
        if len(sec_items) != len(catalog_items):
            errs.append(f"目录 {len(catalog_items)} 条与 section {len(sec_items)} 个数量不一致")
        else:
            for i, (a, b) in enumerate(zip(catalog_items, sec_items)):
                if a != b:
                    errs.append(f"目录第 {i + 1} 条「{a}」与 section「{b}」标题不一致")
    cards_pages = 0
    prev_cardish = False
    seen_kinds = set()
    for idx, s in enumerate(slides):
        if not isinstance(s, dict):
            continue
        blocks = s.get("blocks", [])
        if not isinstance(blocks, list):
            continue
        for b in blocks:
            if not isinstance(b, dict):
                errs.append(f"第 {idx + 1} 页 block 必须是对象")
                continue
            if any(k in b for k in ("color", "fontSize", "hex")):
                errs.append(f"第 {idx + 1} 页 block 携带了视觉字段 color/fontSize/hex，禁止")
            kind = b.get("kind") if isinstance(b, dict) else None
            if kind not in known_kinds:
                errs.append(f"第 {idx + 1} 页出现未知 block kind「{kind}」")
            elif kind == "formula" and not str(b.get("latex", "")).strip():
                errs.append(f"第 {idx + 1} 页 formula block 缺少 latex")
            elif kind == "chart":
                errs.extend(_check_chart(b, idx + 1))
        if s.get("type") == "content" and not template_mode:
            total = sum(len(str(x)) for x in collect_text(s))
            if total < 110:
                errs.append(f"第 {idx + 1} 页正文总字数 {total} < 110")
            kinds = [b.get("kind") for b in blocks if isinstance(b, dict)]
            seen_kinds.update(kinds)
            if kinds and set(kinds) <= {"lead", "insight", "chips"}:
                errs.append(f"第 {idx + 1} 页只放了 lead/insight/chips，没有结构")
            for b in blocks:
                for it in b.get("items", []):
                    if isinstance(it, dict) and "body" in it:
                        body = (it.get("body") or "").strip()
                        if len(body) < 10:
                            errs.append(f"第 {idx + 1} 页「{it.get('title', '')}」的 body 只有 {len(body)} 字")
            cardish = kinds and set(kinds) & {"cards", "columns"}
            if cardish:
                cards_pages += 1
                if prev_cardish:
                    errs.append(f"第 {idx + 1} 页与上一页连续使用 cards/columns")
            prev_cardish = bool(cardish)
        if s.get("type") == "content" and len(s.get("title", "")) < 8:
            errs.append(f"第 {idx + 1} 页标题过短（{len(s.get('title',''))} 字），疑似关键词标题，应为结论句")
    if template_mode:
        return errs
    if cards_pages > 2:
        errs.append(f"cards 页共 {cards_pages} 页 > 2")
    struct_kinds = {"layers", "steps", "rows", "pipeline", "hub", "timeline", "compare"}
    if len(set(seen_kinds) & struct_kinds) < 1 and any(t == "content" for t in types):
        errs.append("全稿没有结构页（layers/steps/rows/pipeline/hub/timeline/compare）")
    counted = seen_kinds - NON_STRUCTURE_KINDS
    if any(t == "content" for t in types) and len(counted) < MIN_KINDS:
        errs.append(f"全稿版式种类过少（{len(counted)} 种，lead/insight/chips 不计），至少 {MIN_KINDS} 种")
    return errs


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--extra-kinds", default="")
    ap.add_argument("--template-mode", action="store_true")
    args = ap.parse_args()
    extra = {k.strip() for k in args.extra_kinds.split(",") if k.strip()}
    path = Path(args.deck)
    try:
        deck = json.loads(path.read_text(encoding="utf-8"))
    except Exception as e:
        sys.stderr.write(f"deck.json 读取失败: {e}\n")
        return 2
    errs = validate_deck(deck, extra, args.template_mode)
    if errs:
        for e in errs:
            sys.stderr.write(f"✗ {e}\n")
        sys.stderr.write("请按 references/deck-schema.md §4 修好 deck.json 再渲染，禁止跳过。\n")
        return 1
    print("deck.json 校验通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
