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
    "agenda", "figure", "lead", "insight", "stats", "columns", "steps", "chips",
    "layers", "table", "split", "callouts", "cards", "rows", "compare",
    "pipeline", "hub", "timeline", "points", "formula",
}
NON_STRUCTURE_KINDS = {"lead", "insight", "chips"}
MIN_KINDS = 5
NON_TEXT_KEYS = {"latex", "kind", "type", "src", "image", "path", "file", "layout"}
BARE_DOLLAR_RE = re.compile(r"(?<!\\)\$")


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
        for k in ("text", "title", "body", "note", "kicker", "phase", "metric", "tag"):
            v = b.get(k)
            if isinstance(v, str):
                out.append(v)
        if isinstance(b.get("caption"), str):
            out.append(b["caption"])
        for side in ("left", "right", "side"):
            if isinstance(b.get(side), dict):
                out.extend(_collect_side(b[side]))
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
    types = [s.get("type") for s in slides]
    has_section = "section" in types
    has_catalog = "catalog" in types
    if has_section and not has_catalog:
        errs.append("存在 section 但没有 catalog 目录页")
    if has_catalog and not has_section:
        errs.append("存在 catalog 目录页但没有对应 section 章节页（目录声明了几章，就必须有几章）")
    catalog_items = []
    for s in slides:
        if s.get("type") == "catalog":
            for b in s.get("blocks", []):
                if b.get("kind") == "agenda":
                    catalog_items = [i.get("title", "").strip() for i in b.get("items", [])]
    if has_section:
        sec_items = [s.get("title", "").strip() for s in slides if s.get("type") == "section"]
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
        blocks = s.get("blocks", [])
        for b in blocks:
            if any(k in b for k in ("color", "fontSize", "hex")):
                errs.append(f"第 {idx + 1} 页 block 携带了视觉字段 color/fontSize/hex，禁止")
            kind = b.get("kind") if isinstance(b, dict) else None
            if kind not in known_kinds:
                errs.append(f"第 {idx + 1} 页出现未知 block kind「{kind}」")
            elif kind == "formula" and not str(b.get("latex", "")).strip():
                errs.append(f"第 {idx + 1} 页 formula block 缺少 latex")
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
