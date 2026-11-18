"""容量与密度规则：为 26 个核心版式提供可执行的内容边界。

规则是保守的字符预算，不替代最终 PPTX 渲染检查；超过 hard_limit 时应缩短、拆页或换版式。
"""
from __future__ import annotations

LAYOUT_CAPACITY = {
    "agenda": (5, 40, 0),
    "lead": (1, 80, 0), "insight": (1, 120, 0), "chips": (6, 24, 0),
    "points": (4, 42, 18), "callouts": (4, 52, 18), "columns": (4, 48, 18),
    "split": (2, 115, 24), "asym": (2, 150, 28), "marginalia": (3, 105, 18),
    "compare": (3, 70, 20), "table": (8, 28, 12), "quadrant": (4, 52, 18),
    "layers": (5, 48, 16), "rows": (6, 44, 16), "steps": (6, 46, 16),
    "pipeline": (6, 42, 16), "timeline": (7, 44, 16), "hub": (6, 38, 14),
    "bignum": (3, 180, 30), "stats": (5, 42, 12), "chart": (6, 34, 10),
    "quote": (1, 160, 24), "figure": (1, 90, 20), "formula": (1, 90, 0),
    "cards": (4, 52, 18),
}


def _text(item):
    if isinstance(item, str):
        return item
    if isinstance(item, dict):
        return " ".join(str(item.get(k, "")) for k in
                         ("title", "body", "text", "note", "label", "value", "metric"))
    return ""


def check_block(block, slide_no=0):
    kind = str(block.get("kind", "")).lower()
    if kind not in LAYOUT_CAPACITY:
        return []
    max_items, per_item, min_body = LAYOUT_CAPACITY[kind]
    items = block.get("items") or block.get("rows") or block.get("points") or []
    if not isinstance(items, list):
        items = []
    errors = []
    if len(items) > max_items:
        errors.append(f"第 {slide_no} 页 {kind} 条目 {len(items)} > {max_items}，应拆页或换版式")
    for i, item in enumerate(items, 1):
        text = _text(item).strip()
        if len(text) > per_item:
            errors.append(f"第 {slide_no} 页 {kind}[{i}] {len(text)} 字 > {per_item} 字容量，应缩短或换版式")
        # min_body is a planning hint, never a license to fabricate filler.
        # Density diagnostics are returned separately from hard capacity errors.
    return errors


def check_deck(deck):
    errors = []
    for n, slide in enumerate(deck.get("slides", []), 1):
        if not isinstance(slide, dict) or not isinstance(slide.get("blocks", []), list):
            continue
        for block in slide.get("blocks", []):
            if isinstance(block, dict):
                errors.extend(check_block(block, n))
    return errors
