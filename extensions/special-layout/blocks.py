"""特殊排版 block：bignum（超大数字结论）与 asym（7:3 非对称分栏）。

粘进渲染器副本（紧挨 BLOCK_RENDERERS 之前），再登记：
    "bignum": _bignum, "asym": _asym,

依赖渲染器副本里已有的名字（不要再 import）：
    常量   TEXT SUBTLE FAINT ACCENT ACCENT_DK ACCENT_LT LINE PANEL
    样式   MSO_ANCHOR PP_ALIGN
    函数   add_text add_rect add_hline text_height
签名与 _stats/_split 一致：fn(slide, x, y, w, b, style) -> 本 block 占用高度（英寸）。
"""


def _point_texts(items):
    return [it if isinstance(it, str) else it.get("text", "") for it in items]


def _bignum(slide, x, y, w, b, style):
    lw = w * 0.5
    rx = x + lw + 0.6
    rw = w - lw - 0.6
    ly = y
    if b.get("kicker"):
        add_text(slide, x, ly, lw, 0.3, b["kicker"], size=11, bold=True, color=FAINT)
        ly += 0.4
    add_text(slide, x, ly, lw, 1.35, str(b.get("value", "")), size=72, bold=True,
             color=ACCENT_DK, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
    ly += 1.5
    t = b.get("title", "")
    if t:
        th = text_height(t, 18, lw)
        add_text(slide, x, ly, lw, th + 0.1, t, size=18, bold=True)
        ly += th + 0.2
    body = b.get("body", "")
    if body:
        bh = text_height(body, 14, lw)
        add_text(slide, x, ly, lw, bh + 0.1, body, size=14, color=SUBTLE, line_spacing=1.25)
        ly += bh + 0.1
    add_rect(slide, x + lw + 0.3, y, 0.02, max(ly - y, 2.4), LINE)
    ry = y + 0.1
    for i, text in enumerate(_point_texts(b.get("items", []))[:3]):
        ih = max(0.8, text_height(text, 14, rw - 0.7) + 0.35)
        add_text(slide, rx, ry, 0.5, 0.3, f"{i + 1:02d}", size=13, bold=True, color=ACCENT)
        add_text(slide, rx + 0.6, ry, rw - 0.6, ih - 0.2, text, size=14, color=TEXT, line_spacing=1.25)
        add_hline(slide, rx, rx + rw, ry + ih - 0.1, color=LINE, w=0.5)
        ry += ih
    return max(ly - y, ry - y, 2.4)


def _asym(slide, x, y, w, b, style):
    left = b.get("left", {})
    right = b.get("right", {})
    gap = 0.35
    lw = (w - gap) * 0.7
    rw = (w - gap) * 0.3
    ly = y
    if left.get("kicker"):
        add_text(slide, x, ly, lw, 0.3, left["kicker"], size=11, bold=True, color=FAINT)
        ly += 0.35
    t = left.get("title", "")
    if t:
        th = sum(text_height(seg, 24, lw) for seg in t.split("\n")) + 0.1 * t.count("\n")
        add_text(slide, x, ly, lw, th + 0.2, t, size=24, bold=True, line_spacing=1.1)
        ly += th + 0.35
    body = left.get("body", "")
    if body:
        bh = text_height(body, 14, lw)
        add_text(slide, x, ly, lw, bh + 0.1, body, size=14, color=SUBTLE, line_spacing=1.25)
        ly += bh + 0.1
    items = _point_texts(right.get("items", []))
    ihs = [max(0.5, text_height(it, 13, rw - 0.4) + 0.16) for it in items]
    ph = 0.55 + sum(ihs) + 0.25
    h = max(ly - y, ph, 2.2)
    rx = x + lw + gap
    add_rect(slide, rx, y, rw, h, PANEL)
    if right.get("kicker"):
        add_text(slide, rx + 0.2, y + 0.15, rw - 0.4, 0.3, right["kicker"], size=11, bold=True,
                 color=ACCENT_DK)
    ry = y + 0.55
    for it, ih in zip(items, ihs):
        add_text(slide, rx + 0.2, ry, rw - 0.4, ih, it, size=13, color=TEXT)
        ry += ih
    return h
