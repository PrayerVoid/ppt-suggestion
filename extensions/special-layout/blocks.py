"""特殊排版 block：核心数字、非对称、引文、斜切、关系与层级指标。

粘进渲染器副本（紧挨 BLOCK_RENDERERS 之前），再登记：
    "bignum": _bignum, "asym": _asym, "quote": _quote,
    "diagonal": _diagonal, "radial": _radial, "ring": _ring,
    "marginalia": _marginalia, "concentric": _concentric,
    "quadrant": _quadrant,

依赖渲染器副本里已有的名字（不要再 import）：
    常量   TEXT SUBTLE FAINT ACCENT ACCENT_DK ACCENT_LT LINE PANEL WHITE
    样式   MSO_ANCHOR PP_ALIGN MSO_CONNECTOR MSO_SHAPE
    函数   add_text add_rect add_hline text_height no_shadow
    工具   Inches Pt
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


def _quote(slide, x, y, w, b, style):
    text = str(b.get("text", "")).strip()
    if not text:
        return 0.8
    mark = str(b.get("mark", "“"))
    body_w = w - 0.8
    th = max(0.9, text_height(text, 24, body_w, 1.2))
    add_text(slide, x, y - 0.05, 0.45, 0.55, mark, size=48, bold=True, color=ACCENT,
             wrap=False)
    add_text(slide, x + 0.55, y + 0.12, body_w, th + 0.1, text, size=24, bold=True,
             color=TEXT, line_spacing=1.2)
    source = str(b.get("source", "")).strip()
    h = max(1.5, 0.12 + th)
    if source:
        add_hline(slide, x + 0.55, x + 2.0, y + h + 0.22, color=LINE, w=0.5)
        add_text(slide, x + 0.55, y + h + 0.32, body_w, 0.25, source, size=11, color=FAINT)
        h += 0.55
    return h


def _diagonal(slide, x, y, w, b, style):
    """白底斜切构图：把大色面作为方向提示，文字仍落在留白区。"""
    side = b.get("side", "right")
    h = max(2.65, min(3.6, CONTENT_BOTTOM - y))
    wedge_w = w * 0.38
    if side == "left":
        wedge_x = x
        text_x, text_w = x + w * 0.38, w * 0.56
        rotation = 0
    else:
        wedge_x = x + w - wedge_w
        text_x, text_w = x, w * 0.56
        rotation = 180
    wedge = add_rect(slide, wedge_x, y - 0.18, wedge_w, h + 0.36, ACCENT_LT,
                     shape=MSO_SHAPE.RIGHT_TRIANGLE)
    wedge.rotation = rotation
    kicker = b.get("kicker", "")
    if kicker:
        add_text(slide, text_x, y + 0.35, text_w, 0.3, kicker, size=11, bold=True,
                 color=ACCENT_DK)
    title = b.get("title", "")
    ty = y + 0.75 if kicker else y + 0.45
    th = text_height(title, 25, text_w, 1.1) if title else 0.0
    if title:
        add_text(slide, text_x, ty, text_w, th + 0.1, title, size=25, bold=True,
                 line_spacing=1.1)
    body = b.get("body", "")
    by = ty + th + 0.35
    if body:
        bh = text_height(body, 14, text_w, 1.25)
        add_text(slide, text_x, by, text_w, bh + 0.1, body, size=14, color=SUBTLE,
                 line_spacing=1.25)
    note = b.get("note", "")
    if note:
        add_text(slide, text_x, y + h - 0.45, text_w, 0.25, note, size=11, color=FAINT)
    return h


def _link(slide, x1, y1, x2, y2):
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1),
                                      Inches(x2), Inches(y2))
    conn.line.color.rgb = LINE
    conn.line.width = Pt(0.75)
    no_shadow(conn)
    return conn


def _radial(slide, x, y, w, b, style):
    center = b.get("center", {})
    items = b.get("items", [])[:6]
    if not items:
        return 0
    n = len(items)
    cx, cy = x + w / 2, y + 1.55
    cw, ch = 1.9, 0.78
    nw, nh = (1.75 if n <= 4 else 1.55), 1.1
    rx = min(3.7, w * 0.34)
    ry = 1.2
    positions = []
    for i in range(n):
        angle = math.radians(-90 + i * 360 / n)
        positions.append((cx + rx * math.cos(angle) - nw / 2,
                          cy + ry * math.sin(angle) - nh / 2))
    for px, py in positions:
        _link(slide, cx, cy, px + nw / 2, py + nh / 2)
    add_rect(slide, cx - cw / 2, cy - ch / 2, cw, ch, ACCENT,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(slide, cx - cw / 2 + 0.15, cy - ch / 2 + 0.12, cw - 0.3, 0.22,
             center.get("kicker", ""), size=10, bold=True, color=WHITE)
    add_text(slide, cx - cw / 2 + 0.15, cy - ch / 2 + 0.35, cw - 0.3, 0.28,
             center.get("title", ""), size=13, bold=True, color=WHITE)
    for (px, py), item in zip(positions, items):
        add_rect(slide, px, py, nw, nh, PANEL, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(slide, px + 0.12, py + 0.1, nw - 0.24, 0.22,
                 item.get("title", ""), size=11, bold=True)
        body = item.get("body", "")
        if body:
            add_text(slide, px + 0.12, py + 0.34, nw - 0.24, nh - 0.4, body,
                     size=9, color=SUBTLE, line_spacing=1.05)
    return 3.15


def _ring(slide, x, y, w, b, style):
    value = str(b.get("value", "")).strip()
    if not value:
        return 0.8
    outer = min(2.35, max(1.8, w * 0.22))
    cx, cy = x + outer / 2 + 0.15, y + outer / 2 + 0.1
    add_rect(slide, cx - outer / 2, cy - outer / 2, outer, outer, None,
             line=LINE, shape=MSO_SHAPE.OVAL)
    add_rect(slide, cx - outer * 0.39, cy - outer * 0.39, outer * 0.78, outer * 0.78,
             ACCENT_LT, shape=MSO_SHAPE.OVAL)
    add_text(slide, cx - outer * 0.25, cy - 0.16, outer * 0.5, 0.32, value,
             size=22, bold=True, color=ACCENT_DK, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, wrap=False)
    tx = x + outer + 0.55
    tw = w - outer - 0.55
    ty = y + 0.08
    kicker = b.get("kicker", "")
    if kicker:
        add_text(slide, tx, ty, tw, 0.25, kicker, size=11, bold=True, color=FAINT)
        ty += 0.35
    title = b.get("title", "")
    if title:
        th = text_height(title, 18, tw)
        add_text(slide, tx, ty, tw, th + 0.05, title, size=18, bold=True)
        ty += th + 0.18
    body = b.get("body", "")
    if body:
        bh = text_height(body, 13, tw)
        add_text(slide, tx, ty, tw, bh + 0.05, body, size=13, color=SUBTLE,
                 line_spacing=1.2)
        ty += bh + 0.15
    for item in b.get("items", [])[:3]:
        label = item.get("label", "") if isinstance(item, dict) else ""
        text = item.get("text", "") if isinstance(item, dict) else str(item)
        add_text(slide, tx, ty, tw, 0.2, label, size=10, bold=True, color=ACCENT_DK)
        add_text(slide, tx + 0.8, ty, tw - 0.8, 0.2, text, size=10, color=SUBTLE)
        ty += 0.28
    return max(outer + 0.2, ty - y)


def _marginalia(slide, x, y, w, b, style):
    """主文加边注：边注只承载定义、边界或补充证据。"""
    main = b.get("main", {})
    notes = b.get("notes", [])[:3]
    gap = 0.45
    note_w = min(3.0, w * 0.25)
    main_w = w - note_w - gap
    my = y
    if main.get("kicker"):
        add_text(slide, x, my, main_w, 0.3, main["kicker"], size=11, bold=True, color=FAINT)
        my += 0.38
    title = str(main.get("title", ""))
    if title:
        th = text_height(title, 25, main_w, 1.1)
        add_text(slide, x, my, main_w, th + 0.1, title, size=25, bold=True, line_spacing=1.1)
        my += th + 0.3
    body = str(main.get("body", ""))
    if body:
        bh = text_height(body, 14, main_w, 1.25)
        add_text(slide, x, my, main_w, bh + 0.1, body, size=14, color=SUBTLE, line_spacing=1.25)
        my += bh + 0.1
    h = max(2.0, my - y)
    nx = x + main_w + gap
    ny = y + 0.05
    for note in notes:
        if not isinstance(note, dict):
            continue
        label = str(note.get("label", "")).strip()
        text = str(note.get("text", "")).strip()
        if label:
            add_text(slide, nx, ny, note_w, 0.25, label, size=10, bold=True, color=ACCENT_DK)
            ny += 0.3
        if text:
            nh = text_height(text, 12, note_w, 1.2)
            add_text(slide, nx, ny, note_w, nh + 0.05, text, size=12, color=SUBTLE, line_spacing=1.2)
            ny += nh + 0.25
        h = max(h, ny - y)
    add_rect(slide, nx - gap * 0.5, y, 0.02, max(2.0, h), LINE)
    return h


def _concentric(slide, x, y, w, b, style):
    """同心层级：环内只放短标签，解释放在右侧。"""
    center = b.get("center", {})
    rings = b.get("rings", [])[:3]
    if not rings:
        return 0.8
    diagram_w = min(4.3, w * 0.38)
    cx, cy = x + diagram_w * 0.48, y + 1.35
    sizes = [2.85, 2.05, 1.25][:len(rings)]
    ring_colors = [ACCENT, ACCENT_DK, LINE]
    for size, color in zip(sizes, ring_colors):
        add_rect(slide, cx - size / 2, cy - size / 2, size, size, None,
                 line=color, line_w=1.1, shape=MSO_SHAPE.OVAL)
    if center.get("title"):
        add_text(slide, cx - 0.48, cy - 0.18, 0.96, 0.36, str(center["title"]),
                 size=12, bold=True, color=ACCENT_DK, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
    if center.get("kicker"):
        add_text(slide, x, y, diagram_w, 0.25, str(center["kicker"]),
                 size=10, bold=True, color=FAINT, align=PP_ALIGN.CENTER)
    tx = x + diagram_w + 0.55
    tw = w - diagram_w - 0.55
    ty = y + 0.12
    for i, ring in enumerate(rings):
        if not isinstance(ring, dict):
            continue
        label = str(ring.get("label", ring.get("title", ""))).strip()
        text = str(ring.get("text", "")).strip()
        if label:
            add_text(slide, tx, ty, tw, 0.32, label, size=14, bold=True, color=ACCENT_DK)
            ty += 0.3
        if text:
            rh = text_height(text, 13, tw, 1.2)
            add_text(slide, tx, ty, tw, rh + 0.05, text, size=13, color=SUBTLE, line_spacing=1.2)
            ty += rh + 0.12
        if i < len(rings) - 1:
            add_hline(slide, tx, tx + tw, ty + 0.08, color=LINE, w=0.5)
            ty += 0.22
    return max(2.8, min(4.1, ty - y))


def _quadrant(slide, x, y, w, b, style):
    """四象限判断：四项必须互斥，中心只保留一个短标题。"""
    items = [it for it in b.get("items", [])[:4] if isinstance(it, dict)]
    if len(items) != 4:
        return 0.8
    gap = 0.28
    col_w = (w - gap) / 2
    box_h = 1.25
    row_gap = 0.38
    total_h = box_h * 2 + row_gap
    for i, item in enumerate(items):
        col, row = i % 2, i // 2
        bx = x + col * (col_w + gap)
        by = y + row * (box_h + row_gap)
        add_rect(slide, bx, by, col_w, box_h, PANEL, line=LINE, line_w=0.5,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        key = str(item.get("key", "")).strip()
        title = str(item.get("title", "")).strip()
        body = str(item.get("body", "")).strip()
        if key:
            add_text(slide, bx + 0.18, by + 0.14, col_w - 0.36, 0.3, key,
                     size=10, bold=True, color=ACCENT_DK)
        if title:
            add_text(slide, bx + 0.18, by + 0.42, col_w - 0.36, 0.35, title,
                     size=14, bold=True)
        if body:
            add_text(slide, bx + 0.18, by + 0.74, col_w - 0.36, 0.5, body,
                     size=11, color=SUBTLE, line_spacing=1.1)
    center = b.get("center", {})
    title = str(center.get("title", "")).strip()
    if title:
        size = 0.72
        add_rect(slide, x + w / 2 - size / 2, y + total_h / 2 - size / 2,
                 size, size, ACCENT, shape=MSO_SHAPE.OVAL)
        add_text(slide, x + w / 2 - size / 2 + 0.06, y + total_h / 2 - 0.16,
                 size - 0.12, 0.32, title, size=10, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return total_h
