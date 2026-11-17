"""特殊构图：斜切、短节点环绕、指标圆环与同心层级。

复制到当前核心渲染器副本的 BLOCK_RENDERERS 之前，仅登记：
    diagonal: _diagonal, radial: _radial, ring: _ring, concentric: _concentric
日常版式由核心渲染器提供，不在本模块重复登记。
依赖当前渲染器的 math、常量、形状工具与 _layout_items；不要重复导入。
"""


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
    wedge = add_rect(slide, wedge_x, y, wedge_w, h, ACCENT_LT,
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
    items = _layout_items(b, "items", 3, 6, "radial")
    if not items:
        return 0
    n = len(items)
    cx, cy = x + w / 2, y + 1.8
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
    return 3.6



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
    add_text(slide, cx - outer * 0.25, cy - 0.25, outer * 0.5, 0.5, value,
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
    for item in _layout_items(b, "items", 0, 3, "ring"):
        label = item.get("label", "") if isinstance(item, dict) else ""
        text = item.get("text", "") if isinstance(item, dict) else str(item)
        add_text(slide, tx, ty, tw, 0.2, label, size=10, bold=True, color=ACCENT_DK)
        add_text(slide, tx + 0.8, ty, tw - 0.8, 0.2, text, size=10, color=SUBTLE)
        ty += 0.28
    return max(outer + 0.2, ty - y)



def _concentric(slide, x, y, w, b, style):
    """同心层级：环内只放短标签，解释放在右侧。"""
    center = b.get("center", {})
    rings = _layout_items(b, "rings", 2, 3, "concentric")
    if not rings:
        return 0.8
    diagram_w = min(4.3, w * 0.38)
    cx, cy = x + diagram_w * 0.48, y + 1.7
    sizes = [2.85, 2.05, 1.25][:len(rings)]
    ring_colors = [ACCENT, ACCENT_DK, LINE]
    for i, (size, color) in enumerate(zip(sizes, ring_colors), start=1):
        add_rect(slide, cx - size / 2, cy - size / 2, size, size, None,
                 line=color, line_w=1.1, shape=MSO_SHAPE.OVAL)
        # 与右侧说明同序，避免空环无法对应层级。
        add_rect(slide, cx - 0.16, cy + size / 2 - 0.13, 0.32, 0.26, WHITE)
        add_text(slide, cx - 0.16, cy + size / 2 - 0.13, 0.32, 0.26,
                 str(i), size=10, bold=True, color=ACCENT_DK, align=PP_ALIGN.CENTER)
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
            add_text(slide, tx, ty, tw, 0.32, f"{i + 1}  {label}", size=14, bold=True, color=ACCENT_DK)
            ty += 0.3
        if text:
            rh = text_height(text, 13, tw, 1.2)
            add_text(slide, tx, ty, tw, rh + 0.05, text, size=13, color=SUBTLE, line_spacing=1.2)
            ty += rh + 0.12
        if i < len(rings) - 1:
            add_hline(slide, tx, tx + tw, ty + 0.08, color=LINE, w=0.5)
            ty += 0.22
    return max(3.3, ty - y)
