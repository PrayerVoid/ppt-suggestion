# -*- coding: utf-8 -*-
"""decor_templates.py — 几何装饰模板库（power_pptx 执行层）

41 个参数化模板：13 个经典派生形（v2 样例已验证）+ 28 个新纹样家族。
统一接口：tpl_xxx(slide, x, y, w, h, pal=None, **opts)
  - 坐标单位 px（画布 1280×720），模板自动缩放适配 (x,y,w,h) 区域
  - pal 为配色 dict：main/sub/accent/ink/gray/card/hair
  - 每个模板只依赖基本几何（圆/方/三角/弧）+ 重叠/布尔/旋转/网格四种组合技法

依赖：power_pptx, lxml。统一由 geometry-deco/add_deco.py 调用。
"""
import math
import random
from power_pptx.util import Emu, Pt
from power_pptx.dml.color import RGBColor
from power_pptx.enum.shapes import MSO_SHAPE
from power_pptx.oxml.ns import qn
from lxml import etree

PX = 9525
PAL = dict(main="1B4B9B", sub="C1272D", accent="F2C500",
           ink="1A1A1A", gray="666660", card="F5F5F2", hair="D9D9D4")


def rgb(h):
    return RGBColor.from_string(h)


# ---------------------------------------------------------------- 基础助手
def shape(slide, prst, x, y, w, h, fill=None, alpha=None, stroke=None, sw=1.0, rot=None):
    sp = slide.shapes.add_shape(prst, Emu(int(x * PX)), Emu(int(y * PX)),
                                Emu(int(w * PX)), Emu(int(h * PX)))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = rgb(fill)
        if alpha is not None:
            clr = sp._element.spPr.find(qn('a:solidFill')).find(qn('a:srgbClr'))
            etree.SubElement(clr, qn('a:alpha')).set('val', str(int(alpha * 100000)))
    if stroke is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = rgb(stroke)
        sp.line.width = Pt(sw)
    if rot is not None:
        sp.rotation = rot
    st = sp._element.find(qn('p:style'))
    if st is not None:
        sp._element.remove(st)          # 去掉主题阴影/效果引用
    sp.shadow.inherit = False
    return sp


def custgeom(sp, w, h, paths):
    """把 preset 形状替换为自定义路径。cmds: move/line/arc(wR,hR,stAng,swAng)/close，角度 1/60000 度。"""
    spPr = sp._element.spPr
    old = spPr.find(qn('a:prstGeom'))
    if old is not None:
        spPr.remove(old)
    cg = spPr.makeelement(qn('a:custGeom'), {})
    spPr.find(qn('a:xfrm')).addnext(cg)
    for tag in ('a:avLst', 'a:gdLst', 'a:ahLst', 'a:cxnLst'):
        etree.SubElement(cg, qn(tag))
    rect = etree.SubElement(cg, qn('a:rect'))
    rect.set('l', '0'); rect.set('t', '0')
    rect.set('r', str(int(w * PX))); rect.set('b', str(int(h * PX)))
    pl = etree.SubElement(cg, qn('a:pathLst'))
    for path in paths:
        p = etree.SubElement(pl, qn('a:path'))
        p.set('w', str(int(w * PX))); p.set('h', str(int(h * PX)))
        for cmd in path:
            if cmd[0] == 'move':
                e = etree.SubElement(p, qn('a:moveTo'))
                pt = etree.SubElement(e, qn('a:pt'))
                pt.set('x', str(int(cmd[1] * PX))); pt.set('y', str(int(cmd[2] * PX)))
            elif cmd[0] == 'line':
                e = etree.SubElement(p, qn('a:lnTo'))
                pt = etree.SubElement(e, qn('a:pt'))
                pt.set('x', str(int(cmd[1] * PX))); pt.set('y', str(int(cmd[2] * PX)))
            elif cmd[0] == 'arc':
                e = etree.SubElement(p, qn('a:arcTo'))
                e.set('wR', str(int(cmd[1] * PX))); e.set('hR', str(int(cmd[2] * PX)))
                e.set('stAng', str(int(cmd[3]))); e.set('swAng', str(int(cmd[4])))
            elif cmd[0] == 'close':
                etree.SubElement(p, qn('a:close'))


def lens(r, d):
    """双弓形（两等圆交集）：w=2r-d, h=√(4r²-d²)。返回 (w,h,paths)。"""
    w = 2 * r - d
    h = 2 * math.sqrt(r * r - d * d / 4)
    th = math.degrees(math.acos(d / (2 * r)))
    return w, h, [[('move', w / 2, 0),
                   ('arc', r, r, int(-th * 60000), int(2 * th * 60000)),
                   ('arc', r, r, int((180 - th) * 60000), int(2 * th * 60000)),
                   ('close',)]]


def dot(slide, cx, cy, d, fill, alpha=None):
    return shape(slide, MSO_SHAPE.OVAL, cx - d / 2, cy - d / 2, d, d, fill=fill, alpha=alpha)


def ring(slide, cx, cy, r, stroke, sw=1.0, alpha=None):
    return shape(slide, MSO_SHAPE.OVAL, cx - r, cy - r, 2 * r, 2 * r, stroke=stroke, sw=sw)


def seg(slide, x1, y1, x2, y2, color, sw=1.0, alpha=None):
    """线段 = 细矩形旋转。"""
    sw = max(sw, 1.2)
    L = math.hypot(x2 - x1, y2 - y1)
    if L < 0.5:
        return None
    ang = math.degrees(math.atan2(y2 - y1, x2 - x1))
    return shape(slide, MSO_SHAPE.RECTANGLE, (x1 + x2) / 2 - L / 2, (y1 + y2) / 2 - sw / 2,
                 L, sw, fill=color, alpha=alpha, rot=ang)


def poly(slide, x, y, w, h, pts, fill=None, alpha=None, stroke=None, sw=1.0):
    """多边形（pts 为区域内局部 px 坐标）。"""
    path = [('move', pts[0][0], pts[0][1])]
    path += [('line', px, py) for px, py in pts[1:]]
    path += [('close',)]
    sp = shape(slide, MSO_SHAPE.RECTANGLE, x, y, w, h, fill=fill, alpha=alpha, stroke=stroke, sw=sw)
    custgeom(sp, w, h, [path])
    return sp


def hex_pts(cx, cy, s):
    """尖顶六边形顶点（y 轴向下）。"""
    return [(cx + s * math.cos(math.radians(30 + 60 * k)),
             cy + s * math.sin(math.radians(30 + 60 * k))) for k in range(6)]


def hex_centers(x, y, w, h, s, mode):
    """六边形网格中心，只保留完整落在区域内的。mode: kagome / honey。"""
    out = []
    if mode == 'kagome':
        a = 3 * math.sqrt(3) / 2 * s
        dx, dy = a, a * math.sqrt(3) / 2
    else:
        dx, dy = math.sqrt(3) * s, 1.5 * s
    hw, hh = math.sqrt(3) / 2 * s, s          # 六边形半宽/半高
    row = 0
    cy = y + hh + 2
    while cy + hh <= y + h - 2:
        off = (dx / 2) if (row % 2 and mode == 'honey') else ((dx / 2) if row % 2 else 0)
        cx = x + hw + 2 + off
        while cx + hw <= x + w - 2:
            out.append((cx, cy))
            cx += dx
        cy += dy
        row += 1
    return out


def clip_seg(x1, y1, x2, y2, x, y, w, h):
    """Liang-Barsky 线段裁剪到矩形，返回端点或 None。"""
    dx, dy = x2 - x1, y2 - y1
    t0, t1 = 0.0, 1.0
    for p, q in ((-dx, x1 - x), (dx, x + w - x1), (-dy, y1 - y), (dy, y + h - y1)):
        if abs(p) < 1e-9:
            if q < 0:
                return None
        else:
            t = q / p
            if p < 0:
                t0 = max(t0, t)
            else:
                t1 = min(t1, t)
    if t0 > t1:
        return None
    return (x1 + t0 * dx, y1 + t0 * dy, x1 + t1 * dx, y1 + t1 * dy)


def P(pal, key):
    return (pal or PAL)[key]


# ================================================================ 经典派生形（13）
def tpl_vesica(slide, x, y, w, h, pal=None, d_ratio=1.0, color=None):
    """双弓形：两等圆相交，交集即花瓣。d_ratio=d/r：0.8 胖 / 1.0 经典 / 1.2 瘦。"""
    c = color or P(pal, 'main')
    r = min(w / (4 - d_ratio), h / 2) * 0.96
    d = d_ratio * r
    cx, cy = x + w / 2, y + h / 2
    x0 = cx - (4 * r - d) / 2
    shape(slide, MSO_SHAPE.OVAL, x0, cy - r, 2 * r, 2 * r, stroke=P(pal, 'gray'), sw=1.2)
    shape(slide, MSO_SHAPE.OVAL, x0 + 2 * r - d, cy - r, 2 * r, 2 * r, stroke=P(pal, 'gray'), sw=1.2)
    lw, lh, paths = lens(r, d)
    sp = shape(slide, MSO_SHAPE.RECTANGLE, cx - lw / 2, cy - lh / 2, lw, lh, fill=c)
    custgeom(sp, lw, lh, paths)


def tpl_rosette(slide, x, y, w, h, pal=None, n=6, d_ratio=1.0, color=None):
    """玫瑰窗：n 片双弓形花瓣旋转对称。"""
    c = color or P(pal, 'main')
    r = min(w, h) / 2 * 0.92
    cx, cy = x + w / 2, y + h / 2
    for k in range(n):
        ang = -90 + k * 360 / n
        a = math.radians(ang)
        lw, lh, paths = lens(r, d_ratio * r)
        pcx, pcy = cx + lh / 2 * math.cos(a), cy + lh / 2 * math.sin(a)
        sp = shape(slide, MSO_SHAPE.RECTANGLE, pcx - lw / 2, pcy - lh / 2, lw, lh, fill=c, rot=ang + 90)
        custgeom(sp, lw, lh, paths)
    dot(slide, cx, cy, r * 0.14, P(pal, 'sub'))


def tpl_seed_of_life(slide, x, y, w, h, pal=None, color=None):
    """生命之胚：7 等圆线稿重叠（圆心距=r）。"""
    c = color or P(pal, 'main')
    r = min(w, h) / 4 * 0.95
    cx, cy = x + w / 2, y + h / 2
    ring(slide, cx, cy, r, c, 1.2)
    for k in range(6):
        a = math.radians(k * 60)
        ring(slide, cx + r * math.cos(a), cy + r * math.sin(a), r, c, 1.2)


def tpl_moon(slide, x, y, w, h, pal=None, color=None):
    """月牙 = 圆 − 圆（preset moon）。"""
    s = min(w, h) * 0.8
    shape(slide, MSO_SHAPE.MOON, x + (w - s) / 2, y + (h - s) / 2, s, s,
          fill=color or P(pal, 'ink'))


def tpl_octagram(slide, x, y, w, h, pal=None, color=None):
    """八角星 = 方 ∪ 方45°（同色相并）。"""
    c = color or P(pal, 'main')
    s = min(w, h) * 0.68
    x0, y0 = x + (w - s) / 2, y + (h - s) / 2
    shape(slide, MSO_SHAPE.RECTANGLE, x0, y0, s, s, fill=c)
    shape(slide, MSO_SHAPE.RECTANGLE, x0, y0, s, s, fill=c, rot=45)


def tpl_octagon_emergence(slide, x, y, w, h, pal=None, color=None, alpha=0.35):
    """八边形 = 方 ∩ 方45°（低透明交叠，交集涌现）。"""
    c = color or P(pal, 'sub')
    s = min(w, h) * 0.68
    x0, y0 = x + (w - s) / 2, y + (h - s) / 2
    shape(slide, MSO_SHAPE.RECTANGLE, x0, y0, s, s, fill=c, alpha=alpha)
    shape(slide, MSO_SHAPE.RECTANGLE, x0, y0, s, s, fill=c, alpha=alpha, rot=45)


def tpl_hexagram(slide, x, y, w, h, pal=None, color=None):
    """六角星 = 三角 ∪ 倒三角（倒置须下移 h/3 对齐重心）。"""
    c = color or P(pal, 'accent')
    s = min(w, h) * 0.78
    x0, y0 = x + (w - s) / 2, y + (h - s) / 2 - s / 6
    shape(slide, MSO_SHAPE.ISOSCELES_TRIANGLE, x0, y0, s, s, fill=c)
    shape(slide, MSO_SHAPE.ISOSCELES_TRIANGLE, x0, y0 + s / 3, s, s, fill=c, rot=180)


def tpl_seigaiha(slide, x, y, w, h, pal=None, color=None):
    """青海波：同心半圆扇，隔行错半位、后行遮前行。"""
    c = color or P(pal, 'main')
    bg = P(pal, 'card')
    R = min(w / 10, h / 4.4)
    rows = 3
    cols = int(w / (2 * R)) + 1
    y0 = y + h - R * 1.15
    for row in range(rows - 1, -1, -1):
        for col in range(cols):
            fcx = x + R + col * 2 * R + (row % 2) * R
            fcy = y0 - row * R
            if fcx - R < x - 1 or fcx + R > x + w + 1:
                continue
            for rad, fl in zip([R, .8 * R, .6 * R, .4 * R, .2 * R], [c, bg, c, bg, c]):
                sp = shape(slide, MSO_SHAPE.RECTANGLE, fcx - rad, fcy - rad, 2 * rad, rad, fill=fl)
                custgeom(sp, 2 * rad, rad,
                         [[('move', 0, rad), ('arc', rad, rad, 180 * 60000, 180 * 60000), ('close',)]])


def tpl_pinwheel(slide, x, y, w, h, pal=None, color=None):
    """风车：4 个四分圆共心留隙（g≈0.17R，否则拼成整圆）。"""
    c = color or P(pal, 'sub')
    R = min(w, h) / 2 * 0.88
    g = 0.17 * R
    cx, cy = x + w / 2, y + h / 2
    for k in range(4):
        th = math.radians(k * 90 + 45)
        bcx = cx + (g + R / math.sqrt(2)) * math.cos(th)
        bcy = cy + (g + R / math.sqrt(2)) * math.sin(th)
        sp = shape(slide, MSO_SHAPE.RECTANGLE, bcx - R / 2, bcy - R / 2, R, R, fill=c, rot=k * 90)
        custgeom(sp, R, R, [[('move', 0, 0), ('line', R, 0), ('arc', R, R, 0, 90 * 60000), ('close',)]])


def tpl_scallop(slide, x, y, w, h, pal=None, color=None):
    """扇贝边：半圆相切排列，可替代发丝分隔线。"""
    c = color or P(pal, 'main')
    c2 = P(pal, 'accent')
    r = h * 0.45
    n = int(w / (2.05 * r))
    x0 = x + (w - n * 2.05 * r) / 2 + r
    for i in range(n):
        sp = shape(slide, MSO_SHAPE.RECTANGLE, x0 + i * 2.05 * r - r, y + h / 2 - r, 2 * r, r,
                   fill=c if i % 2 == 0 else c2)
        custgeom(sp, 2 * r, r, [[('move', 0, r), ('arc', r, r, 180 * 60000, 180 * 60000), ('close',)]])


def tpl_shippo(slide, x, y, w, h, pal=None, color=None):
    """七宝つなぎ：方格圆 + 隙心圆 → 四向花瓣（r = s/√2）。"""
    c = color or P(pal, 'main')
    s = min(w / 5.5, h / 4.5)
    r = s / math.sqrt(2)
    nx, ny = int(w / s), int(h / s)
    x0, y0 = x + (w - nx * s) / 2 + s / 2, y + (h - ny * s) / 2 + s / 2
    for i in range(nx):
        for j in range(ny):
            ring(slide, x0 + i * s, y0 + j * s, r, c, 1.0)
    for i in range(nx - 1):
        for j in range(ny - 1):
            ring(slide, x0 + (i + .5) * s, y0 + (j + .5) * s, r, c, 1.0)


def tpl_truchet(slide, x, y, w, h, pal=None, color=None, seed=7):
    """Truchet 瓦：两条四分圆弧 × 随机直角旋转（随机只用在朝向）。"""
    c = color or P(pal, 'ink')
    T = min(w / 9, h / 4)
    nx, ny = int(w / T), int(h / T)
    x0, y0 = x + (w - nx * T) / 2, y + (h - ny * T) / 2
    random.seed(seed)
    for row in range(ny):
        for col in range(nx):
            sp = shape(slide, MSO_SHAPE.RECTANGLE, x0 + col * T, y0 + row * T, T, T,
                       stroke=c, sw=1.4, rot=random.choice([0, 90]))
            custgeom(sp, T, T, [
                [('move', T / 2, 0), ('arc', T / 2, T / 2, 0, 90 * 60000)],
                [('move', T / 2, T), ('arc', T / 2, T / 2, 180 * 60000, 90 * 60000)],
            ])


def tpl_reuleaux(slide, x, y, w, h, pal=None, color=None):
    """勒洛三角：三圆交集，3 段 60° 弧（半径=边长）。"""
    c = color or P(pal, 'main')
    S = min(w, h) * 0.82
    y0 = S - math.sqrt(3) / 2 * S
    x0, yy = x + (w - S) / 2, y + (h - S) / 2 + (S - y0) / 2 - y0 / 2
    paths = [[('move', 0, y0),
              ('arc', S, S, 240 * 60000, 60 * 60000),
              ('arc', S, S, 0, 60 * 60000),
              ('arc', S, S, 120 * 60000, 60 * 60000),
              ('close',)]]
    sp = shape(slide, MSO_SHAPE.RECTANGLE, x0, yy, S, S, fill=c)
    custgeom(sp, S, S, paths)


# ================================================================ 圆派生（4）
def tpl_laurel(slide, x, y, w, h, pal=None, color=None, leaves=12):
    """月桂花环：双弓形叶片沿圆弧切向排布，顶部留口。"""
    c = color or P(pal, 'main')
    R = min(w, h) / 2 * 0.8
    cx, cy = x + w / 2, y + h / 2
    lr = R * 0.24
    lw, lh, paths = lens(lr, 0.8 * lr)
    for k in range(leaves):
        ang = -30 + k * 240 / (leaves - 1)
        a = math.radians(ang)
        pcx, pcy = cx + R * math.cos(a), cy + R * math.sin(a)
        sp = shape(slide, MSO_SHAPE.RECTANGLE, pcx - lw / 2, pcy - lh / 2, lw, lh,
                   fill=c, rot=ang + 90)
        custgeom(sp, lw, lh, paths)
    dot(slide, cx, cy + R, R * 0.1, P(pal, 'sub'))


def tpl_cloud(slide, x, y, w, h, pal=None, color=None):
    """云 = 圆之并：同色圆簇重叠，轮廓即云。"""
    c = color or P(pal, 'main')
    r = min(w / 5.2, h / 2.3)
    cx, cy = x + w / 2, y + h / 2 + r * 0.25
    for dx, dy, rr in [(-1.5, 0, 1), (-0.4, -0.55, 1.15), (0.8, -0.3, 0.95),
                       (1.7, 0.1, 0.8), (-0.2, 0.35, 1.05), (1.0, 0.4, 0.85)]:
        dot(slide, cx + dx * r, cy + dy * r, 2 * rr * r, c)


def tpl_ring_chain(slide, x, y, w, h, pal=None, color=None, n=5):
    """连环：等圆描边水平扣连（间距 1.3r）。"""
    c = color or P(pal, 'main')
    r = min(h / 2.4, w / (2 + (n - 1) * 1.3))
    sw = max(1.5, r * 0.16)
    cx0 = x + w / 2 - (n - 1) * 1.3 * r / 2
    cy = y + h / 2
    for i in range(n):
        ring(slide, cx0 + i * 1.3 * r, cy, r, c, sw)


def tpl_guilloche(slide, x, y, w, h, pal=None, color=None, n=14):
    """玑镂玫瑰窗：n 个等圆描边，圆心偏移 e≈0.35r 旋转排布（万花尺效果）。"""
    c = color or P(pal, 'main')
    r = min(w, h) / 2 * 0.6
    e = 0.35 * r
    cx, cy = x + w / 2, y + h / 2
    for k in range(n):
        a = math.radians(k * 360 / n)
        ring(slide, cx + e * math.cos(a), cy + e * math.sin(a), r, c, 1.0)


# ================================================================ 弧与螺旋（4）
def tpl_golden_spiral(slide, x, y, w, h, pal=None, color=None):
    """黄金螺旋：斐波那契正方形 [1,1,2,3,5,8] + 内切四分圆弧，弧首尾相接。"""
    c = color or P(pal, 'main')
    fib = [1, 1, 2, 3, 5, 8]
    sc = min(w / 13, h / 8) * 0.94
    ox, oy = x + (w - 13 * sc) / 2 + 3 * sc, y + (h - 8 * sc) / 2 + 5 * sc
    bx, by, bw, bh = 0.0, 0.0, 1.0, 1.0
    dirs = ['R', 'D', 'L', 'U', 'R']
    # 弧心角规则：种子 BR(180°)；R→BL(270°) D→TL(0°) L→TR(90°) U→BR(180°)
    arc_cfg = {'R': (0, 1, 270), 'D': (0, 0, 0), 'L': (1, 0, 90), 'U': (1, 1, 180)}
    squares = [(0, 0, 1)]
    for i, f in enumerate(fib[1:]):
        d = dirs[i]
        if d == 'R':
            sx, sy = bx + bw, by
            bx, bw = bx, bw + f
        elif d == 'D':
            sx, sy = bx, by + bh
            by, bh = by, bh + f
        elif d == 'L':
            sx, sy = bx - f, by
            bx, bw = bx - f, bw + f
        else:
            sx, sy = bx, by - f
            by, bh = by - f, bh + f
        squares.append((sx, sy, f))
    for idx, (sx, sy, f) in enumerate(squares):
        X, Y, S = ox + sx * sc, oy + sy * sc, f * sc
        shape(slide, MSO_SHAPE.RECTANGLE, X, Y, S, S, stroke=P(pal, 'gray'), sw=0.9, alpha=None)
        if idx == 0:
            ccx, ccy, st = 1, 1, 180
        else:
            fx, fy, st = arc_cfg[dirs[idx - 1]]
            ccx, ccy = fx, fy
        a0 = math.radians(st)
        p0 = (ccx + math.cos(a0), ccy + math.sin(a0))
        sp = shape(slide, MSO_SHAPE.RECTANGLE, X, Y, S, S, stroke=c, sw=2.2)
        custgeom(sp, S, S, [[('move', p0[0] * S, p0[1] * S),
                             ('arc', S, S, st * 60000, 90 * 60000)]])
    dot(slide, ox + 0.5 * sc, oy + 0.5 * sc, 4, P(pal, 'sub'))


def tpl_rainbow(slide, x, y, w, h, pal=None, colors=None):
    """虹带：同心半圆环（custGeom 双弧），三色递进。"""
    cs = colors or [P(pal, 'main'), P(pal, 'accent'), P(pal, 'sub')]
    R = min(w / 2, h) * 0.95
    cx, cy = x + w / 2, y + h
    t = R / 3.4
    for i, c in enumerate(cs):
        Ro, Ri = R - i * t, R - i * t - t * 0.8
        sp = shape(slide, MSO_SHAPE.RECTANGLE, cx - Ro, cy - Ro, 2 * Ro, Ro, fill=c)
        custgeom(sp, 2 * Ro, Ro, [[
            ('move', 0, Ro), ('arc', Ro, Ro, 180 * 60000, 180 * 60000),
            ('line', Ro + Ri, Ro), ('arc', Ri, Ri, 0, -180 * 60000), ('close',)]])


def tpl_wave(slide, x, y, w, h, pal=None, color=None):
    """波浪线：半圆弧交替上下（开放路径，只描边）。"""
    c = color or P(pal, 'main')
    r = h * 0.3
    n = max(2, int(w / (2 * r)))
    r = w / (2 * n)
    cy = y + h / 2
    cmds = [('move', 0, r)]
    for i in range(n):
        cmds.append(('arc', r, r, 180 * 60000, (180 if i % 2 == 0 else -180) * 60000))
    sp = shape(slide, MSO_SHAPE.RECTANGLE, x, cy - r, w, 2 * r, stroke=c, sw=2.0)
    custgeom(sp, w, 2 * r, [cmds])


def tpl_square_whirl(slide, x, y, w, h, pal=None, color=None, n=10):
    """方旋：嵌套正方形，逐层缩 0.87 转 7°（op-art 涡旋）。"""
    c = color or P(pal, 'main')
    s = min(w, h) * 0.9
    cx, cy = x + w / 2, y + h / 2
    for k in range(n):
        shape(slide, MSO_SHAPE.RECTANGLE, cx - s / 2, cy - s / 2, s, s,
              stroke=c, sw=1.2, rot=k * 7)
        s *= 0.87


# ================================================================ 涌现场（4）
def tpl_moire(slide, x, y, w, h, pal=None, color=None, angle=6):
    """莫尔条纹：两组平行线网格错角叠加，干涉涌现。"""
    c = color or P(pal, 'ink')
    g = min(w, h) / 16
    cx, cy = x + w / 2, y + h / 2
    L = math.hypot(w, h)
    n = int(L / g) + 2
    a = math.radians(angle)
    dx, dy = math.sin(a), math.cos(a)          # 线方向（近垂直）
    nx, ny = math.cos(a), -math.sin(a)         # 法向
    for k in range(-n, n + 1):
        for rot in (0, 1):
            if rot == 0:
                ddx, ddy, nnx, nny = 0.0, 1.0, 1.0, 0.0
            else:
                ddx, ddy, nnx, nny = dx, dy, nx, ny
            px, py = cx + k * g * nnx, cy + k * g * nny
            seg_pts = clip_seg(px - ddx * L, py - ddy * L, px + ddx * L, py + ddy * L, x, y, w, h)
            if seg_pts:
                seg(slide, *seg_pts, c, 1.0, alpha=0.55)


def tpl_phyllotaxis(slide, x, y, w, h, pal=None, color=None, n=160):
    """叶序点阵：r=c√i, θ=i·137.5°（向日葵种子排布）。"""
    c = color or P(pal, 'main')
    cx, cy = x + w / 2, y + h / 2
    cmax = (min(w, h) / 2 - 5) / math.sqrt(n)
    for i in range(1, n + 1):
        r = cmax * math.sqrt(i)
        a = math.radians(i * 137.508)
        d = 2.2 + 3.0 * i / n
        dot(slide, cx + r * math.cos(a), cy + r * math.sin(a), d, c, alpha=0.85)
    dot(slide, cx, cy, 5, P(pal, 'sub'))


def tpl_halftone(slide, x, y, w, h, pal=None, color=None, cols=18, rows=10):
    """半调渐变：点阵半径沿 x 线性增大（印刷网目效果）。"""
    c = color or P(pal, 'ink')
    cw, ch = w / cols, h / rows
    for i in range(cols):
        for j in range(rows):
            r = 0.7 + 3.0 * i / (cols - 1)
            dot(slide, x + (i + .5) * cw, y + (j + .5) * ch, 2 * r, c, alpha=0.8)


def tpl_tartan(slide, x, y, w, h, pal=None):
    """塔坦格纹：经纬色带 α 叠加 + 细强调线（苏格兰 plaid）。"""
    m, s, a = P(pal, 'main'), P(pal, 'sub'), P(pal, 'accent')
    unit = min(w, h) / 6
    pat = [1.4, 0.5, 1.4, 0.5, 2.2]
    tot = sum(pat)
    scx, scy = w / tot, h / tot
    posx = posy = 0
    for k, pw in enumerate(pat):
        col = m if k % 2 == 0 else s
        shape(slide, MSO_SHAPE.RECTANGLE, x + posx, y, pw * scx, h, fill=col, alpha=0.28)
        shape(slide, MSO_SHAPE.RECTANGLE, x, y + posy, w, pw * scy, fill=col, alpha=0.28)
        posx += pw * scx
        posy += pw * scy
    shape(slide, MSO_SHAPE.RECTANGLE, x + w * 0.45, y, 2, h, fill=a, alpha=0.6)
    shape(slide, MSO_SHAPE.RECTANGLE, x, y + h * 0.45, w, 2, fill=a, alpha=0.6)


# ================================================================ 东亚铺砌（4）
def tpl_kagome(slide, x, y, w, h, pal=None, color=None):
    """笼目：三角 lattice 上的六边形线稿（三六角形铺砌，竹编同源）。"""
    c = color or P(pal, 'main')
    s = min(w / 7, h / 5.6)
    for cx, cy in hex_centers(x, y, w, h, s, 'kagome'):
        pts = hex_pts(cx, cy, s)
        poly(slide, cx - s, cy - s, 2 * s, 2 * s,
             [(px - (cx - s), py - (cy - s)) for px, py in pts], stroke=c, sw=1.1)


def tpl_asanoha(slide, x, y, w, h, pal=None, color=None):
    """麻の葉：六边形 + 六辐条 + 六角星线（刺子绣经典）。"""
    c = color or P(pal, 'main')
    s = min(w / 5.5, h / 4)
    for cx, cy in hex_centers(x, y, w, h, s, 'honey'):
        pts = hex_pts(cx, cy, s)
        lp = [(px - (cx - s), py - (cy - s)) for px, py in pts]
        poly(slide, cx - s, cy - s, 2 * s, 2 * s, lp, stroke=c, sw=0.9)
        for px, py in pts:
            seg(slide, cx, cy, px, py, c, 0.9)
        for tri in ((0, 2, 4), (1, 3, 5)):
            tp = [pts[i] for i in tri]
            poly(slide, cx - s, cy - s, 2 * s, 2 * s,
                 [(px - (cx - s), py - (cy - s)) for px, py in tp], stroke=c, sw=0.9)


def tpl_kikko(slide, x, y, w, h, pal=None, color=None):
    """亀甲：六边形边共享铺砌（龟甲纹）。"""
    c = color or P(pal, 'main')
    s = min(w / 5.5, h / 4)
    for cx, cy in hex_centers(x, y, w, h, s, 'honey'):
        pts = hex_pts(cx, cy, s)
        poly(slide, cx - s, cy - s, 2 * s, 2 * s,
             [(px - (cx - s), py - (cy - s)) for px, py in pts],
             fill=c, alpha=0.08, stroke=c, sw=1.0)


def tpl_uroko(slide, x, y, w, h, pal=None):
    """鱗：等边三角上下相错、双色交替（鱼鳞/三角鳞）。"""
    m, a = P(pal, 'main'), P(pal, 'accent')
    s = min(w / 8, h / 3.6)
    hh = math.sqrt(3) / 2 * s
    rows = int(h / hh)
    cols = int(w / s) + 1
    x0, y0 = x + (w - cols * s) / 2, y + (h - rows * hh) / 2
    for row in range(rows):
        off = (row % 2) * s / 2
        for col in range(cols):
            bx, by = x0 + col * s + off, y0 + row * hh
            up = [(0, hh), (s, hh), (s / 2, 0)]
            dn = [(s / 2, 0), (s * 1.5, 0), (s, hh)]
            for pts, k in ((up, 0), (dn, 1)):
                if bx + s * 1.5 > x + w + 1:
                    continue
                poly(slide, bx, by, s * 1.5, hh, pts, fill=m if k == 0 else a, alpha=0.85)


# ================================================================ 西亚/欧洲铺砌（4）
def tpl_tumbling_blocks(slide, x, y, w, h, pal=None, color=None):
    """翻滚块：六边形 split 三菱形，三档透明度 → 3D 立方体错觉（rhombille）。"""
    c = color or P(pal, 'main')
    s = min(w / 5, h / 3.6)
    for cx, cy in hex_centers(x, y, w, h, s, 'honey'):
        v = hex_pts(cx, cy, s)
        rhombi = [([0, 1, 2], 0.22), ([2, 3, 4], 0.5), ([4, 5, 0], 0.92)]
        for idx, al in rhombi:
            pts = [(cx, cy)] + [v[i] for i in idx]
            poly(slide, cx - s, cy - s, 2 * s, 2 * s,
                 [(px - (cx - s), py - (cy - s)) for px, py in pts], fill=c, alpha=al)


def tpl_star_cross(slide, x, y, w, h, pal=None):
    """星与十字：八角星节点 + 十字填隙（伊斯兰 star-and-cross 简化）。"""
    m, a = P(pal, 'main'), P(pal, 'accent')
    s = min(w / 5, h / 3.2)
    nx, ny = max(2, int(w / (s * 1.5))), max(2, int(h / (s * 1.5)))
    dx, dy = w / nx, h / ny
    for i in range(nx):
        for j in range(ny):
            cx, cy = x + (i + .5) * dx, y + (j + .5) * dy
            q = s * 0.62
            shape(slide, MSO_SHAPE.RECTANGLE, cx - q / 2, cy - q / 2, q, q, fill=m, alpha=0.85)
            shape(slide, MSO_SHAPE.RECTANGLE, cx - q / 2, cy - q / 2, q, q, fill=m, alpha=0.85, rot=45)
            if i < nx - 1 and j < ny - 1:
                mx, my = cx + dx / 2, cy + dy / 2
                q2 = s * 0.3
                shape(slide, MSO_SHAPE.CROSS, mx - q2 / 2, my - q2 / 2, q2, q2, fill=a, alpha=0.8)


def tpl_ichimatsu(slide, x, y, w, h, pal=None):
    """市松：双色棋盘格（歌舞伎幕布纹）。"""
    m, s = P(pal, 'main'), P(pal, 'sub')
    n = 6
    cell = min(w, h) / n
    nx, ny = int(w / cell), int(h / cell)
    x0, y0 = x + (w - nx * cell) / 2, y + (h - ny * cell) / 2
    for i in range(nx):
        for j in range(ny):
            shape(slide, MSO_SHAPE.RECTANGLE, x0 + i * cell, y0 + j * cell, cell, cell,
                  fill=m if (i + j) % 2 == 0 else s, alpha=0.88)


def tpl_argyle(slide, x, y, w, h, pal=None):
    """菱格：45° 方格双色 + 对角细线（argyle 毛衣纹）。"""
    m, s, hc = P(pal, 'main'), P(pal, 'sub'), P(pal, 'hair')
    cell = min(w / 5, h / 3)
    a = cell * 0.72
    nx, ny = int(w / cell) + 1, int(h / cell) + 1
    x0, y0 = x + (w - (nx - 1) * cell) / 2, y + (h - (ny - 1) * cell) / 2
    for i in range(nx):
        for j in range(ny):
            cx, cy = x0 + i * cell + (j % 2) * cell / 2, y0 + j * cell
            if cx - a / 2 < x or cx + a / 2 > x + w or cy - a / 2 < y or cy + a / 2 > y + h:
                continue
            shape(slide, MSO_SHAPE.RECTANGLE, cx - a / 2, cy - a / 2, a, a,
                  fill=m if (i + j) % 2 == 0 else s, alpha=0.8, rot=45)
    g = cell / 2
    k = -int(h / g)
    while k * g < w + h:
        for (x1, y1, x2, y2) in ((x + k * g, y, x + k * g - h, y + h),
                                 (x + k * g, y, x + k * g + h, y + h)):
            cp = clip_seg(x1, y1, x2, y2, x, y, w, h)
            if cp:
                seg(slide, *cp, hc, 1.0, alpha=0.9)
        k += 1


# ================================================================ 边饰带（4）
def tpl_meander(slide, x, y, w, h, pal=None, color=None):
    """回纹/希腊钥匙：连续方折螺旋单元 + 基线（富贵不断头）。"""
    c = color or P(pal, 'main')
    u = h * 0.92
    n = max(2, int(w / u))
    u = w / n
    for i in range(n):
        x0 = x + i * u
        pts = [(0, h), (0, 0), (u, 0), (u, .7 * h), (.3 * u, .7 * h),
               (.3 * u, .3 * h), (.7 * u, .3 * h), (.7 * u, .85 * h)]
        sp = shape(slide, MSO_SHAPE.RECTANGLE, x0, y, u, h, stroke=c, sw=1.6)
        custgeom(sp, u, h, [[('move', pts[0][0], pts[0][1])] +
                            [('line', qx, qy) for qx, qy in pts[1:]]])
    seg(slide, x, y + h, x + w, y + h, c, 1.6)


def tpl_egg_dart(slide, x, y, w, h, pal=None):
    """卵箭纹：椭圆(卵)与倒三角(箭)交替，上下发丝线（古典建筑线脚）。"""
    m, s, hc = P(pal, 'main'), P(pal, 'sub'), P(pal, 'hair')
    u = h * 0.75
    n = max(3, int(w / u))
    u = w / n
    seg(slide, x, y + 2, x + w, y + 2, hc, 1.0)
    seg(slide, x, y + h - 2, x + w, y + h - 2, hc, 1.0)
    for i in range(n):
        cx = x + (i + .5) * u
        if i % 2 == 0:
            shape(slide, MSO_SHAPE.OVAL, cx - u * 0.26, y + h * 0.12, u * 0.52, h * 0.76,
                  fill=m, alpha=0.85)
        else:
            shape(slide, MSO_SHAPE.ISOSCELES_TRIANGLE, cx - u * 0.18, y + h * 0.2,
                  u * 0.36, h * 0.6, fill=s, alpha=0.85, rot=180)


def tpl_herringbone(slide, x, y, w, h, pal=None, color=None):
    """人字纹：±45° 砖交错（herringbone 铺地）。"""
    c = color or P(pal, 'main')
    L = min(w / 6, h / 3)
    t = L / 2
    step = L * math.sqrt(2) / 2
    nx, ny = int(w / step), int(h / step)
    x0, y0 = x + (w - nx * step) / 2 + step / 2, y + (h - ny * step) / 2 + step / 2
    for i in range(nx):
        for j in range(ny):
            cx, cy = x0 + i * step, y0 + j * step + (i % 2) * step / 2
            if cy + L / 2 > y + h:
                continue
            shape(slide, MSO_SHAPE.RECTANGLE, cx - L / 2, cy - t / 2, L, t,
                  fill=c, alpha=0.8 if (i + j) % 2 == 0 else 0.45,
                  rot=45 if (i + j) % 2 == 0 else -45)


def tpl_basket_weave(slide, x, y, w, h, pal=None):
    """ basket 编织：横竖双条交替成棋盘（竹编/篮纹）。"""
    m, s = P(pal, 'main'), P(pal, 'sub')
    cell = min(w / 6, h / 3)
    nx, ny = int(w / cell), int(h / cell)
    x0, y0 = x + (w - nx * cell) / 2, y + (h - ny * cell) / 2
    bw, gap = cell * 0.88, cell * 0.1
    for i in range(nx):
        for j in range(ny):
            bx, by = x0 + i * cell, y0 + j * cell
            c = m if (i + j) % 2 == 0 else s
            if (i + j) % 2 == 0:
                for k in range(2):
                    shape(slide, MSO_SHAPE.RECTANGLE, bx + (cell - bw) / 2,
                          by + gap / 2 + k * (bw / 2 + gap / 2), bw, bw / 2 - gap / 2,
                          fill=c, alpha=0.8)
            else:
                for k in range(2):
                    shape(slide, MSO_SHAPE.RECTANGLE,
                          bx + gap / 2 + k * (bw / 2 + gap / 2), by + (cell - bw) / 2,
                          bw / 2 - gap / 2, bw, fill=c, alpha=0.55)


# ================================================================ 放射与中式（4）
def tpl_sunburst(slide, x, y, w, h, pal=None, color=None, n=24):
    """旭日：细长三角放射排布，长短相间（Art Deco 日出母题）。"""
    c = color or P(pal, 'accent')
    R = min(w, h) / 2 * 0.95
    cx, cy = x + w / 2, y + h / 2
    for k in range(n):
        ang = k * 360 / n
        a = math.radians(ang)
        ln = R if k % 2 == 0 else R * 0.7
        tw = R * 0.085
        pcx, pcy = cx + ln / 2 * math.cos(a), cy + ln / 2 * math.sin(a)
        shape(slide, MSO_SHAPE.ISOSCELES_TRIANGLE, pcx - tw / 2, pcy - ln / 2, tw, ln,
              fill=c, alpha=0.9, rot=ang + 90)
    dot(slide, cx, cy, R * 0.16, P(pal, 'main'))


def tpl_fangsheng(slide, x, y, w, h, pal=None):
    """方胜：两菱形角部相叠成对，横向重复（中式吉祥纹）。"""
    m, s = P(pal, 'main'), P(pal, 'sub')
    q = min(h * 0.52, w / 7)
    n = max(1, int(w / (q * 2.6)))
    x0 = x + (w - n * q * 2.6) / 2 + q * 0.5
    cy = y + h / 2
    for i in range(n):
        cx = x0 + i * q * 2.6
        shape(slide, MSO_SHAPE.RECTANGLE, cx - q / 2, cy - q / 2 - q * 0.18, q, q,
              fill=m, alpha=0.85, rot=45)
        shape(slide, MSO_SHAPE.RECTANGLE, cx - q / 2 + q * 0.72, cy - q / 2 + q * 0.54, q, q,
              fill=s, alpha=0.85, rot=45)


def tpl_bead_reel(slide, x, y, w, h, pal=None, color=None):
    """串珠：圆珠与扁轮交替（古典 bead-and-reel 线脚）。"""
    c = color or P(pal, 'main')
    d = h * 0.5
    u = d * 1.5
    n = max(3, int(w / u))
    u = w / n
    cy = y + h / 2
    for i in range(n):
        cx = x + (i + .5) * u
        if i % 2 == 0:
            dot(slide, cx, cy, d, c, alpha=0.9)
        else:
            shape(slide, MSO_SHAPE.OVAL, cx - d * 0.55, cy - d * 0.26, d * 1.1, d * 0.52,
                  fill=c, alpha=0.6)


def tpl_qiulu(slide, x, y, w, h, pal=None, color=None):
    """球路纹：大圆 + 四正交小圆 + 四斜独立圆，四向连续（宋锦经典）。"""
    c = color or P(pal, 'main')
    R = min(w / 4.4, h / 2.6)
    nx, ny = max(1, int(w / (2 * R))), max(1, int(h / (2 * R)))
    x0, y0 = x + (w - nx * 2 * R) / 2 + R, y + (h - ny * 2 * R) / 2 + R
    for i in range(nx):
        for j in range(ny):
            cx, cy = x0 + i * 2 * R, y0 + j * 2 * R
            ring(slide, cx, cy, R, c, 1.1)
            for k in range(4):
                a = math.radians(k * 90)
                ring(slide, cx + R * math.cos(a), cy + R * math.sin(a), 0.42 * R, c, 1.1)
            for k in range(4):
                a = math.radians(45 + k * 90)
                ring(slide, cx + 1.5 * R * math.cos(a), cy + 1.5 * R * math.sin(a),
                     0.26 * R, c, 1.1)


# ================================================================ 注册表
TEMPLATES = [
    # id, 名称, 家族, 函数, 角色, 关键参数, 出处/依据
    ("vesica", "双弓形", "经典派生", tpl_vesica, "焦点", "d_ratio 0.8/1.0/1.2", "两圆交集；v2 P2"),
    ("rosette", "玫瑰窗", "经典派生", tpl_rosette, "焦点", "n 瓣数, d_ratio", "旋转对称；v2 P3"),
    ("seed-of-life", "生命之胚", "经典派生", tpl_seed_of_life, "焦点", "7 等圆", "神圣几何；v2 P3"),
    ("moon", "月牙", "经典派生", tpl_moon, "点缀", "preset moon", "圆−圆；v2 P4"),
    ("octagram", "八角星", "经典派生", tpl_octagram, "焦点", "方∪方45°", "伊斯兰几何；v2 P4"),
    ("octagon", "八边形涌现", "经典派生", tpl_octagon_emergence, "焦点", "α≈0.35", "方∩方45°；v2 P4"),
    ("hexagram", "六角星", "经典派生", tpl_hexagram, "焦点", "倒三角下移 h/3", "三角∪三角；v2 P4"),
    ("seigaiha", "青海波", "经典派生", tpl_seigaiha, "背景/边饰", "同心半圆扇", "日本纹样；v2 P5"),
    ("pinwheel", "风车", "经典派生", tpl_pinwheel, "焦点", "g≈0.17R 留隙", "四分圆共心；v2 P5"),
    ("scallop", "扇贝边", "经典派生", tpl_scallop, "分隔", "半圆相切", "v2 P5"),
    ("shippo", "七宝", "经典派生", tpl_shippo, "背景", "r=s/√2", "刺子绣；v2 P6"),
    ("truchet", "Truchet 瓦", "经典派生", tpl_truchet, "背景", "seed, 随机朝向", "Smith 1987；v2 P6"),
    ("reuleaux", "勒洛三角", "经典派生", tpl_reuleaux, "焦点", "3×60° 弧", "三圆交集"),
    ("laurel", "月桂花环", "圆派生", tpl_laurel, "焦点/徽章", "leaves, 顶留口", "双弓形叶沿弧"),
    ("cloud", "云簇", "圆派生", tpl_cloud, "焦点", "同色圆之并", "轮廓即云"),
    ("ring-chain", "连环", "圆派生", tpl_ring_chain, "边饰", "n, 间距 1.3r", "奥运环同源"),
    ("guilloche", "玑镂窗", "圆派生", tpl_guilloche, "背景/焦点", "n, e≈0.35r", "万花尺/钞票纹"),
    ("golden-spiral", "黄金螺旋", "弧与螺旋", tpl_golden_spiral, "焦点", "fib [1,1,2,3,5,8]", "斐波那契方+四分弧"),
    ("rainbow", "虹带", "弧与螺旋", tpl_rainbow, "焦点/徽章", "半圆环三色", "同心弧差"),
    ("wave", "波浪线", "弧与螺旋", tpl_wave, "分隔", "半圆交替弧", "开放路径描边"),
    ("square-whirl", "方旋", "弧与螺旋", tpl_square_whirl, "背景/焦点", "缩 .87 转 7°", "op-art 嵌套"),
    ("moire", "莫尔条纹", "涌现场", tpl_moire, "背景", "angle≈6°", "双网格干涉"),
    ("phyllotaxis", "叶序点阵", "涌现场", tpl_phyllotaxis, "背景/焦点", "137.5°, r=c√i", "向日葵种子"),
    ("halftone", "半调渐变", "涌现场", tpl_halftone, "背景/分隔", "cols, rows", "印刷网目"),
    ("tartan", "塔坦格纹", "涌现场", tpl_tartan, "背景", "α 经纬叠加", "苏格兰 plaid"),
    ("kagome", "笼目", "东亚铺砌", tpl_kagome, "背景", "a=3√3/2·s", "三六角形铺砌"),
    ("asanoha", "麻の葉", "东亚铺砌", tpl_asanoha, "背景", "六边+辐条+星", "刺子绣"),
    ("kikko", "亀甲", "东亚铺砌", tpl_kikko, "背景", "边共享六边", "龟甲纹"),
    ("uroko", "鱗", "东亚铺砌", tpl_uroko, "背景/边饰", "三角双色", "鱼鳞纹"),
    ("tumbling-blocks", "翻滚块", "西亚欧洲", tpl_tumbling_blocks, "背景/焦点", "三菱形 α .92/.5/.22", "rhombille 拼布"),
    ("star-cross", "星与十字", "西亚欧洲", tpl_star_cross, "背景", "八角星+十字", "伊斯兰铺砌"),
    ("ichimatsu", "市松", "西亚欧洲", tpl_ichimatsu, "背景/边饰", "棋盘格", "歌舞伎幕"),
    ("argyle", "菱格", "西亚欧洲", tpl_argyle, "背景", "45°方+对角线", "argyle 毛衣"),
    ("meander", "回纹", "边饰带", tpl_meander, "边饰/分隔", "方折螺旋+基线", "希腊钥匙/雷纹"),
    ("egg-dart", "卵箭纹", "边饰带", tpl_egg_dart, "边饰", "卵+箭交替", "古典线脚"),
    ("herringbone", "人字纹", "边饰带", tpl_herringbone, "背景/边饰", "±45° 砖", "铺地纹"),
    ("basket-weave", "篮纹", "边饰带", tpl_basket_weave, "背景", "横竖双条", "竹编"),
    ("sunburst", "旭日", "放射中式", tpl_sunburst, "焦点/角饰", "n, 长短相间", "Art Deco"),
    ("fangsheng", "方胜", "放射中式", tpl_fangsheng, "边饰/点缀", "菱角相叠", "中式吉祥"),
    ("bead-reel", "串珠", "放射中式", tpl_bead_reel, "边饰/分隔", "珠+轮交替", "古典线脚"),
    ("qiulu", "球路纹", "放射中式", tpl_qiulu, "背景", "大圆+8 小圆", "宋锦"),
]

if __name__ == "__main__":
    print("decor_templates: %d templates" % len(TEMPLATES))
    for tid, name, fam, fn, role, params, src in TEMPLATES:
        print("  %-16s %-8s %s" % (tid, fam, name))