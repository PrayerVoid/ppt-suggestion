#!/usr/bin/env python3
"""配色 / 基本版式 / 内容版式随机采样器。

给 AI 提供带随机性的候选，避免千篇一律；最终取舍与覆盖仍由 AI 决定。
文档规则分两类：硬约束（采样前后强制过滤）、软偏好（转成下方权重倍数）。

用法:
  sample_style.py --seed 7 --count 3 --scene 学术 --pages 14
  sample_style.py --pages cover,catalog,section,content,close --force-color C0392B
  sample_style.py --seed 7 --emit-constants [候选序号]   # 输出可粘进渲染脚本副本的常量块
"""
import argparse
import colorsys
import json
import random
import sys

# ---------------- 软偏好：权重倍数（倍数越小概率越低） ----------------
SAT_HIGH, SAT_HIGH_W = 0.6, 0.05    # design.md:26 禁鲜艳颜色：高饱和强调色概率大幅下降
SAT_LOW, SAT_LOW_W = 0.18, 0.3      # design.md:16 低饱和但要可辨：近灰色略降
PRIMARY_DEG, PRIMARY_W = 15, 0.2    # design.md:16 避免纯色：近纯红/绿/蓝（且 S>0.8）再降
L_MID, L_SIGMA = 0.38, 0.10         # design.md:16 低亮度：强调色亮度集中在中低段
PURE_BLACK_W = 0.1                  # design.md:16 避免纯色：正文近纯黑概率下降
SCENE_W = 1.35                      # 场景色相仅轻微加权；低饱和不等于偏绿
ASYM_W, CARDS_W = 1.3, 0.4          # design.md:68 非对称/结构版式加权；design.md:27 卡片降权

SCENE_HUE = {"学术": (200, 260), "工程": (170, 230), "商务": (190, 250),
             "政务": (345, 20), "教育": (90, 170)}

# ---------------- 硬约束 ----------------
MIN_TEXT = 4.5     # validation.md:72 小字对比度 ≥ 4.5:1
MIN_FAINT = 2.5    # 仅页脚/图注使用的次要元数据，放宽
CHECKS = [("TEXT", "WHITE", MIN_TEXT), ("TEXT", "PANEL", MIN_TEXT), ("TEXT", "ACCENT_LT", MIN_TEXT),
          ("SUBTLE", "WHITE", MIN_TEXT), ("SUBTLE", "PANEL", MIN_TEXT),
          ("ACCENT", "WHITE", MIN_TEXT), ("ACCENT_DK", "WHITE", MIN_TEXT),
          ("ACCENT_DK", "ACCENT_LT", MIN_TEXT), ("FAINT", "WHITE", MIN_FAINT)]

# 版式分类：非对称/结构版式（design.md:68、story.md:23）、结构页（deck-schema.md:341）、不计种类的辅助块
ASYM = {"split", "layers", "rows", "stats", "compare", "pipeline", "hub", "timeline"}
STRUCT = {"layers", "steps", "rows", "pipeline", "hub", "timeline", "compare"}
AUX = {"lead", "insight", "chips"}
GRID = {"cards", "columns"}
CARDS_ROWS = {4: 2, 6: 3, 8: 4, 9: 3}   # deck-schema.md:221 条目数 → columns
ITEMS = {"stats": (3, 4), "columns": (2, 3), "steps": (3, 5), "chips": (6, 10), "pipeline": (3, 5),
         "hub": (3, 6), "timeline": (3, 5), "points": (3, 5), "compare": (3, 5), "callouts": (2, 2)}
RECIPES = ["stats", "stats+insight", "lead+stats", "layers", "layers+insight", "pipeline",
           "pipeline+insight", "hub", "timeline", "timeline+insight", "compare", "compare+insight",
           "rows", "steps+callouts", "steps+chips", "split", "split+insight", "table",
           "table+insight", "lead+columns", "lead+points", "points", "cards"]
MAX_INSIGHT, MIN_INSIGHT, MAX_CARDS, MAX_BLOCKS = 6, 3, 2, 3


# ---------------- 颜色工具 ----------------
def hls_hex(h, l, s):
    r, g, b = colorsys.hls_to_rgb((h % 360) / 360, l, s)
    return "%02X%02X%02X" % tuple(round(c * 255) for c in (r, g, b))


def hex_hls(hx):
    r, g, b = (int(hx[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    return h * 360, l, s


def rel_lum(hx):
    def f(c):
        c /= 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (int(hx[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def contrast(a, b):
    hi, lo = sorted((rel_lum(a), rel_lum(b)), reverse=True)
    return (hi + 0.05) / (lo + 0.05)


def gray(v):
    return "%02X" % v * 3


def in_hue(h, rng_):
    a, b = rng_
    return a <= h <= b if a <= b else (h >= a or h <= b)


def palette_violations(p, skip_accent_white=False):
    return [f"{a}/{b}<{m}" for a, b, m in CHECKS
            if not (skip_accent_white and (a, b) == ("ACCENT", "WHITE")) and contrast(p[a], p[b]) < m]


# ---------------- 配色采样 ----------------
def accent_weight(h, s, l):
    w = 1.0
    if s > SAT_HIGH:
        w *= SAT_HIGH_W
    if s < SAT_LOW:
        w *= SAT_LOW_W
    if s > 0.8 and min(abs((h - p + 180) % 360 - 180) for p in (0, 120, 240)) < PRIMARY_DEG:
        w *= PRIMARY_W
    return w * max(0.05, pow(2.718281828, -0.5 * ((l - L_MID) / L_SIGMA) ** 2))


def pick_hue(rng, scene):
    # 先确定色相，再在该色相内找满足对比度的 S/L，避免过滤通过率改变色相分布。
    while True:
        h = rng.uniform(0, 360)
        if scene not in SCENE_HUE:
            return h
        weight = SCENE_W if in_hue(h, SCENE_HUE[scene]) else 1.0
        if rng.random() < weight / SCENE_W:
            return h


def pick_accent(rng, scene):
    h = pick_hue(rng, scene)
    pool = []
    while not pool:
        for _ in range(80):
            s, l = rng.uniform(0.10, 0.95), rng.uniform(0.18, 0.50)
            hx = hls_hex(h, l, s)
            if contrast(hx, "FFFFFF") >= MIN_TEXT:
                pool.append((accent_weight(h, s, l), hx))
    return rng.choices([p[1] for p in pool], weights=[p[0] for p in pool])[0]


def derive_dark(h, s, l, lt):
    l = min(l - 0.10, 0.30)
    while l > 0.08 and (contrast(hls_hex(h, l, s), lt) < MIN_TEXT or contrast(hls_hex(h, l, s), "FFFFFF") < MIN_TEXT):
        l -= 0.02
    return hls_hex(h, max(l, 0.05), s)


def make_palette(rng, scene, forced=None):
    for _ in range(200):
        acc = forced or pick_accent(rng, scene)
        h, l, s = hex_hls(acc)
        lt = hls_hex(h, rng.uniform(0.93, 0.95), rng.uniform(0.15, 0.30))
        tv = rng.choices(range(0x00, 0x34), weights=[PURE_BLACK_W if v < 0x10 else 1 for v in range(0x00, 0x34)])[0]
        p = {"WHITE": "FFFFFF", "TEXT": gray(tv), "SUBTLE": gray(rng.randint(0x58, 0x6C)),
             "FAINT": gray(rng.randint(0x8E, 0x9E)), "ACCENT": acc,
             "ACCENT_DK": derive_dark(h, s, l, lt), "ACCENT_LT": lt,
             "LINE": hls_hex(h, rng.uniform(0.85, 0.89), rng.uniform(0.10, 0.20)),
             "PANEL": hls_hex(h, rng.uniform(0.955, 0.97), rng.uniform(0.12, 0.22))}
        if not palette_violations(p, skip_accent_white=bool(forced)):
            return p
    raise RuntimeError("采样不到满足对比度的配色")


def make_base(rng):
    m = round(rng.uniform(0.8, 1.1) * 20) / 20
    return {"MARGIN_L": m, "MARGIN_R": m, "BLOCK_GAP": round(rng.uniform(0.18, 0.30), 2),
            "TITLE_SIZE": rng.randint(24, 28)}


# ---------------- 内容版式采样 ----------------
def page_types(spec):
    if "," in spec or not spec.isdigit():
        types = spec.split(",")
        bad = [t for t in types if t not in ("cover", "catalog", "section", "content", "close")]
        if bad:
            raise SystemExit(f"未知页面类型: {bad}")
        return types
    n = int(spec)
    if n < 3:
        raise SystemExit("页数至少 3")
    body = n - 3
    if body < 6:   # 页数太少时不做目录与章节
        return ["cover"] + ["content"] * (n - 2) + ["close"]
    c = max(3, min(8, round(body / 4.5), body // 2))
    per, extra = divmod(body - c, c)
    seq = ["cover", "catalog"]
    for i in range(c):
        seq += ["section"] + ["content"] * (per + (1 if i < extra else 0))
    return seq + ["close"]


def primary(kinds):
    return next(k for k in kinds if k != "lead")


def block_spec(rng, kind):
    if kind == "cards":
        n = rng.choice(list(CARDS_ROWS))
        return {"kind": kind, "items": n, "columns": CARDS_ROWS[n]}
    if kind in ITEMS:
        return {"kind": kind, "items": rng.randint(*ITEMS[kind])}
    return {"kind": kind}


def sample_content(rng, n, allow_figure):
    recipes = RECIPES + (["figure"] if allow_figure else [])
    best, best_miss = None, None
    for _ in range(300):
        seq, prev, cards = [], None, 0
        insight = 0
        for _i in range(n):
            ok, ws = [], []
            for r in recipes:
                kinds = r.split("+")
                pr = primary(kinds)
                if pr == prev or (pr in GRID and prev in GRID):
                    continue
                if pr == "cards" and cards >= MAX_CARDS:
                    continue
                if "insight" in kinds and insight >= MAX_INSIGHT:
                    continue
                w = (ASYM_W if pr in ASYM else 1.0) * (CARDS_W if pr == "cards" else 1.0)
                ok.append(kinds)
                ws.append(w)
            kinds = rng.choices(ok, weights=ws)[0]
            seq.append(list(kinds))
            prev = primary(kinds)
            cards += prev == "cards"
            insight += "insight" in kinds
        # 补足全稿 insight 下限
        short = [i for i, k in enumerate(seq) if "insight" not in k and len(k) < MAX_BLOCKS and primary(k) not in GRID]
        rng.shuffle(short)
        total = sum("insight" in k for k in seq)
        for i in short[:max(0, min(MIN_INSIGHT, n) - total)]:
            seq[i].append("insight")
        miss = unmet(seq, n)
        if best is None or len(miss) < len(best_miss):
            best, best_miss = seq, miss
        if not miss:
            break
    return best, best_miss


def unmet(seq, n):
    prim = [primary(k) for k in seq]
    kinds = {x for k in seq for x in k} - AUX
    miss = []
    if sum(p in ASYM for p in prim) < 0.4 * n:
        miss.append("非对称/结构版式<40%")
    if sum(any(x in STRUCT for x in k) for k in seq) < min(2, n):
        miss.append("结构页<2")
    if len(kinds) < min(5, n):
        miss.append("版式种类<5")
    if sum("insight" in k for k in seq) < min(MIN_INSIGHT, n):
        miss.append("insight<3")
    return miss


def make_layout(rng, types, allow_figure):
    n = types.count("content")
    seq, miss = sample_content(rng, n, allow_figure) if n else ([], [])
    chapters = types.count("section")
    it, out = iter(seq), []
    for i, t in enumerate(types, 1):
        page = {"page": i, "type": t, "blocks": []}
        if t == "catalog":
            page["agenda_items"] = chapters
        if t == "content":
            kinds = next(it)
            page["blocks"] = [block_spec(rng, k) for k in kinds]
            page["structural"] = any(k in STRUCT for k in kinds)
            page["asymmetric"] = primary(kinds) in ASYM
        out.append(page)
    return out, miss


# ---------------- 候选与输出 ----------------
def make_candidate(seed, idx, scene, forced, types, allow_figure):
    rng = random.Random(f"{seed}-{idx}")
    palette = make_palette(rng, scene, forced)
    warn = [] if not forced or contrast(forced, "FFFFFF") >= MIN_TEXT else ["指定色与白底对比度<4.5，仅适合大字/色块"]
    layout, miss = make_layout(rng, types, allow_figure)
    n = types.count("content")
    summary = {"content_pages": n, "unmet_soft_targets": miss, "warnings": warn,
               "asymmetric_ratio": round(sum(p.get("asymmetric", False) for p in layout) / n, 2) if n else None,
               "insight_count": sum(b["kind"] == "insight" for p in layout for b in p["blocks"])}
    return {"id": idx, "palette": palette, "base": make_base(rng), "layout": layout, "summary": summary}


def emit_constants(c):
    pal = "\n".join(f"{k} = RGBColor(0x{v[0:2]}, 0x{v[2:4]}, 0x{v[4:6]})" for k, v in c["palette"].items())
    b = c["base"]
    return (f"{pal}\nMARGIN_L, MARGIN_R = {b['MARGIN_L']}, {b['MARGIN_R']}\n"
            f"BLOCK_GAP = {b['BLOCK_GAP']}\nTITLE_SIZE = {b['TITLE_SIZE']}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--seed", default=None, help="随机种子，缺省自动生成并写入输出")
    ap.add_argument("--count", type=int, default=3, help="候选数量")
    ap.add_argument("--scene", default="", help="学术/工程/商务/政务/教育，仅影响色相倾向")
    ap.add_argument("--force-color", default=None, help="用户指定强调色（十六进制），直接采用")
    ap.add_argument("--pages", default="12", help="页数，或逗号分隔的页面类型序列")
    ap.add_argument("--figure", action="store_true", help="允许采样 figure 版式（有可用原图时）")
    ap.add_argument("--emit-constants", nargs="?", const=0, type=int, metavar="序号",
                    help="只输出指定候选的渲染脚本常量块")
    a = ap.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    forced = None
    if a.force_color:
        forced = a.force_color.lstrip("#").upper()
        if len(forced) != 6 or any(ch not in "0123456789ABCDEF" for ch in forced):
            raise SystemExit("--force-color 需要 6 位十六进制")
    seed = a.seed if a.seed is not None else str(random.SystemRandom().randrange(10 ** 6))
    types = page_types(a.pages)
    cands = [make_candidate(seed, i, a.scene, forced, types, a.figure) for i in range(max(1, a.count))]
    if a.emit_constants is not None:
        print(emit_constants(cands[a.emit_constants]))
        return
    print(json.dumps({"seed": seed, "scene": a.scene, "force_color": forced, "page_types": types,
                      "candidates": cands}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
