#!/usr/bin/env python3
"""把 .pptx 逐页渲染成 PNG，供阶段六「结果校验」目视自查。

引擎（--engine）:
  officecli  默认优先。officecli 逐页截图，能显示原生公式。
             --render html    自带 HTML 引擎 + 无头浏览器，跨平台；会多画母版里未实例化的
                              日期/页码占位（属渲染假象，不是 pptx 的问题）。默认。
             --render native  本机 PowerPoint 渲染，最保真，仅 Windows 且已装 PowerPoint。
             --render auto    有 PowerPoint 用 native，否则 html。
  pillow     内置近似渲染（只依赖 power-pptx + Pillow），看不到公式与动画，
             只适合检查重叠、溢出、失衡。
  auto       默认。先试 officecli，不可用或失败再退回 pillow。

Usage:
  python render_preview.py deck.pptx [-o 输出目录] [--engine auto|officecli|pillow]
                           [--render html|native|auto] [--scale 1.0] [--sheet]

  -o      输出目录，默认 <pptx 同名>_preview/。预览图不交付，别混进交付物
  --scale 输出宽度倍率，1.0 对应 1600 px 宽（officecli 引擎）
  --sheet 额外生成 2x2 的 contact sheet（sheetN.png），一次看四页

Exit codes:
  0 成功（stdout 最后一行是输出目录；stderr 说明实际使用的引擎）
  1 指定引擎失败
  2 输入不存在
"""

from __future__ import annotations

import argparse
import io
import sys
from pathlib import Path

FONT_CANDIDATES = [
    ("C:/Windows/Fonts/msyh.ttc", "C:/Windows/Fonts/msyhbd.ttc"),
    ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
     "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"),
]


# ---------------------------------------------------------------------------
# officecli engine
# ---------------------------------------------------------------------------
def render_officecli(src: Path, out: Path, scale: float, render: str) -> list:
    import ocli
    n = ocli.slide_count(str(src))
    if render == "html":
        sys.stderr.write("[render_preview] html 渲染会多画日期/页码占位，属渲染假象\n")
    width = int(1600 * scale) if scale != 1.0 else None
    files = []
    for i in range(1, n + 1):
        p = out / f"s{i:02d}.png"
        for attempt in range(3):  # native 渲染偶发单页失败，重试同一页即可成功
            try:
                ocli.screenshot(str(src), str(p), start=i, end=i, render=render, width=width)
                break
            except ocli.OfficeCliError:
                if attempt == 2:
                    raise
        if not p.exists():
            raise RuntimeError(f"officecli 没有生成 {p.name}")
        files.append(p)
    return files


# ---------------------------------------------------------------------------
# Pillow engine
# ---------------------------------------------------------------------------
def render_pillow(src: Path, out: Path) -> list:
    from PIL import Image, ImageDraw, ImageFont
    from power_pptx import Presentation

    reg = bld = None
    for r, b in FONT_CANDIDATES:
        if Path(r).exists():
            reg, bld = r, b if Path(b).exists() else r
            break
    cache = {}

    def font(sz, bold):
        k = (round(sz), bool(bold))
        if k not in cache:
            px = max(6, int(round(sz * 100 / 72)))
            cache[k] = ImageFont.truetype(bld if bold else reg, px) if reg else ImageFont.load_default()
        return cache[k]

    sc = 100 / 914400  # px per EMU

    def fill_of(fill):
        try:
            if fill.type == 1:
                return tuple(fill.fore_color.rgb)
        except Exception:
            pass
        return None

    def draw_text(d, tf, x, y, w):
        cy = y
        for p in tf.paragraphs:
            if not p.runs:
                continue
            r0 = p.runs[0]
            sz = r0.font.size.pt if r0.font.size else 12
            f = font(sz, r0.font.bold)
            try:
                col = tuple(r0.font.color.rgb)
            except Exception:
                col = (38, 38, 38)
            text = "".join(r.text for r in p.runs)
            ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.2
            lh = sz * 100 / 72 * ls * 1.15
            lines, cur = [], ""
            for ch in text:
                if cur and d.textlength(cur + ch, font=f) > w:
                    lines.append(cur)
                    cur = ch
                else:
                    cur += ch
            lines.append(cur)
            for ln in lines:
                ax = x
                if p.alignment == 2:
                    ax = x + (w - d.textlength(ln, font=f)) / 2
                elif p.alignment == 3:
                    ax = x + w - d.textlength(ln, font=f)
                d.text((ax, cy), ln, font=f, fill=col)
                cy += lh

    def shape_kind(sh):
        # 原生公式形状没有预设几何，取 shape_type 会抛 NotImplementedError
        try:
            return sh.shape_type
        except NotImplementedError:
            return None

    prs = Presentation(str(src))
    W, H = int(prs.slide_width * sc), int(prs.slide_height * sc)
    files = []
    skipped = 0
    for i, slide in enumerate(prs.slides, 1):
        im = Image.new("RGB", (W, H), "white")
        d = ImageDraw.Draw(im)
        for sh in slide.shapes:
            kind = shape_kind(sh)
            if kind is None:
                skipped += 1
                continue
            x, y, w, h = [(v or 0) * sc for v in (sh.left, sh.top, sh.width, sh.height)]
            if kind == 13:  # picture
                pic = Image.open(io.BytesIO(sh.image.blob))
                if pic.mode in ("RGBA", "LA", "P"):
                    pic = pic.convert("RGBA")
                    bg = Image.new("RGBA", pic.size, (255, 255, 255, 255))
                    pic = Image.alpha_composite(bg, pic)
                pic = pic.convert("RGB").resize((max(1, int(w)), max(1, int(h))))
                im.paste(pic, (int(x), int(y)))
            elif getattr(sh, "has_table", False) and sh.has_table:
                t = sh.table
                cw, cy = w / len(t.columns), y
                for r in t.rows:
                    for ci, c in enumerate(r.cells):
                        d.rectangle([x + ci * cw, cy, x + (ci + 1) * cw, cy + r.height * sc],
                                    fill=fill_of(c.fill) or (255, 255, 255), outline=(217, 221, 227))
                        draw_text(d, c.text_frame, x + ci * cw + 8, cy + 4, cw - 16)
                    cy += r.height * sc
            elif kind == 9:  # connector / line
                try:
                    col = tuple(sh.line.color.rgb)
                except Exception:
                    col = (200, 200, 200)
                d.line([sh.begin_x * sc, sh.begin_y * sc, sh.end_x * sc, sh.end_y * sc],
                       fill=col, width=max(1, int((sh.line.width or 9525) * sc)))
            else:
                if kind == 1:  # autoshape
                    ln = None
                    try:
                        if sh.line.fill.type == 1:
                            ln = tuple(sh.line.color.rgb)
                    except Exception:
                        pass
                    box = [x, y, x + w, y + h]
                    if "OVAL" in str(getattr(sh, "auto_shape_type", "")):
                        d.ellipse(box, fill=fill_of(sh.fill), outline=ln)
                    else:
                        d.rectangle(box, fill=fill_of(sh.fill), outline=ln)
                if sh.has_text_frame and sh.text_frame.text.strip():
                    draw_text(d, sh.text_frame, x, y, w)
        p = out / f"s{i:02d}.png"
        im.save(p)
        files.append(p)
    if skipped:
        sys.stderr.write(f"[render_preview] 公式形状已跳过 {skipped} 个（Pillow 画不出原生公式）\n")
    return files


# ---------------------------------------------------------------------------
def make_sheets(files: list, out: Path) -> None:
    from PIL import Image
    for k in range(0, len(files), 4):
        ims = [Image.open(f).convert("RGB") for f in files[k:k + 4]]
        w, h = ims[0].size
        w2, h2 = min(w, 840), int(min(w, 840) * h / w)
        sheet = Image.new("RGB", (w2 * 2 + 10, h2 * 2 + 10), (120, 120, 120))
        for j, im in enumerate(ims):
            sheet.paste(im.resize((w2, h2)), ((j % 2) * (w2 + 10), (j // 2) * (h2 + 10)))
        sheet.save(out / f"sheet{k // 4 + 1}.png")


def main() -> int:
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("-o", "--out")
    ap.add_argument("--engine", choices=["auto", "officecli", "pillow"], default="auto")
    ap.add_argument("--render", choices=["html", "native", "auto"], default="html")
    ap.add_argument("--scale", type=float, default=1.0)
    ap.add_argument("--sheet", action="store_true")
    args = ap.parse_args()

    src = Path(args.pptx).expanduser().resolve()
    if not src.exists():
        sys.stderr.write(f"文件不存在: {src}\n")
        return 2
    out = Path(args.out).expanduser() if args.out else src.with_name(src.stem + "_preview")
    out.mkdir(parents=True, exist_ok=True)

    files, used = None, None
    if args.engine in ("auto", "officecli"):
        try:
            files, used = render_officecli(src, out, args.scale, args.render), f"officecli({args.render})"
        except Exception as e:
            msg = str(e).strip().splitlines()[0] if str(e).strip() else type(e).__name__
            sys.stderr.write(f"[render_preview] officecli 不可用（{msg}）\n")
            if args.engine == "officecli":
                return 1
    if files is None:
        sys.stderr.write("[render_preview] 使用 Pillow 近似渲染，看不到公式与动画\n")
        files, used = render_pillow(src, out), "pillow"
    if args.sheet:
        make_sheets(files, out)
    sys.stderr.write(f"[render_preview] 引擎: {used}，共 {len(files)} 页\n")
    print(str(out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
