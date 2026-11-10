#!/usr/bin/env python3
"""把 .pptx 逐页渲染成 PNG，供阶段四「模型自查」目视检查。

引擎（--engine）:
  aspose  Aspose.Slides 渲染，接近 PowerPoint 效果。依赖 aspose-slides，
          Linux/macOS 还需要系统装 libgdiplus（macOS: brew install mono-libgdiplus）。
          缺 libgdiplus 时会报 "type initializer for 'Gdip' threw an exception"。
  pillow  内置近似渲染，只依赖 python-pptx + Pillow，任何机器都能跑。
          按形状位置画矩形/圆/文字/图片/表格/线条，字体与换行与 PowerPoint 有出入，
          适合检查重叠、溢出、失衡，不适合检查字体细节。
  auto    默认。先试 aspose，失败自动退回 pillow。

Usage:
  python render_preview.py deck.pptx [-o 输出目录] [--engine auto|aspose|pillow]
                           [--scale 1.0] [--sheet]

  -o      输出目录，默认 <pptx 同名>_preview/。建议用系统临时目录，别混进交付物
  --scale aspose 引擎的缩放；pillow 引擎固定 100 px/in（13.33in 宽 = 1333 px）
  --sheet 额外生成 2x2 的 contact sheet（sheetN.png），一次看四页

说明:
  - Aspose 免费评估模式会在图片上叠加 Evaluation 水印，预览图只用于自查，不交付。
  - 安装 aspose-slides（国内用清华源）:
      uv pip install --python <venv-python> -i https://pypi.tuna.tsinghua.edu.cn/simple aspose-slides

Exit codes:
  0 成功（stdout 最后一行是输出目录；stderr 说明实际使用的引擎）
  2 输入不存在
"""

from __future__ import annotations

import argparse
import io
import sys
from pathlib import Path

REG = "/System/Library/Fonts/STHeiti Light.ttc"
BLD = "/System/Library/Fonts/STHeiti Medium.ttc"
FONT_CANDIDATES = [
    ("/System/Library/Fonts/STHeiti Light.ttc", "/System/Library/Fonts/STHeiti Medium.ttc"),
    ("/System/Library/Fonts/Hiragino Sans GB.ttc", "/System/Library/Fonts/Hiragino Sans GB.ttc"),
    ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
     "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"),
    ("C:/Windows/Fonts/msyh.ttc", "C:/Windows/Fonts/msyhbd.ttc"),
]


# ---------------------------------------------------------------------------
# Aspose engine
# ---------------------------------------------------------------------------
def render_aspose(src: Path, out: Path, scale: float) -> list:
    import aspose.slides as slides
    files = []
    with slides.Presentation(str(src)) as prs:
        for i, sld in enumerate(prs.slides, 1):
            p = out / f"s{i:02d}.png"
            with sld.get_image(scale, scale) as img:
                img.save(str(p), slides.ImageFormat.PNG)
            files.append(p)
    return files


# ---------------------------------------------------------------------------
# Pillow engine
# ---------------------------------------------------------------------------
def render_pillow(src: Path, out: Path) -> list:
    from PIL import Image, ImageDraw, ImageFont
    from pptx import Presentation

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

    prs = Presentation(str(src))
    W, H = int(prs.slide_width * sc), int(prs.slide_height * sc)
    files = []
    for i, slide in enumerate(prs.slides, 1):
        im = Image.new("RGB", (W, H), "white")
        d = ImageDraw.Draw(im)
        for sh in slide.shapes:
            x, y, w, h = [(v or 0) * sc for v in (sh.left, sh.top, sh.width, sh.height)]
            if sh.shape_type == 13:  # picture
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
            elif sh.shape_type == 9:  # connector / line
                try:
                    col = tuple(sh.line.color.rgb)
                except Exception:
                    col = (200, 200, 200)
                d.line([sh.begin_x * sc, sh.begin_y * sc, sh.end_x * sc, sh.end_y * sc],
                       fill=col, width=max(1, int((sh.line.width or 9525) * sc)))
            else:
                if sh.shape_type == 1:  # autoshape
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


def _ensure_gdiplus_env() -> None:
    """macOS：若装了 setup_libgdiplus_macos.sh 编译的 libgdiplus，带上 DYLD_LIBRARY_PATH 重启自己。
    dyld 只在进程启动时读这个变量，所以必须 exec，不能在进程内 os.environ 赋值。"""
    import os
    if sys.platform != "darwin":
        return
    lib = Path(os.environ.get("GDIP_HOME", Path.home() / ".comate-engine" / "gdip")) / "lib"
    if not (lib / "libgdiplus.dylib").exists():
        return
    cur = os.environ.get("DYLD_LIBRARY_PATH", "")
    if str(lib) in cur.split(":"):
        return
    env = dict(os.environ, DYLD_LIBRARY_PATH=f"{lib}:{cur}" if cur else str(lib))
    os.execve(sys.executable, [sys.executable] + sys.argv, env)


def main() -> int:
    _ensure_gdiplus_env()
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("-o", "--out")
    ap.add_argument("--engine", choices=["auto", "aspose", "pillow"], default="auto")
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
    if args.engine in ("auto", "aspose"):
        try:
            files, used = render_aspose(src, out, args.scale), "aspose"
        except Exception as e:  # ImportError / .NET Gdip 异常等
            msg = str(e).splitlines()[0] if str(e) else type(e).__name__
            sys.stderr.write(f"[render_preview] aspose 不可用（{msg}）\n")
            if args.engine == "aspose":
                return 1
    if files is None:
        files, used = render_pillow(src, out), "pillow"
    if args.sheet:
        make_sheets(files, out)
    sys.stderr.write(f"[render_preview] 引擎: {used}，共 {len(files)} 页\n")
    print(str(out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
