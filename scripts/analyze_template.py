#!/usr/bin/env python3
"""Analyze a PPTX template for ppt-suggestion (README §1.1 模板分析).

Extracts: slide size, masters, layouts with placeholders, theme colors,
header/footer/page-number clues. Writes a JSON sidecar and a human-readable
template summary (Markdown) consumed by the planning phase.

Usage:
  python analyze_template.py <template.pptx> [-o 模板摘要.md]

Exit codes:
  0 success
  2 template unreadable / python-pptx missing
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.util import Emu
except ImportError:
    sys.stderr.write("python-pptx 未安装，请用 venv python 运行（见 README.md 快速上手）\n")
    raise SystemExit(2)

THEME_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

PH_TYPES = {
    0: "TITLE", 1: "BITMAP", 3: "CENTER_TITLE", 4: "SUBTITLE", 5: "BODY",
    6: "CHART", 7: "OBJECT", 8: "TABLE", 9: "HEADER", 12: "SLIDE_NUMBER",
    13: "TITLE", 15: "FOOTER", 16: "DATE", 18: "PICTURE", 25: "VERTICAL_BODY",
}


def read_theme_colors(prs) -> dict:
    colors = {}
    try:
        master = prs.slide_masters[0]
        theme_part = master.part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        from lxml import etree
        root = etree.fromstring(theme_part.blob)
        for tag in ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
                    "accent4", "accent5", "accent6", "tx1", "tx2"):
            node = root.find(f".//{{{THEME_NS}}}clrScheme/{{{THEME_NS}}}{tag}")
            if node is None:
                continue
            srgb = node.find(f"{{{THEME_NS}}}srgbClr")
            if srgb is not None:
                colors[tag] = srgb.get("val")
            else:
                sysc = node.find(f"{{{THEME_NS}}}sysClr")
                if sysc is not None and sysc.get("lastClr"):
                    colors[tag] = sysc.get("lastClr")
    except Exception:
        pass
    return colors


def analyze(prs) -> dict:
    w_in = Emu(prs.slide_width).inches if prs.slide_width else None
    h_in = Emu(prs.slide_height).inches if prs.slide_height else None
    layouts = []
    chrome = set()
    for mi, master in enumerate(prs.slide_masters):
        for lo in master.slide_layouts:
            phs = []
            for sh in lo.placeholders:
                try:
                    ptype = sh.placeholder_format.type
                    ptype_name = PH_TYPES.get(int(ptype), str(ptype))
                except Exception:
                    ptype_name = "UNKNOWN"
                phs.append({
                    "idx": int(sh.placeholder_format.idx) if sh.placeholder_format.idx is not None else None,
                    "type": ptype_name,
                    "pos": [round(Emu(sh.left).inches, 2) if sh.left is not None else None,
                            round(Emu(sh.top).inches, 2) if sh.top is not None else None],
                })
                if ptype_name in ("HEADER", "FOOTER", "SLIDE_NUMBER", "DATE"):
                    chrome.add(ptype_name)
            layouts.append({"name": lo.name or "", "placeholders": phs})
    return {
        "slide_size_inches": [round(w_in, 3) if w_in else None,
                              round(h_in, 3) if h_in else None],
        "masters": len(prs.slide_masters),
        "layouts": layouts,
        "theme_colors": read_theme_colors(prs),
        "chrome": sorted(chrome),
        "slide_count": len(prs.slides._sldIdLst),
    }


def to_markdown(info: dict) -> str:
    lines = ["# 模板摘要", ""]
    w, h = info["slide_size_inches"]
    ratio = "16:9" if w and h and abs(w / h - 16 / 9) < 0.05 else (
        "4:3" if w and h and abs(w / h - 4 / 3) < 0.05 else f"{w}×{h} in")
    lines.append(f"- 页面比例：{ratio}（{w} × {h} in）")
    lines.append(f"- 母版数量：{info['masters']}")
    if info["chrome"]:
        lines.append(f"- 页眉页脚页码：包含 {' / '.join(info['chrome'])} 占位符")
    else:
        lines.append("- 页眉页脚页码：未在版式中发现占位符（可能画在母版图形里，需人工确认）")
    tc = info["theme_colors"]
    if tc:
        lines.append("- 主题色：")
        for k in ("dk1", "dk2", "lt1", "lt2", "accent1", "accent2", "accent3", "accent4"):
            if k in tc:
                lines.append(f"  - {k}: #{tc[k]}")
    else:
        lines.append("- 主题色：读取失败（渲染器将回退默认白底低饱和色）")
    lines.append("")
    lines.append("## 版式清单")
    for lo in info["layouts"]:
        names = [f"{p['idx']}:{p['type']}" for p in lo["placeholders"]]
        lines.append(f"- **{lo['name'] or '(未命名)'}**" + (f" — {', '.join(names)}" if names else " — (无占位符)"))
    lines.append("")
    lines.append("## 待人工核对")
    lines.append("- [ ] 页眉、页脚、页码实际位置与样式")
    lines.append("- [ ] 分隔线、色块及装饰元素在哪些页、放哪里")
    lines.append("- [ ] 图片放置逻辑（位置/大小/留白规律）")
    lines.append("- [ ] 各类页面的版式规律（标题/目录/章节/正文/图文/表格/总结）")
    return "\n".join(lines)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("template")
    ap.add_argument("-o", "--output")
    args = ap.parse_args()
    tpl = Path(args.template).expanduser()
    try:
        prs = Presentation(str(tpl))
    except Exception as e:
        sys.stderr.write(f"打开模板失败: {e}\n")
        return 2
    info = analyze(prs)
    if args.output:
        out = Path(args.output).expanduser()
        if not out.is_absolute():
            out = Path.cwd() / out
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(to_markdown(info), encoding="utf-8")
        sidecar = out.with_suffix(".template.json")
        sidecar.write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
        print(str(out))
    else:
        print(to_markdown(info))
    return 0


if __name__ == "__main__":
    sys.exit(main())
