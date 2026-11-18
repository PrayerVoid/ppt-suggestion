#!/usr/bin/env python3
"""Analyze a PPTX template for ppt-suggestion (references/source-intake.md).

Extracts: slide size, masters, layouts with placeholders, theme colors,
header/footer/page-number clues. Writes a JSON sidecar and a human-readable
template summary (Markdown) consumed by the planning phase.

Usage:
  python analyze_template.py <template.pptx> [-o 模板摘要.md]
  python analyze_template.py <template.pptx> --slides-json   # 逐页清单，供模板模式选页写 map

Exit codes:
  0 success
  2 template unreadable / power-pptx missing
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

try:
    from power_pptx import Presentation
    from power_pptx.util import Emu
except ImportError:
    sys.stderr.write("power-pptx 未安装，请先 pip install -r requirements.txt（见 README.md）\n")
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
    from animation_inventory import inventory
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
        "animation": inventory(prs),
    }


def _shape_kind(sh) -> str:
    if sh.shape_type is not None and "PICTURE" in str(sh.shape_type):
        return "picture"
    if getattr(sh, "has_table", False) and sh.has_table:
        return "table"
    if getattr(sh, "has_chart", False) and sh.has_chart:
        return "chart"
    if sh.has_text_frame:
        return "text"
    return "other"


# officecli 路径里每类元素各自从 1 计数（shape[K] 只数 p:sp，connector[K] 只数 p:cxnSp，依此类推）
_CATEGORY_BY_TAG = {"sp": "shape", "pic": "picture", "cxnSp": "connector", "grpSp": "group"}


def _category(sh) -> str:
    tag = sh._element.tag.rsplit("}", 1)[-1]
    if tag == "graphicFrame":
        return "table" if getattr(sh, "has_table", False) and sh.has_table else "chart"
    return _CATEGORY_BY_TAG.get(tag, tag)


def walk_shapes(shapes, parent_key: str = "", steps: tuple = (), depth: int = 0):
    """深度优先遍历形状（含组合内子形状），逐个产出记录。

    key：形状名；同一层有同名形状时写成 `名称#序号`（序号按 z 序从 1 起）；组合内的子形状写成
    `组合key/子key`。steps：officecli 路径的逐级 (类别, 序号)，复制页后顺序与名称不变而 id 会重编，
    所以按位置寻址才稳定。
    """
    names = [sh.name for sh in shapes]
    seen, per_cat = {}, {}
    for sh, name in zip(shapes, names):
        seen[name] = seen.get(name, 0) + 1
        label = f"{name}#{seen[name]}" if names.count(name) > 1 else name
        key = f"{parent_key}/{label}" if parent_key else label
        cat = _category(sh)
        per_cat[cat] = per_cat.get(cat, 0) + 1
        here = steps + ((cat, per_cat[cat]),)
        kind = "group" if cat == "group" else _shape_kind(sh)
        yield {"key": key, "name": name, "id": sh.shape_id, "category": cat, "kind": kind,
               "steps": here, "depth": depth, "shape": sh}
        if cat == "group":
            yield from walk_shapes(list(sh.shapes), key, here, depth + 1)


def slides_inventory(prs) -> list:
    """逐页清单：页号、版式名、每个形状（含组合内子形状）的 key/name/id/类型/占位符/位置(EMU)/文字摘要。

    组合子形状的位置是组合内坐标，不是页面坐标。
    """
    from animation_inventory import inventory
    animations = inventory(prs)["slides"]
    out = []
    for i, s in enumerate(prs.slides, 1):
        shapes = []
        for rec in walk_shapes(list(s.shapes)):
            sh = rec["shape"]
            ph = None
            if sh.is_placeholder:
                try:
                    ph = PH_TYPES.get(int(sh.placeholder_format.type), str(sh.placeholder_format.type))
                except Exception:
                    ph = "UNKNOWN"
            text = sh.text_frame.text.strip().replace("\n", " / ")[:40] if sh.has_text_frame else ""
            shapes.append({
                "key": rec["key"], "depth": rec["depth"], "name": sh.name, "id": sh.shape_id,
                "kind": rec["kind"], "placeholder": ph,
                "left": sh.left, "top": sh.top, "width": sh.width, "height": sh.height,
                "text": text,
            })
        out.append({"slide": i, "layout": s.slide_layout.name or "", "shapes": shapes,
                    "animation": animations[i - 1]})
    return out


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
        lines.append("- 主题色：读取失败；模板模式仍保留原主题关联，须依据实际预览人工解析显式字体配色，不自动改用自主白底方案")
    lines.append("")
    lines.append("## 版式清单")
    for lo in info["layouts"]:
        names = [f"{p['idx']}:{p['type']}" for p in lo["placeholders"]]
        lines.append(f"- **{lo['name'] or '(未命名)'}**" + (f" — {', '.join(names)}" if names else " — (无占位符)"))
    lines.append("")
    animation = info["animation"]
    lines.extend(["## 动画与转场预检", "",
                  "扫描实际页、版式和母版的全部 MC 分支；以下为 XML 清单，未验证放映播放。", "",
                  "| 输出源页 | 对象时序 | 构建列表 | 继承时序 | 转场 | 自动翻页 | 声音/Morph |",
                  "| --- | --- | --- | --- | --- | --- | --- |"])
    for page in animation["slides"]:
        timing, transition = page["timing"], page["transition"]
        advance = f"{transition['advance_after_ms']}ms" if transition["automatic"] else "否"
        lines.append(f"| {page['slide']} | {'有' if timing['present'] else '无'} | "
                     f"{'有' if timing['has_builds'] else '无'} | "
                     f"{'有' if page['has_inherited_timing'] else '无'} | {transition['type']} | "
                     f"{advance} | {transition['has_sound']}/{transition['is_morph']} |")
    lines.extend(["", "母版/版式的目标 ID 只在所属部件内解析，不能借用页面对象。详情见同名 .template.json。",
                  "生成前必须按编辑动作形成逐页动画计划，见 references/template-animation.md。", ""])
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
    ap.add_argument("--slides-json", action="store_true", help="输出逐页清单 JSON 到 stdout")
    args = ap.parse_args()
    tpl = Path(args.template).expanduser()
    try:
        prs = Presentation(str(tpl))
    except Exception as e:
        sys.stderr.write(f"打开模板失败: {e}\n")
        return 2
    if args.slides_json:
        print(json.dumps(slides_inventory(prs), ensure_ascii=False, indent=2))
        return 0
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
