#!/usr/bin/env python3
"""给成品 pptx 加页间转场与内容页入场动画。

用法:
  python apply_animation.py in.pptx --deck deck.json --out out.pptx \
      [--transition fade|push|wipe|cover|cut|morph|none] [--entrance fade|wipe|appear|none] [--com-check]

入场动画按 PowerPoint 自己写出的标准结构生成（mainSeq 主序列、每个 block 一次点击、
块内其余形状 with-previous），不使用 power_pptx 的 animations（它不写 mainSeq，
PowerPoint 读回主序列为空，播放时部分形状不出现）。必须在 formula.py 之后运行，
否则公式占位框被替换后动画目标 id 失效。
--com-check 用本机 PowerPoint（需 pywin32）读回每页效果数，仅 Windows。

退出码: 0 成功并通过自检；1 自检失败；2 参数/文件问题。
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from lxml import etree
from power_pptx import Presentation
from power_pptx.enum.presentation import MSO_TRANSITION
from power_pptx.oxml.ns import qn

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
TRANSITIONS = {
    "fade": MSO_TRANSITION.FADE,
    "push": MSO_TRANSITION.PUSH,
    "wipe": MSO_TRANSITION.WIPE,
    "cover": MSO_TRANSITION.COVER,
    "cut": MSO_TRANSITION.CUT,
    "morph": MSO_TRANSITION.MORPH,
}
# presetID, presetSubtype, animEffect 的 filter（None 表示只有 set）；取自 PowerPoint 写出的文件
PRESETS = {
    "fade": (10, 0, "fade"),
    "wipe": (8, 2, "wipe(dir=left)"),
    "appear": (1, 0, None),
}
ANIMATED_TYPES = {"content", "catalog"}
CONTENT_TOP_FRAC = 1.39 / 7.5      # 标题区与分隔线之下
CONTENT_BOTTOM_FRAC = 7.05 / 7.5   # 页脚之上
BAND_GAP_IN = 0.05                 # 小于该垂直间隙的形状视为同一行


def content_shapes(slide, slide_h):
    lo, hi = slide_h * CONTENT_TOP_FRAC, slide_h * CONTENT_BOTTOM_FRAC
    out = []
    for sh in slide.shapes:
        if (sh.name or "").startswith("DECO_"):
            continue
        if sh.top is None or sh.height is None or not lo <= sh.top < hi:
            continue
        out.append(sh)
    return out


def split_bands(shapes, n_blocks):
    """按垂直范围把形状聚成 block 行；行数多于 block 数时合并间隙最小的相邻行。"""
    if not shapes:
        return []
    shapes = sorted(shapes, key=lambda s: (s.top, s.left))
    thr = int(BAND_GAP_IN * 914400)
    bands = [[shapes[0]]]
    ends = [shapes[0].top + shapes[0].height]
    for sh in shapes[1:]:
        if sh.top < ends[-1] + thr:
            bands[-1].append(sh)
            ends[-1] = max(ends[-1], sh.top + sh.height)
        else:
            bands.append([sh])
            ends.append(sh.top + sh.height)
    while len(bands) > max(n_blocks, 1):
        gaps = [min(s.top for s in bands[i + 1]) - ends[i] for i in range(len(bands) - 1)]
        i = gaps.index(min(gaps))
        bands[i].extend(bands[i + 1])
        ends[i] = max(ends[i], ends[i + 1])
        del bands[i + 1], ends[i + 1]
    return bands


def _has_visible_fill(el):
    sp_pr = el.find(qn("p:spPr"))
    if sp_pr is None:
        return False
    if any(sp_pr.find(f"{{{A_NS}}}{t}") is not None for t in ("solidFill", "gradFill", "blipFill")):
        return True
    ln = sp_pr.find(f"{{{A_NS}}}ln")
    return ln is not None and ln.find(f"{{{A_NS}}}solidFill") is not None


def xp(el, expr):
    return etree.XPath(expr, namespaces={"p": P_NS})(el)


class _Ids:
    def __init__(self, start):
        self.n = start - 1

    def __call__(self):
        self.n += 1
        return self.n


def _effect(ids, preset, spid, node, grp):
    pid, sub, flt = PRESETS[preset]
    grp_attr = ' grpId="0"' if grp else ""
    eid, set_id = ids(), ids()
    anim = (f'<p:animEffect transition="in" filter="{flt}"><p:cBhvr><p:cTn id="{ids()}" dur="500"/>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>') if flt else ""
    return (f'<p:par><p:cTn id="{eid}" presetID="{pid}" presetClass="entr" presetSubtype="{sub}" '
            f'fill="hold"{grp_attr} nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst><p:set><p:cBhvr><p:cTn id="{set_id}" dur="1" fill="hold"><p:stCondLst>'
            f'<p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
            f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:to><p:strVal val="visible"/></p:to></p:set>{anim}</p:childTnLst></p:cTn></p:par>')


def build_timing(bands, preset):
    """每个 band 一次点击：首个形状 clickEffect，其余 withEffect，均在同一个内层 par 里。"""
    ids = _Ids(3)
    clicks, blds = [], []
    for band in bands:
        outer, inner = ids(), ids()
        effects = []
        for k, sh in enumerate(sorted(band, key=lambda s: (s.top, s.left))):
            is_sp = sh._element.tag == qn("p:sp")
            effects.append(_effect(ids, preset, sh.shape_id, "clickEffect" if k == 0 else "withEffect", is_sp))
            if is_sp:
                anim_bg = ' animBg="1"' if _has_visible_fill(sh._element) else ""
                blds.append(f'<p:bldP spid="{sh.shape_id}" grpId="0"{anim_bg}/>')
        clicks.append(
            f'<p:par><p:cTn id="{outer}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
            f'<p:childTnLst><p:par><p:cTn id="{inner}" fill="hold"><p:stCondLst><p:cond delay="0"/>'
            f'</p:stCondLst><p:childTnLst>{"".join(effects)}</p:childTnLst></p:cTn></p:par></p:childTnLst>'
            f'</p:cTn></p:par>')
    xml = (f'<p:timing xmlns:p="{P_NS}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" '
           f'nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" '
           f'dur="indefinite" nodeType="mainSeq"><p:childTnLst>{"".join(clicks)}</p:childTnLst></p:cTn>'
           f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
           f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
           f'</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
           f'{"<p:bldLst>" + "".join(blds) + "</p:bldLst>" if blds else ""}</p:timing>')
    return etree.fromstring(xml)


def set_timing(slide, timing):
    sld = slide._element
    for old in sld.findall(qn("p:timing")):
        sld.remove(old)
    ext = sld.find(qn("p:extLst"))
    if ext is not None:
        ext.addprevious(timing)
    else:
        sld.append(timing)


def apply(src, deck, dst, transition, entrance):
    prs = Presentation(str(src))
    slide_h = prs.slide_height
    slides_meta = deck.get("slides", [])
    planned = {}
    for i, slide in enumerate(prs.slides):
        meta = slides_meta[i] if i < len(slides_meta) else {}
        if transition != "none" and i > 0:
            slide.transition.kind = TRANSITIONS[transition]
            slide.transition.duration = 500
        if entrance != "none" and meta.get("type") in ANIMATED_TYPES:
            n_blocks = len([b for b in meta.get("blocks", []) if isinstance(b, dict)])
            bands = split_bands(content_shapes(slide, slide_h), n_blocks)
            if bands:
                set_timing(slide, build_timing(bands, entrance))
                planned[i] = [[sh.name for sh in band] for band in bands]
    prs.save(str(dst))
    return planned


def verify(dst, transition, planned):
    """结构自检：目标 id 存在且唯一、主序列/点击分组符合 PowerPoint 写法、效果数与计划一致。"""
    errors = []
    prs = Presentation(str(dst))
    for i, slide in enumerate(prs.slides):
        n = i + 1
        if transition != "none" and i > 0 and slide.transition.kind != TRANSITIONS[transition]:
            errors.append(f"第 {n} 页转场缺失")
        sld = slide._element
        timing = sld.find(qn("p:timing"))
        if i not in planned:
            continue
        if timing is None:
            errors.append(f"第 {n} 页没有 p:timing")
            continue
        top_ids = {str(sh.shape_id) for sh in slide.shapes}
        seqs = xp(timing, './/p:cTn[@nodeType="mainSeq"]')
        if len(seqs) != 1:
            errors.append(f"第 {n} 页主序列数量为 {len(seqs)}")
            continue
        effects = xp(timing, './/p:cTn[@presetClass]')
        targets = [xp(e, './/p:spTgt/@spid')[0] for e in effects]
        if any(e.get("presetClass") != "entr" for e in effects):
            errors.append(f"第 {n} 页含非入场效果")
        if len(set(targets)) != len(targets):
            errors.append(f"第 {n} 页有形状被重复加入场")
        missing = [t for t in targets if t not in top_ids]
        if missing:
            errors.append(f"第 {n} 页动画目标不存在: {missing}")
        groups = planned[i]
        if len(effects) != sum(len(g) for g in groups):
            errors.append(f"第 {n} 页效果数 {len(effects)} 与计划 {sum(len(g) for g in groups)} 不符")
        clicks = xp(seqs[0], './p:childTnLst/p:par')
        if len(clicks) != len(groups):
            errors.append(f"第 {n} 页点击数 {len(clicks)} 与 block 数 {len(groups)} 不符")
        for c in clicks:
            nodes = [e.get("nodeType") for e in xp(c, './/p:cTn[@presetClass]')]
            if not nodes or nodes[0] != "clickEffect" or any(x != "withEffect" for x in nodes[1:]):
                errors.append(f"第 {n} 页点击分组的触发类型异常: {nodes}")
        bad = [b for b in xp(timing, './/p:bldP/@spid') if b not in targets]
        if bad:
            errors.append(f"第 {n} 页 bldP 指向未动画的形状: {bad}")
    return errors


def com_check(dst, planned):
    """用本机 PowerPoint 读回主序列，核对每页效果数、点击数与形状名。"""
    try:
        import win32com.client
    except ImportError:
        print("跳过 --com-check：未安装 pywin32", file=sys.stderr)
        return []
    errors = []
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(str(Path(dst).resolve()), True, False, False)
    try:
        for i, groups in planned.items():
            seq = pres.Slides(i + 1).TimeLine.MainSequence
            names = sorted(seq(k).Shape.Name for k in range(1, seq.Count + 1))
            clicks = sum(1 for k in range(1, seq.Count + 1) if seq(k).Timing.TriggerType == 1)
            want = sorted(n for g in groups for n in g)
            if seq.Count != len(want) or names != want or clicks != len(groups):
                errors.append(f"第 {i + 1} 页 PowerPoint 读回: 效果 {seq.Count}/{len(want)}，点击 {clicks}/{len(groups)}")
    finally:
        pres.Close()
    return errors


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx")
    ap.add_argument("--deck", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--transition", choices=["fade", "push", "wipe", "cover", "cut", "morph", "none"], default="fade")
    ap.add_argument("--entrance", choices=["fade", "wipe", "appear", "none"], default="fade")
    ap.add_argument("--com-check", action="store_true")
    a = ap.parse_args()
    src, deck_path, dst = Path(a.pptx), Path(a.deck), Path(a.out)
    if not src.is_file() or not deck_path.is_file():
        print("输入文件不存在", file=sys.stderr)
        return 2
    deck = json.loads(deck_path.read_text(encoding="utf-8"))
    planned = apply(src, deck, dst, a.transition, a.entrance)
    errors = verify(dst, a.transition, planned)
    if a.com_check:
        errors += com_check(dst, planned)
    for e in errors:
        print("校验失败:", e, file=sys.stderr)
    total = sum(len(g) for gs in planned.values() for g in gs)
    clicks = sum(len(gs) for gs in planned.values())
    print(f"{dst}  入场动画 {total} 个，点击 {clicks} 次 / {len(planned)} 页")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
