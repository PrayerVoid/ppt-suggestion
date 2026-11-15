"""把公式占位框与行内 `$...$` 替换为 PowerPoint 原生可编辑公式（OMML）。

用法: python formula.py out.pptx [--sidecar out.pptx.formulas.json]
1. 块级公式：读取渲染器写出的 sidecar，用 officecli 在原位置新增公式对象并删除占位框。
2. 行内公式：扫描所有幻灯片的全部文字（含组合形状、表格单元格），把 `$LaTeX$` 换成行内公式；
   `\\$` 表示字面美元符号。同一段落里被拆成多个 run 的 `$...$` 会先合并再判断。
任一公式写入失败则非 0 退出、不改文件，不会降级成图片。
"""
import argparse
import copy
import json
import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

import ocli

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
}
A, M, MC, A14 = (f"{{{NS[k]}}}" for k in ("a", "m", "mc", "a14"))
MATH_FONT = "Cambria Math"
INLINE_RE = re.compile(r"(?<!\\)\$(.+?)(?<!\\)\$")
BARE_DOLLAR_RE = re.compile(r"(?<!\\)\$")


def build_ops(items):
    ops = []
    for it in items:
        s = it["slide"]
        ops.append({
            "command": "add", "parent": f"/slide[{s}]", "type": "equation",
            "props": {
                "formula": it["latex"],
                "x": f"{it['left']}emu", "y": f"{it['top']}emu",
                "width": f"{it['width']}emu", "height": f"{it['height']}emu",
            },
        })
        ops.append({"command": "remove", "path": f"/slide[{s}]/shape[@name={it['shape']}]"})
    return ops


def apply_block_formulas(pptx, sidecar):
    items = json.loads(sidecar.read_text(encoding="utf-8"))
    if not items:
        return 0
    try:
        ocli.batch(pptx, build_ops(items))
    finally:
        ocli.close(pptx)
    return len(items)


def _make_blank_pptx(path):
    from power_pptx import Presentation
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.save(path)


def latex_to_omml(latexes):
    """用 officecli 把 LaTeX 转成 OMML，返回 {latex: (m:oMath 元素, 回退纯文本)}。"""
    uniq = list(dict.fromkeys(latexes))
    out = {}
    with tempfile.TemporaryDirectory() as tmp:
        probe = str(Path(tmp) / "probe.pptx")
        _make_blank_pptx(probe)
        try:
            ocli.batch(probe, [
                {"command": "add", "parent": "/slide[1]", "type": "equation",
                 "props": {"formula": lx}} for lx in uniq
            ])
        except ocli.OfficeCliError as e:
            raise ocli.OfficeCliError(f"officecli 无法写入这些行内公式 {uniq}：\n{e}")
        finally:
            ocli.close(probe)
        with zipfile.ZipFile(probe) as z:
            root = etree.fromstring(z.read("ppt/slides/slide1.xml"))
    omaths = list(root.iter(f"{M}oMath"))
    if len(omaths) != len(uniq):
        raise ocli.OfficeCliError(f"行内公式转换数量不符：期望 {len(uniq)}，得到 {len(omaths)}")
    for lx, om in zip(uniq, omaths):
        fb = om
        while fb is not None and fb.tag != f"{MC}AlternateContent":
            fb = fb.getparent()
        fallback = "".join(fb.xpath(".//mc:Fallback//a:t/text()", namespaces=NS)) if fb is not None else lx
        out[lx] = (om, fallback or lx)
    return out


def _math_rpr(src_rpr, italic=True):
    """公式 run 属性：沿用所在 run 的字号与颜色，字体用 Cambria Math 才能按数学排版。"""
    rpr = etree.Element(f"{A}rPr", lang="en-US")
    if src_rpr is not None:
        for k in ("sz", "b"):
            if src_rpr.get(k):
                rpr.set(k, src_rpr.get(k))
        fill = src_rpr.find(f"{A}solidFill")
        if fill is not None:
            rpr.append(copy.deepcopy(fill))
    if italic:
        rpr.set("i", "1")
    for tag in ("latin", "cs"):
        etree.SubElement(rpr, f"{A}{tag}", typeface=MATH_FONT)
    return rpr


def _inline_element(omath, fallback, src_rpr):
    """生成 mc:AlternateContent：Choice 里是行内 m:oMath，Fallback 是普通文字 run。"""
    ac = etree.Element(f"{MC}AlternateContent", nsmap={"mc": NS["mc"]})
    choice = etree.SubElement(ac, f"{MC}Choice", Requires="a14", nsmap={"a14": NS["a14"]})
    holder = etree.SubElement(choice, f"{A14}m")
    om = copy.deepcopy(omath)
    for r in om.iter(f"{M}r"):
        for old in r.findall(f"{A}rPr"):
            r.remove(old)
        pos = 1 if r.find(f"{M}rPr") is not None else 0
        upright = r.find(f"{M}rPr/{M}nor") is not None
        r.insert(pos, _math_rpr(src_rpr, italic=not upright))
    holder.append(om)
    fb = etree.SubElement(ac, f"{MC}Fallback")
    run = etree.SubElement(fb, f"{A}r")
    if src_rpr is not None:
        run.append(copy.deepcopy(src_rpr))
    etree.SubElement(run, f"{A}t").text = fallback
    return ac


def _run_groups(p):
    """段落里连续的 a:r 视为一组；a:br 等把段落切成多组，`$...$` 不能跨组。"""
    groups, cur = [], []
    for ch in p:
        if ch.tag == f"{A}r":
            cur.append(ch)
        elif cur:
            groups.append(cur)
            cur = []
    if cur:
        groups.append(cur)
    return groups


def _run_text(r):
    return "".join(t.text or "" for t in r.findall(f"{A}t"))


def _split_text(text):
    """把文本切成 ('t', 起, 止) 与 ('m', 起, 止, latex) 片段；未配对的 `$` 返回错误。"""
    segs, pos = [], 0
    for m in INLINE_RE.finditer(text):
        if m.start() > pos:
            segs.append(("t", pos, m.start()))
        segs.append(("m", m.start(), m.end(), m.group(1).strip()))
        pos = m.end()
    if pos < len(text):
        segs.append(("t", pos, len(text)))
    for s in segs:
        if s[0] == "t" and BARE_DOLLAR_RE.search(text[s[1]:s[2]]):
            return None
        if s[0] == "m" and not s[3]:
            return None
    return segs


def scan_inline(slide_roots):
    """返回 (待替换的 [(段落, 组, 片段, 全文)], 需要转换的 LaTeX 列表, 错误列表)。"""
    jobs, latexes, errs = [], [], []
    for sidx, root in slide_roots:
        for p in root.iter(f"{A}p"):
            for grp in _run_groups(p):
                text = "".join(_run_text(r) for r in grp)
                if "$" not in text:
                    continue
                segs = _split_text(text)
                if segs is None:
                    errs.append(f"第 {sidx} 页文字「{text[:30]}」里的 $ 没有配对或公式为空（字面美元符号写成 \\$）")
                    continue
                if not any(s[0] == "m" for s in segs):
                    if "\\$" in text:
                        jobs.append((p, grp, segs, text))
                    continue
                jobs.append((p, grp, segs, text))
                latexes += [s[3] for s in segs if s[0] == "m"]
    return jobs, latexes, errs


def _runs_for_range(grp, text, start, end):
    """把 [start,end) 字符范围映射回原 run，返回 [(run, 子串)]。"""
    out, pos = [], 0
    for r in grp:
        t = _run_text(r)
        lo, hi = max(start, pos), min(end, pos + len(t))
        if lo < hi:
            out.append((r, t[lo - pos:hi - pos]))
        pos += len(t)
    return out


def _rpr_at(grp, start):
    pos = 0
    for r in grp:
        t = _run_text(r)
        if pos <= start < pos + max(len(t), 1):
            return r.find(f"{A}rPr")
        pos += len(t)
    return grp[0].find(f"{A}rPr")


def rewrite(jobs, omml):
    for p, grp, segs, text in jobs:
        new = []
        for seg in segs:
            if seg[0] == "t":
                for r, sub in _runs_for_range(grp, text, seg[1], seg[2]):
                    nr = copy.deepcopy(r)
                    for t in nr.findall(f"{A}t")[1:]:
                        nr.remove(t)
                    nr.find(f"{A}t").text = sub.replace("\\$", "$")
                    new.append(nr)
            else:
                om, fb = omml[seg[3]]
                new.append(_inline_element(om, fb, _rpr_at(grp, seg[1])))
        anchor = grp[0]
        for el in new:
            anchor.addprevious(el)
        for r in grp:
            p.remove(r)


def apply_inline_formulas(pptx):
    """扫描并改写所有幻灯片；全部成功才落盘。返回替换的行内公式个数。"""
    try:
        ocli.close(pptx)  # 常驻进程仍持有文件时无法覆盖
    except ocli.OfficeCliError:
        pass
    with zipfile.ZipFile(pptx) as z:
        names = [n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)]
        names.sort(key=lambda n: int(re.search(r"(\d+)\.xml", n).group(1)))
        roots = {n: etree.fromstring(z.read(n)) for n in names}
    slide_roots = [(i + 1, roots[n]) for i, n in enumerate(names)]
    jobs, latexes, errs = scan_inline(slide_roots)
    if errs:
        raise ocli.OfficeCliError("\n".join(errs))
    if not jobs:
        return 0
    omml = latex_to_omml(latexes) if latexes else {}
    rewrite(jobs, omml)
    fd, tmp = tempfile.mkstemp(suffix=".pptx", dir=str(Path(pptx).parent))
    os.close(fd)
    try:
        with zipfile.ZipFile(pptx) as zin, zipfile.ZipFile(tmp, "w") as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename in roots:
                    data = etree.tostring(roots[info.filename], xml_declaration=True,
                                          encoding="UTF-8", standalone=True)
                zout.writestr(info, data, compress_type=info.compress_type)
        os.replace(tmp, pptx)
    except BaseException:
        if os.path.exists(tmp):
            os.unlink(tmp)
        raise
    return len(latexes)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pptx")
    ap.add_argument("--sidecar")
    a = ap.parse_args()
    pptx = str(Path(a.pptx).resolve())
    sidecar = Path(a.sidecar or a.pptx + ".formulas.json")
    try:
        blocks = apply_block_formulas(pptx, sidecar) if sidecar.exists() else 0
        inline = apply_inline_formulas(pptx)
    except ocli.OfficeCliError as e:
        print(f"公式写入失败：\n{e}", file=sys.stderr)
        return 1
    print(f"已写入 {blocks} 个块级公式、{inline} 个行内公式。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
