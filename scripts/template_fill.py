"""模板模式：复制模板副本，用 officecli 复制原页并填充内容，不改原模板。

用法:
  python template_fill.py --template 模板.pptx --deck deck.json --map map.json --output out.pptx

map.json 格式见 references/template-inherit.md。流程：复制模板→批量（原子）复制所选原页并填充
→删除模板原页→落盘→重开校验。有公式时写出 <output>.formulas.json，再跑 formula.py。

退出码: 0 成功；1 执行或校验失败；2 map/模板/参数有误。
"""
import argparse
import hashlib
import json
import re
import shutil
import sys
from pathlib import Path

import ocli
from analyze_template import walk_shapes
from power_pptx import Presentation

SCHEME_COLORS = {"tx1", "tx2", "bg1", "bg2", "dk1", "dk2", "lt1", "lt2",
                 "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"}
UNITS = {"emu": 1, "cm": 360000, "mm": 36000, "in": 914400, "pt": 12700}


class MapError(Exception):
    pass


def to_emu(v, what):
    if isinstance(v, (int, float)):
        return int(v)
    m = re.fullmatch(r"\s*(-?\d+(?:\.\d+)?)\s*(emu|cm|mm|in|pt)\s*", str(v))
    if not m:
        raise MapError(f"{what} 的长度写法无法识别: {v!r}（用 EMU 整数，或 2cm / 10mm / 1in / 12pt）")
    return int(float(m.group(1)) * UNITS[m.group(2)])


def scan_template(path):
    """每页返回 {"keys": key→记录, "ids": id→[key], "dup": 重复的 key 集合}。"""
    prs = Presentation(path)
    pages = []
    for s in prs.slides:
        keys, ids, dup = {}, {}, set()
        for rec in walk_shapes(list(s.shapes)):
            if rec["key"] in keys:
                dup.add(rec["key"])
            keys[rec["key"]] = rec
            ids.setdefault(rec["id"], []).append(rec["key"])
        pages.append({"keys": keys, "ids": ids, "dup": dup})
    return pages


def resolve_key(key, page, tslide):
    """把 map 里的 key 解析为模板页上的形状记录；`@id=N` 按模板页原 id 换算成同一条记录。"""
    if key.startswith("@id="):
        try:
            i = int(key[4:])
        except ValueError:
            raise MapError(f"模板第 {tslide} 页：{key} 不是合法的 @id=数字")
        hits = page["ids"].get(i)
        if not hits:
            raise MapError(f"模板第 {tslide} 页没有 id={i} 的形状；改用 key，可选: {sorted(page['keys'])}")
        if len(hits) > 1:
            raise MapError(f"模板第 {tslide} 页有多个 id={i} 的形状 {hits}，请改用 key")
        return page["keys"][hits[0]]
    rec = page["keys"].get(key)
    if rec is None:
        raise MapError(f"模板第 {tslide} 页没有 key 为 {key!r} 的形状；可选: {sorted(page['keys'])}")
    if key in page["dup"]:
        raise MapError(f"模板第 {tslide} 页 key {key!r} 不唯一（形状名含 / 或 #），请改用 @id=N")
    return rec


def slide_path(cur, steps):
    return f"/slide[{cur}]" + "".join(f"/{cat}[{i}]" for cat, i in steps)


def shifted(steps, removes, key):
    """删除同级同类的前序形状后，剩下形状的序号要前移；返回输出页上的 steps。"""
    out = []
    for j, (cat, idx) in enumerate(steps):
        if steps[:j + 1] in removes:
            raise MapError(f"{key!r} 同时出现在 fields 和 remove（或其所在组合被删除）")
        gone = sum(1 for r in removes if len(r) == j + 1 and r[:j] == steps[:j] and r[j][0] == cat and r[j][1] < idx)
        out.append((cat, idx - gone))
    return tuple(out)


def explain_wrong_kind(key, rec, page, want):
    if rec["kind"] == "group":
        kids = [k for k in page["keys"] if k.startswith(rec["key"] + "/")]
        return f"{key!r} 是组合，不能直接{want}；请写组合内子形状的 key，如 {kids[:6]}"
    return f"{key!r} 是{rec['kind']}（{rec['category']}），不能{want}"


def build_ops(maps, pages, map_dir):
    """返回 (ops, 期望文字校验表, 公式 sidecar 列表)。"""
    n_tpl = len(pages)
    ops, checks, formulas = [], [], []
    fno = 0
    for k, ms in enumerate(maps, 1):
        t = ms.get("template_slide")
        if not isinstance(t, int) or not 1 <= t <= n_tpl:
            raise MapError(f"map 第 {k} 项：template_slide 必须是 1..{n_tpl} 的整数，当前 {t!r}")
        page = pages[t - 1]
        cur = n_tpl + k  # 副本追加在末尾，填充完成后再删除原页
        removes = sorted((resolve_key(key, page, t)["steps"] for key in ms.get("remove") or []), reverse=True)
        ops.append({"command": "add", "parent": "/", "from": f"/slide[{t}]"})
        for key, val in (ms.get("fields") or {}).items():
            rec = resolve_key(key, page, t)
            base = slide_path(cur, rec["steps"])
            if isinstance(val, str) or (isinstance(val, dict) and "text" in val):
                text = val if isinstance(val, str) else val["text"]
                if rec["kind"] != "text":
                    raise MapError(f"map 第 {k} 项：" + explain_wrong_kind(key, rec, page, "填文字"))
                props = {"text": text}
                if isinstance(val, dict) and "size" in val:
                    props["size"] = val["size"]
                ops.append({"command": "set", "path": base, "props": props})
                checks.append(("text", k, shifted(rec["steps"], removes, key), key, text))
            elif isinstance(val, dict) and "image" in val:
                if rec["kind"] != "picture":
                    raise MapError(f"map 第 {k} 项：" + explain_wrong_kind(key, rec, page, "换图片"))
                img = Path(val["image"])
                img = img if img.is_absolute() else (map_dir / img)
                if not img.exists():
                    raise MapError(f"map 第 {k} 项：图片不存在 {img}")
                # 模板图片位带着原照片的裁剪，不重置会把新图裁掉一块
                ops.append({"command": "set", "path": base, "props": {"src": str(img.resolve()), "crop": "0"}})
                checks.append(("image", k, shifted(rec["steps"], removes, key), key,
                               hashlib.sha1(img.read_bytes()).hexdigest()))
            elif isinstance(val, dict) and "table" in val:
                if rec["kind"] != "table":
                    raise MapError(f"map 第 {k} 项：" + explain_wrong_kind(key, rec, page, "填表格"))
                for r, row in enumerate(val["table"], 1):
                    for c, cell in enumerate(row, 1):
                        ops.append({"command": "set", "path": f"{base}/tr[{r}]/tc[{c}]", "props": {"text": str(cell)}})
            else:
                raise MapError(f"map 第 {k} 项：{key!r} 的值必须是字符串、{{\"text\":..,\"size\":..}}、"
                               f"{{\"image\":路径}} 或 {{\"table\":二维数组}}")
        # 位置寻址下删除会让同类后续形状序号前移，所以放在填充之后，并从后往前删
        for steps in removes:
            ops.append({"command": "remove", "path": slide_path(cur, steps)})
        for j, ex in enumerate(ms.get("extra_shapes") or [], 1):
            color = ex.get("color", "tx1")
            if color not in SCHEME_COLORS:
                raise MapError(f"map 第 {k} 项 extra_shapes[{j}]：color 只能用主题色名 {sorted(SCHEME_COLORS)}，当前 {color!r}")
            props = {"text": ex["text"], "color": color, "name": f"EXTRA_{k}_{j}",
                     "size": ex.get("size", 14)}
            for p in ("x", "y", "width", "height"):
                props[p] = f"{to_emu(ex[p], f'map 第 {k} 项 extra_shapes[{j}].{p}')}emu"
            if ex.get("bold"):
                props["bold"] = True
            ops.append({"command": "add", "parent": f"/slide[{cur}]", "type": "textbox", "props": props})
        for f in ms.get("formulas") or []:
            fno += 1
            box = {p: to_emu(f[p], f"map 第 {k} 项 formulas.{p}") for p in ("x", "y", "width", "height")}
            name = f"FORMULA_{fno}"
            ops.append({"command": "add", "parent": f"/slide[{cur}]", "type": "textbox", "props": {
                "text": f["latex"], "name": name, "size": 14,
                **{p: f"{v}emu" for p, v in box.items()}}})
            formulas.append({"slide": k, "shape": name, "latex": f["latex"], "display": bool(f.get("display", True)),
                             "left": box["x"], "top": box["y"], "width": box["width"], "height": box["height"]})
    for i in range(n_tpl, 0, -1):
        ops.append({"command": "remove", "path": f"/slide[{i}]"})
    return ops, checks, formulas


def verify(out, n_expected, checks):
    prs = Presentation(out)
    slides = list(prs.slides)
    problems = []
    if len(slides) != n_expected:
        problems.append(f"页数应为 {n_expected}，实际 {len(slides)}")
        return problems
    index = {}
    for kind, k, steps, key, expect in checks:
        if k not in index:
            index[k] = {rec["steps"]: rec["shape"] for rec in walk_shapes(list(slides[k - 1].shapes))}
        sh = index[k].get(steps)
        if sh is None:
            problems.append(f"第 {k} 页找不到形状 {key!r}")
            continue
        if kind == "text" and sh.text_frame.text != expect:
            problems.append(f"第 {k} 页 {key!r} 文字未落位: 期望 {expect!r}，实际 {sh.text_frame.text!r}")
        if kind == "image" and sh.image.sha1 != expect:
            problems.append(f"第 {k} 页 {key!r} 图片与 map 指定文件不一致")
    return problems


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--template", required=True)
    ap.add_argument("--deck", required=True)
    ap.add_argument("--map", required=True)
    ap.add_argument("--output", required=True)
    a = ap.parse_args()
    tpl, out = Path(a.template).resolve(), Path(a.output).resolve()
    if tpl == out:
        print("输出路径不能等于模板路径（模板原件不允许被修改）", file=sys.stderr)
        return 2
    try:
        deck = json.loads(Path(a.deck).read_text(encoding="utf-8"))
        maps = json.loads(Path(a.map).read_text(encoding="utf-8"))["slides"]
        n_deck = len(deck["slides"])
        if len(maps) != n_deck:
            raise MapError(f"map 有 {len(maps)} 项，deck 有 {n_deck} 页，必须一一对应")
        pages = scan_template(str(tpl))
        ops, checks, formulas = build_ops(maps, pages, Path(a.map).resolve().parent)
    except MapError as e:
        print(f"map 有误: {e}", file=sys.stderr)
        return 2
    except (OSError, KeyError, json.JSONDecodeError) as e:
        print(f"读取输入失败: {e!r}", file=sys.stderr)
        return 2

    out.parent.mkdir(parents=True, exist_ok=True)
    if out.exists():
        ocli.close(str(out))  # 上次运行的常驻进程可能还占着输出文件，覆盖前先释放
    shutil.copyfile(tpl, out)
    try:
        ocli.batch(str(out), ops)
    except ocli.OfficeCliError as e:
        ocli.close(str(out))
        out.unlink(missing_ok=True)
        print(f"officecli 批量编辑失败，整批已回滚，未生成输出:\n{e}", file=sys.stderr)
        return 1
    ocli.close(str(out))

    problems = verify(str(out), n_deck, checks)
    if problems:
        print("输出校验未通过:\n  " + "\n  ".join(problems), file=sys.stderr)
        return 1
    sidecar = Path(str(out) + ".formulas.json")
    if formulas:
        sidecar.write_text(json.dumps(formulas, ensure_ascii=False, indent=2), encoding="utf-8")
    else:
        sidecar.unlink(missing_ok=True)
    print(f"{out}（{n_deck} 页，公式 {len(formulas)} 个）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
