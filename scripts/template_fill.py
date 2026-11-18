"""模板模式：独立克隆所选原页及自有部件，再用 OfficeCLI 填充，不改原模板。

用法:
  python template_fill.py --template 模板.pptx --deck deck.json --map map.json --output out.pptx

map.json 格式见 references/template-inherit.md。流程：包级克隆所选页并重映射关系
→OfficeCLI 原子批量填充暂存文件→应用动画计划与讲稿→重开校验→原子发布。
strict 保持对象几何；改变对象集合须显式 custom。有公式时写出 sidecar，再跑 formula.py。

退出码: 0 成功；1 执行或校验失败；2 map/模板/参数有误。
"""
import argparse
import hashlib
import json
import math
import os
import re
import sys
import tempfile
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
    from validate_pptx import positioned_shapes
    for s in prs.slides:
        keys, ids, dup = {}, {}, set()
        displayed = {id(sh._element): sh for sh in positioned_shapes(s.shapes)}
        for rec in walk_shapes(list(s.shapes)):
            sh = displayed.get(id(rec["shape"]._element))
            if sh is not None:
                rec["display_width"] = sh.width
                rec["display_height"] = sh.height
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


CHART_TYPES = {
    "column_clustered": "column",
    "column_stacked": "stackedColumn",
    "column_stacked_100": "percentStackedColumn",
    "bar_clustered": "bar",
    "bar_stacked": "stackedBar",
    "line": "line",
    "line_markers": "line",
    "area": "area",
    "area_stacked": "stackedArea",
    "pie": "pie",
    "doughnut": "doughnut",
    "radar": "radar",
    "radar_filled": "radar",
}


def chart_props(chart, box, k):
    """Convert a map chart object into OfficeCLI native chart properties."""
    fallback = str(chart.get("fallback", "native")).strip().lower()
    if fallback not in {"native", "auto"}:
        raise MapError(f"map 第 {k} 项 chart 只支持原生图表；SVG 回退请先生成图片并使用 image 替换")
    chart_type = str(chart.get("chart_type", "column_clustered")).strip().lower()
    if chart_type not in CHART_TYPES:
        raise MapError(f"map 第 {k} 项 chart_type 暂不支持 OfficeCLI 原生映射: {chart_type!r}")
    categories = [str(value) for value in chart.get("categories") or []]
    series = chart.get("series") or []
    if not categories or not series:
        raise MapError(f"map 第 {k} 项 chart 必须同时有 categories 和 series")
    if len(categories) == 1:
        raise MapError(f"map 第 {k} 项 chart 只有一个数据点，不能表达比较关系；改用 stats/lead 或补齐可比类别")
    if not any(chart.get(field) for field in ("source", "table_view", "accessibility_note")):
        raise MapError(f"map 第 {k} 项 chart 必须提供 source、table_view 或 accessibility_note")
    data = []
    names = set()
    for item in series:
        name = str(item.get("name", "")).strip()
        values = item.get("values") or []
        if not name or name in names or len(values) != len(categories):
            raise MapError(f"map 第 {k} 项 chart 的系列名必须唯一，且数值长度要与 categories 一致")
        names.add(name)
        try:
            if not all(math.isfinite(float(value)) for value in values):
                raise ValueError
        except (TypeError, ValueError):
            raise MapError(f"map 第 {k} 项 chart 的 values 必须是有限数值")
        if any(token in name for token in (":", ";")) or any(
                token in value for value in categories for token in (",", ";")):
            raise MapError(f"map 第 {k} 项 chart 的系列名不能含 : 或 ;，分类不能含 , 或 ;")
        data.append(f"{name}:{','.join(str(value) for value in values)}")
    props = {
        "chartType": CHART_TYPES[chart_type],
        "categories": ",".join(categories),
        "data": ";".join(data),
        "name": chart.get("name", f"CHART_{k}"),
        "x": f"{box[0]}emu", "y": f"{box[1]}emu",
        "w": f"{box[2]}emu", "h": f"{box[3]}emu",
        # 主题色名由 OfficeCLI 解析，避免把模板主题硬编码成另一套颜色。
        "colors": chart.get("colors", "accent1,accent2,accent3,accent4"),
    }
    for source_key, target_key in (("data_labels", "dataLabels"), ("label_pos", "labelPos"),
                                   ("axis_font", "axisfont"), ("gridlines", "gridlines"),
                                   ("chart_area_fill", "chartareafill")):
        if source_key in chart:
            props[target_key] = chart[source_key]
    return props


def build_ops(maps, pages, map_dir):
    """返回 (ops, 期望文字校验表, 公式 sidecar 列表)。"""
    from template_capacity import check_box, check_image, check_text

    n_tpl = len(pages)
    ops, checks, formulas = [], [], []
    fno = 0
    for k, ms in enumerate(maps, 1):
        t = ms.get("template_slide")
        if not isinstance(t, int) or not 1 <= t <= n_tpl:
            raise MapError(f"map 第 {k} 项：template_slide 必须是 1..{n_tpl} 的整数，当前 {t!r}")
        page = pages[t - 1]
        cur = n_tpl + k  # 副本追加在末尾，填充完成后再删除原页
        fields = ms.get("fields") or {}
        remove_keys = list(ms.get("remove") or [])
        field_recs = {key: resolve_key(key, page, t) for key in fields}
        for key, val in fields.items():
            if isinstance(val, dict) and "chart" in val:
                rec = field_recs[key]
                # OfficeCLI 可以稳定地在页面根部添加图表；嵌套图片位需要先拆组，
                # 因此这里明确拒绝，避免生成一个看似成功但位置错误的图表。
                if rec["kind"] not in {"picture", "chart"} or len(rec["steps"]) != 1:
                    raise MapError(f"map 第 {k} 项：chart 替换只支持顶层 picture/chart，当前 {key!r} 不符合")
                remove_keys.append(key)
        removes = sorted({resolve_key(key, page, t)["steps"] for key in remove_keys}, reverse=True)
        ops.append({"command": "add", "parent": "/", "from": f"/slide[{t}]"})
        chart_add_no = 0
        for key, val in fields.items():
            rec = field_recs[key]
            base = slide_path(cur, rec["steps"])
            if isinstance(val, str) or (isinstance(val, dict) and "text" in val):
                text = val if isinstance(val, str) else val["text"]
                if rec["kind"] != "text":
                    raise MapError(f"map 第 {k} 项：" + explain_wrong_kind(key, rec, page, "填文字"))
                # An unchanged inherited label remains covered by the template
                # baseline. Capacity is a gate for new content; re-measuring a
                # mixed-run decorative word can create a false failure.
                capacity = [] if text == rec["shape"].text else check_text(
                    rec["shape"], text, f"map 第 {k} 项 {key!r}",
                    width=rec.get("display_width"), height=rec.get("display_height"),
                    size=val.get("size") if isinstance(val, dict) else None)
                if capacity:
                    raise MapError("；".join(capacity))
                props = {"text": text}
                if isinstance(val, dict) and "size" in val:
                    props["size"] = val["size"]
                if isinstance(val, dict) and "color" in val:
                    if val["color"] not in SCHEME_COLORS:
                        raise MapError("模板字段 color 只接受主题色名")
                    props["color"] = val["color"]
                ops.append({"command": "set", "path": base, "props": props})
                checks.append(("text", k, shifted(rec["steps"], removes, key), key, text))
            elif isinstance(val, dict) and "image" in val:
                if rec["kind"] != "picture":
                    raise MapError(f"map 第 {k} 项：" + explain_wrong_kind(key, rec, page, "换图片"))
                img = Path(val["image"])
                img = img if img.is_absolute() else (map_dir / img)
                if not img.exists():
                    raise MapError(f"map 第 {k} 项：图片不存在 {img}")
                resolution = check_image(img, rec["shape"], f"map 第 {k} 项 {key!r}")
                if resolution:
                    raise MapError("；".join(resolution))
                # 模板图片位带着原照片的裁剪，不重置会把新图裁掉一块
                ops.append({"command": "set", "path": base, "props": {"src": str(img.resolve()), "crop": "0"}})
                checks.append(("image", k, shifted(rec["steps"], removes, key), key,
                               hashlib.sha1(img.read_bytes()).hexdigest()))
            elif isinstance(val, dict) and "chart" in val:
                box = (rec["shape"].left, rec["shape"].top, rec["shape"].width, rec["shape"].height)
                chart_add_no += 1
                ops.append({"command": "add", "parent": f"/slide[{cur}]", "type": "chart",
                            "props": chart_props(val["chart"], box, k)})
                # 原图片/图表已在 removes 中删除，新图表追加到页面根部；不能用
                # 被删除的原 steps 做校验路径。
                existing = sum(1 for item in page["keys"].values()
                               if item["category"] == "chart" and len(item["steps"]) == 1)
                removed_charts = sum(1 for item in removes if len(item) == 1 and item[0][0] == "chart")
                output_chart_index = existing - removed_charts + chart_add_no
                checks.append(("chart", k, (("chart", output_chart_index),), key,
                               str(val["chart"].get("chart_type", "column_clustered"))))
            elif isinstance(val, dict) and "table" in val:
                if rec["kind"] != "table":
                    raise MapError(f"map 第 {k} 项：" + explain_wrong_kind(key, rec, page, "填表格"))
                for r, row in enumerate(val["table"], 1):
                    for c, cell in enumerate(row, 1):
                        ops.append({"command": "set", "path": f"{base}/tr[{r}]/tc[{c}]", "props": {"text": str(cell)}})
            else:
                raise MapError(f"map 第 {k} 项：{key!r} 的值必须是字符串、{{\"text\":..,\"size\":..}}、"
                               f"{{\"image\":路径}}、{{\"chart\":图表对象}} 或 {{\"table\":二维数组}}")
        # Legacy explicit custom interface only. Strict mode rejects cloning; the
        # semi-autonomous path clears the body canvas via template_canvas.py.
        source_group_count = max(
            (rec["steps"][0][1] for rec in page["keys"].values()
             if rec["steps"] and rec["steps"][0][0] == "group"),
            default=0,
        )
        for j, clone in enumerate(ms.get("clone_groups") or [], 1):
            source_key = clone.get("source")
            if not isinstance(source_key, str):
                raise MapError(f"map 第 {k} 项 clone_groups[{j}]：缺少 source")
            source = resolve_key(source_key, page, t)
            if source["kind"] != "group" or len(source["steps"]) != 1:
                raise MapError(f"map 第 {k} 项 clone_groups[{j}]：source 必须是模板页顶层组合")
            ops.append({"command": "add", "parent": f"/slide[{cur}]", "type": "group",
                        "from": slide_path(cur, source["steps"])})
            new_steps = (("group", source_group_count + j),)
            group_props = {}
            for field, prop in (("x", "x"), ("y", "y"), ("width", "width"), ("height", "height")):
                if field in clone:
                    group_props[prop] = f"{to_emu(clone[field], f'map 第 {k} 项 clone_groups[{j}].{field}')}emu"
            if group_props:
                ops.append({"command": "set", "path": slide_path(cur, new_steps), "props": group_props})
                checks.append(("geometry", k, shifted(new_steps, removes, source_key), source_key,
                               {field: to_emu(clone[field], field) for field in
                                ("x", "y", "width", "height") if field in clone}))
            for field_key, field_val in (clone.get("fields") or {}).items():
                child = resolve_key(field_key, page, t)
                if child["kind"] != "text":
                    raise MapError(f"map 第 {k} 项 clone_groups[{j}]：字段 {field_key!r} 必须是文字形状")
                if not field_key.startswith(source_key + "/"):
                    raise MapError(f"map 第 {k} 项 clone_groups[{j}]：字段 {field_key!r} 不在 source 组合内")
                mapped = new_steps + child["steps"][1:]
                text = field_val if isinstance(field_val, str) else (
                    field_val.get("text") if isinstance(field_val, dict) else None
                )
                if text is None:
                    raise MapError(f"map 第 {k} 项 clone_groups[{j}]：字段 {field_key!r} 只支持字符串或 {{\"text\":..}}")
                source_width = max(1, source["shape"].width or 1)
                source_height = max(1, source["shape"].height or 1)
                target_width = to_emu(clone["width"], "clone width") if "width" in clone else source_width
                target_height = to_emu(clone["height"], "clone height") if "height" in clone else source_height
                child_width = child["shape"].width * target_width / source_width
                child_height = child["shape"].height * target_height / source_height
                capacity = check_text(child["shape"], text, f"map 第 {k} 项 clone_groups[{j}] {field_key!r}",
                                      width=child_width, height=child_height,
                                      size=field_val.get("size") if isinstance(field_val, dict) else None)
                if capacity:
                    raise MapError("；".join(capacity))
                props = {"text": text}
                if isinstance(field_val, dict) and "size" in field_val:
                    props["size"] = field_val["size"]
                if isinstance(field_val, dict) and "color" in field_val:
                    if field_val["color"] not in SCHEME_COLORS:
                        raise MapError("模板 clone 字段 color 只接受主题色名")
                    props["color"] = field_val["color"]
                ops.append({"command": "set", "path": slide_path(cur, mapped), "props": props})
                checks.append(("text", k, shifted(mapped, removes, field_key), field_key, text))
        # 位置寻址下删除会让同类后续形状序号前移，所以放在填充之后，并从后往前删
        for steps in removes:
            ops.append({"command": "remove", "path": slide_path(cur, steps)})
        for j, ex in enumerate(ms.get("extra_shapes") or [], 1):
            color = ex.get("color", "tx1")
            if color not in SCHEME_COLORS:
                raise MapError(f"map 第 {k} 项 extra_shapes[{j}]：color 只能用主题色名 {sorted(SCHEME_COLORS)}，当前 {color!r}")
            props = {"text": ex["text"], "color": color, "name": f"EXTRA_{k}_{j}",
                     "size": ex.get("size", 14)}
            dimensions = {}
            for p in ("x", "y", "width", "height"):
                dimensions[p] = to_emu(ex[p], f"map 第 {k} 项 extra_shapes[{j}].{p}")
                props[p] = f"{dimensions[p]}emu"
            capacity = check_box(ex["text"], dimensions["width"], dimensions["height"],
                                 ex.get("size", 14), f"map 第 {k} 项 extra_shapes[{j}]")
            if capacity:
                raise MapError("；".join(capacity))
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


def write_speaker_notes(prs, deck):
    from speaker_notes import apply_notes
    return apply_notes(prs, deck, preserve_template=True)


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
        if kind == "chart" and not getattr(sh, "has_chart", False):
            problems.append(f"第 {k} 页 {key!r} 图表未落位（期望 OfficeCLI 原生图表 {expect}）")
        if kind == "geometry":
            for field, value in expect.items():
                actual = getattr(sh, {"x": "left", "y": "top"}.get(field, field))
                if actual is None or abs(actual - value) > 2:
                    problems.append(f"第 {k} 页 {key!r} 的 {field} 未落位：期望 {value}，实际 {actual}")
    return problems


def direct_text_fill(staging, maps, pages):
    """Apply pure text edits without OfficeCLI rewriting untouched rich text."""
    from slide_clone import replace_text
    prs = Presentation(str(staging))
    for output_index, mapping in enumerate(maps, 1):
        slide = prs.slides[output_index - 1]
        page = pages[mapping['template_slide'] - 1]
        targets = {rec['steps']: rec['shape'] for rec in walk_shapes(list(slide.shapes))}
        for key, value in (mapping.get('fields') or {}).items():
            rec = resolve_key(key, page, mapping['template_slide'])
            if rec['kind'] != 'text' or not isinstance(value, str):
                raise MapError('direct text fill received a non-text field')
            if value == rec['shape'].text:
                continue
            target = targets.get(rec['steps'])
            if target is None:
                raise MapError(f'克隆页找不到文字槽位 {key!r}')
            replace_text(target, value)
    prs.save(str(staging))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--template", required=True)
    ap.add_argument("--deck", required=True)
    ap.add_argument("--map", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--animation-policy", choices=("auto", "preserve", "strip"))
    ap.add_argument("--animation-plan-only", action="store_true", help="只输出编辑前动画计划，不调用 OfficeCLI")
    a = ap.parse_args()
    from template_animation import AnimationError, apply_plan, build_plan, verify_plan, write_plan
    tpl, out = Path(a.template).resolve(), Path(a.output).resolve()
    if tpl == out:
        print("输出路径不能等于模板路径（模板原件不允许被修改）", file=sys.stderr)
        return 2
    try:
        deck = json.loads(Path(a.deck).read_text(encoding="utf-8"))
        mapping = json.loads(Path(a.map).read_text(encoding="utf-8"))
        from template_contract import validate_map_mode
        try:
            map_mode = validate_map_mode(mapping)
        except ValueError as exc:
            raise MapError(str(exc)) from exc
        maps = mapping["slides"]
        n_deck = len(deck["slides"])
        if deck.get('notes_enabled', True) and map_mode == 'strict':
            missing = [i for i,p in enumerate(deck['slides'],1) if not p.get('notes')]
            if missing:
                raise MapError(f'新稿必须逐页显式提供 notes，缺少讲稿的页: {missing}')
        if len(maps) != n_deck:
            raise MapError(f"map 有 {len(maps)} 项，deck 有 {n_deck} 页，必须一一对应")
        pages = scan_template(str(tpl))
        if map_mode == 'strict':
            for index, ms in enumerate(maps,1):
                t = ms.get('template_slide')
                if not isinstance(t,int) or not 1 <= t <= len(pages):
                    raise MapError(f'第 {index} 页 template_slide 无效')
                page = pages[t-1]
                supplied = {resolve_key(key,page,t)['key']:value for key,value in (ms.get('fields') or {}).items()}
                slots = [r for r in page['keys'].values() if r['kind']=='text'
                         and re.fullmatch(r'(?:Text|Bullet)\d+',r['shape'].name)
                         and r['shape'].text.strip()]
                missing = [r['key'] for r in slots if not supplied.get(r['key'])
                           or isinstance(supplied[r['key']],dict) and not supplied[r['key']].get('text')]
                if missing:
                    raise MapError(f'第 {index} 页有未填的原版式语义槽位: {missing}；重组真实材料或换页，不能删卡')
        ops, checks, formulas = build_ops(maps, pages, Path(a.map).resolve().parent)
        plan = build_plan(tpl, maps, pages,
                          a.animation_policy or mapping.get("animation_policy", "auto"))
        if a.animation_plan_only:
            print(json.dumps(plan, ensure_ascii=False, indent=2))
            return 0
        for row in plan["pages"]:
            print(f"动画计划 第 {row['output_slide']} 页: {row['timing']} / 转场 {row['transition']}；"
                  + "；".join(row["reasons"]), file=sys.stderr)
    except (MapError, AnimationError) as e:
        print(f"map 有误: {e}", file=sys.stderr)
        return 2
    except (OSError, KeyError, json.JSONDecodeError) as e:
        print(f"读取输入失败: {e!r}", file=sys.stderr)
        return 2

    from pptx_package import package_problems, repair_package
    from validate_pptx import check_slide, audited_template_page, _inches

    out.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".template-fill-", dir=out.parent) as tmp:
        staging = Path(tmp) / out.name
        try:
            from slide_clone import clone_selected, prepared_edit_ops
            clone_selected(tpl, [m['template_slide'] for m in maps], staging)
            fill_ops = prepared_edit_ops(ops, len(pages))
            pure_text = all(
                op.get('command') == 'set'
                and set((op.get('props') or {})) == {'text'}
                and isinstance((op.get('props') or {}).get('text'), str)
                for op in fill_ops)
            try:
                if pure_text:
                    direct_text_fill(staging, maps, pages)
                else:
                    ocli.batch(str(staging), fill_ops)
            finally:
                ocli.close(str(staging))
            apply_plan(staging, plan)
            repairs = repair_package(staging, repair_animations=False)
            problems = (package_problems(staging) + verify_plan(staging, plan)
                        + verify(str(staging), n_deck, checks))
            prs = Presentation(str(staging))
            notes_report = write_speaker_notes(prs, deck)
            prs.save(str(staging))
            repair_package(staging, repair_animations=False)
            prs = Presentation(str(staging))
            from speaker_notes import verify_notes
            problems += verify_notes(prs, deck)
            original_prs = Presentation(str(tpl))
            warnings = []
            for index, slide in enumerate(prs.slides, 1):
                # Formula conversion is a separate postprocessing step. All
                # other layout checks must pass before publishing this output.
                page_for_check = dict(deck["slides"][index - 1])
                baseline = frozenset()
                page_for_check.pop('_baseline_shapes', None)
                if map_mode == 'strict' and maps[index - 1].get('fields'):
                    page_for_check, baseline = audited_template_page(
                        original_prs.slides[maps[index - 1]['template_slide'] - 1], slide,
                        _inches(prs.slide_width), _inches(prs.slide_height), page_for_check)
                check_slide(slide, index, page_for_check, problems, warnings,
                            _inches(prs.slide_width), _inches(prs.slide_height),
                            baseline, [], check_formulas=False)
            if map_mode == 'strict':
                from template_contract import check_contract
                try:
                    check_contract(tpl, staging, [m['template_slide'] for m in maps])
                except ValueError as exc:
                    problems.append(str(exc))
            if problems:
                print("输出校验未通过，未替换已有输出:\n  " + "\n  ".join(problems), file=sys.stderr)
                return 1
            if warnings:
                print("布局复核提示:\n  " + "\n  ".join(warnings), file=sys.stderr)
            if repairs:
                print(f"已修正包结构: {repairs}", file=sys.stderr)
            if out.exists():
                ocli.close(str(out))
            os.replace(staging, out)
        except Exception as exc:
            print(f"模板生成失败，未替换已有输出: {exc}", file=sys.stderr)
            return 1
    write_plan(plan, out)
    Path(str(out) + ".notes.json").write_text(json.dumps({"enabled": deck.get("notes_enabled", True),
            "pages": notes_report}, ensure_ascii=False, indent=2), encoding="utf-8")
    sidecar = Path(str(out) + ".formulas.json")
    if formulas:
        sidecar.write_text(json.dumps(formulas, ensure_ascii=False, indent=2), encoding="utf-8")
    else:
        sidecar.unlink(missing_ok=True)
    print(f"{out}（{n_deck} 页，公式 {len(formulas)} 个）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
