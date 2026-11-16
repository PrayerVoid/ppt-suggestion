"""Plan template animation handling before editing; enforce the plan after editing.

Policies: auto conservatively adapts edited pages; preserve refuses unsafe edits;
strip removes object timing and transitions. Never modifies the source template.
Semantic fingerprints normalize shape/time/relationship IDs after slide cloning.
"""
import argparse
import hashlib
import json
import os
import tempfile
from pathlib import Path
from zipfile import ZipFile

from lxml import etree
from power_pptx import Presentation

from analyze_template import walk_shapes
from pptx_package import P, R, PackageError, _serialize, _xml, package_problems

NS = {"p": P, "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
POLICIES = ("auto", "preserve", "strip")
DYNAMIC = {f"{{{P}}}timing", f"{{{P}}}bldLst", f"{{{P}}}transition"}


class AnimationError(ValueError):
    pass


def _nodes(root, kind):
    tags = {f"{{{P}}}transition"} if kind == "transition" else DYNAMIC - {f"{{{P}}}transition"}
    return [node for node in root.iter() if node.tag in tags
            and not any(parent.tag in tags for parent in node.iterancestors())]


def _strip(root, kind):
    for node in _nodes(root, kind):
        node.getparent().remove(node)
    # An AlternateContent containing only removed timing/transition branches
    # must not leave an empty Choice/Fallback scaffold behind.
    mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    for node in list(root.iter(f"{{{mc}}}AlternateContent")):
        if all(len(branch) == 0 for branch in node):
            node.getparent().remove(node)


def _shape_paths(owner, removes=()):
    result = {}
    for rec in walk_shapes(list(owner.shapes)):
        steps = rec["steps"]
        if any(steps[:len(rem)] == rem for rem in removes):
            continue
        shifted = []
        for level, (category, index) in enumerate(steps):
            gone = sum(1 for rem in removes if len(rem) == level + 1
                       and rem[:level] == steps[:level] and rem[level][0] == category
                       and rem[level][1] < index)
            shifted.append((category, index - gone))
        result[str(rec["id"])] = "/".join(f"{cat}[{index}]" for cat, index in shifted)
    return result


def _fingerprint(owner, kind, removes=(), transition_action="preserve"):
    root = etree.fromstring(etree.tostring(owner._element))
    if kind == "transition":
        _adapt_transition(root, transition_action)
    nodes = _nodes(root, kind)
    paths = _shape_paths(owner, removes)
    time_ids = {node.get("id"): str(index) for index, node in enumerate(
        (n for top in nodes for n in top.iter(f"{{{P}}}cTn")), 1)}

    def canonical(node):
        attrs = []
        for key, value in sorted(node.attrib.items()):
            if key == "spid":
                if value not in paths:
                    raise AnimationError(f"{owner.part.partname}: 动画目标 {value} 无法映射到保留形状")
                value = paths[value]
            elif key == "id" and node.tag == f"{{{P}}}cTn":
                value = time_ids[value]
            elif key == "val" and node.tag == f"{{{P}}}tn":
                value = time_ids.get(value, "missing-time:" + value)
            elif key.startswith(f"{{{R}}}"):
                try:
                    rel = owner.part.rels[value]
                    target = rel.target_ref if rel.is_external else hashlib.sha256(rel.target_part.blob).hexdigest()
                    value = rel.reltype + ":" + target
                except KeyError as exc:
                    raise AnimationError(f"{owner.part.partname}: 动画关系 {value} 不存在") from exc
            attrs.append((key, value))
        return [node.tag, attrs, (node.text or "").strip(),
                [canonical(child) for child in node if isinstance(child.tag, str)]]

    return hashlib.sha256(json.dumps([canonical(node) for node in nodes],
                                    ensure_ascii=False, sort_keys=True).encode()).hexdigest()


def _adapt_transition(root, action):
    if action == "strip":
        _strip(root, "transition")
    elif action == "manual":
        for transition in _nodes(root, "transition"):
            transition.attrib.pop("advTm", None)
            transition.set("advClick", "1")
            for sound in list(transition.iter(f"{{{P}}}sndAc")):
                sound.getparent().remove(sound)


def _has_morph(owner):
    return any(etree.QName(node).localname.lower() == "morph"
               for top in _nodes(owner._element, "transition") for node in top.iter()
               if isinstance(node.tag, str))


def _edited(item):
    return any(item.get(key) for key in ("fields", "remove", "clone_groups", "extra_shapes", "formulas"))


def _edit_risks(slide, item, page):
    from template_fill import resolve_key

    risks = []
    removes = [resolve_key(key, page, item["template_slide"])["steps"]
               for key in item.get("remove") or []]
    for key, value in (item.get("fields") or {}).items():
        if isinstance(value, dict) and "chart" in value:
            removes.append(resolve_key(key, page, item["template_slide"])["steps"])
    removes = sorted(set(removes))
    targets = set(slide._element.xpath(".//p:timing//@spid | .//p:bldLst//@spid"))
    removed_ids = {str(rec["id"]) for rec in page["keys"].values()
                   if any(rec["steps"][:len(rem)] == rem for rem in removes)}
    if targets & removed_ids:
        risks.append("删除或图表替换涉及动画目标/触发对象")
    if _nodes(slide._element, "timing") and any(item.get(key) for key in ("clone_groups", "extra_shapes", "formulas")):
        risks.append("新增/复制对象没有定义对应的点击顺序与动画映射")
    sensitive = set(slide._element.xpath(
        './/p:bldP[not(@build="allAtOnce")]/@spid | .//p:bldGraphic/@spid | '
        './/p:spTgt[p:txEl]/@spid'))
    for key, value in (item.get("fields") or {}).items():
        rec = resolve_key(key, page, item["template_slide"])
        if str(rec["id"]) in sensitive and (isinstance(value, str)
                or isinstance(value, dict) and any(k in value for k in ("text", "table", "chart"))):
            risks.append(f"{key} 有段落/字符范围或图形构建，内容变化需要重建时序")
    if _edited(item):
        advanced = slide._element.xpath(
            './/p:timing//p:animMotion | .//p:timing//p:animRot | .//p:timing//p:animScale | '
            './/p:timing//p:cmd | .//p:timing//p:audio | .//p:timing//p:video | '
            './/p:timing//p:cTn[@nodeType="interactiveSeq"]')
        if advanced:
            risks.append("编辑页含动作路径、交互触发、媒体或其它高级时序")
    return risks, removes


def build_plan(template, items, pages, policy="auto"):
    if policy not in POLICIES:
        raise AnimationError(f"animation_policy 必须为 {POLICIES}，当前 {policy!r}")
    from animation_inventory import inventory
    prs = Presentation(str(template))
    animation = inventory(prs)
    sources = [item["template_slide"] for item in items]
    plan = {"version": 1, "template_sha256": hashlib.sha256(Path(template).read_bytes()).hexdigest(),
            "default_policy": policy, "pages": [], "shared_parts": [],
            "playback_verified": False}
    shared = {}
    for output_index, item in enumerate(items, 1):
        selected = item.get("animation_policy", policy)
        if selected not in POLICIES:
            raise AnimationError(f"第 {output_index} 页 animation_policy 无效: {selected!r}")
        source = item["template_slide"]
        slide = prs.slides[source - 1]
        risks, removes = _edit_risks(slide, item, pages[source - 1])
        observed = animation["slides"][source - 1]
        if _edited(item) and observed["timing"]["advanced_features"]:
            risks.append("动画清单标记高级/未知特性：" + ", ".join(observed["timing"]["advanced_features"]))
        # Layout/master transitions are stored as shared defaults by PowerPoint
        # but do not carry page object targets. Only inherited timing makes a
        # page edit unsafe; page transitions are handled on the output slide.
        inherited = [owner for owner in (slide.slide_layout, slide.slide_layout.slide_master)
                     if _nodes(owner._element, "timing")]
        inherited_timing = bool(inherited)
        # Morph depends on adjacency AND the objects on both sides of a boundary.
        intact_pair = (output_index > 1 and sources[output_index - 2] == source - 1
                       and not _edited(item) and not _edited(items[output_index - 2]))
        morph = any(_has_morph(owner) for owner in (slide, *inherited))
        if morph and not intact_pair:
            risks.append("Morph 的前后页/内容对应关系已改变或首张输出缺少前页")
            if selected != "strip" and any(_has_morph(owner) for owner in inherited):
                raise AnimationError(f"第 {output_index} 页的共享母版/版式 Morph 依赖已改变；请统一 strip 或换页")
        if inherited_timing and _edited(item) and selected != "strip":
            raise AnimationError(f"第 {output_index} 页存在母版/版式动画，编辑无法安全局部适配；"
                                 "请换用无继承动画页，或显式对共享该部件的输出页使用 strip")
        if selected == "preserve" and risks:
            raise AnimationError(f"第 {output_index} 页 preserve 无法满足：" + "；".join(risks))
        timing = "strip" if selected == "strip" or selected == "auto" and risks else "preserve"
        transition = "preserve"
        if selected == "strip" or selected == "auto" and morph and not intact_pair:
            transition = "strip"
        elif selected == "auto" and _edited(item):
            transition = "manual"
        reasons = list(risks)
        if selected == "strip":
            reasons.append("显式 strip：取消对象动画、转场、自动翻页和转场声音")
        elif transition == "manual":
            reasons.append("内容已变：保留普通转场效果，改为手动翻页并取消转场声音")
        elif not reasons:
            reasons.append("动画目标和时序可保持，生成后核对语义指纹")
        row = {"output_slide": output_index, "template_slide": source, "policy": selected,
               "timing": timing, "transition": transition, "reasons": reasons,
               "layout_part": str(slide.slide_layout.part.partname),
               "master_part": str(slide.slide_layout.slide_master.part.partname),
               "source_animation": {"has_timing": observed["timing"]["present"],
                                    "has_builds": observed["timing"]["has_builds"],
                                    "inherited_timing_parts": observed["inherited_timing_parts"],
                                    "target_ids": observed["timing"]["target_ids"],
                                    "transition": observed["transition"]},
               "expected_timing": _empty_fingerprint() if timing == "strip" else _fingerprint(slide, "timing", removes),
               "expected_transition": _fingerprint(slide, "transition", transition_action=transition)}
        plan["pages"].append(row)
        for owner in inherited:
            name = str(owner.part.partname).lstrip("/")
            action = "strip" if selected == "strip" else "preserve"
            # A shared transition cannot be adapted for one page in place.
            if action == "preserve" and selected == "auto" and _edited(item) and _nodes(owner._element, "transition"):
                raise AnimationError(f"第 {output_index} 页有共享版式/母版转场；请显式 preserve、strip 或换页")
            if name in shared and shared[name]["action"] != action:
                raise AnimationError(f"共享部件 {name} 同时被要求保留与清除；请统一相关页策略或换页")
            shared[name] = {"part": name, "action": action,
                            "expected_timing": _empty_fingerprint() if action == "strip" else _fingerprint(owner, "timing"),
                            "expected_transition": _empty_fingerprint() if action == "strip" else _fingerprint(owner, "transition")}
    plan["shared_parts"] = list(shared.values())
    return plan


def _empty_fingerprint():
    return hashlib.sha256(b"[]").hexdigest()


def write_plan(plan, output):
    target = Path(str(output) + ".animation-plan.json")
    target.write_text(json.dumps(plan, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return target


def _rewrite(path, replacements):
    if not replacements:
        return
    path = Path(path)
    temporary = None
    try:
        with ZipFile(path) as package:
            fd, temporary = tempfile.mkstemp(prefix=".animation-", suffix=".pptx", dir=path.parent)
            os.close(fd)
            with ZipFile(temporary, "w") as result:
                result.comment = package.comment
                for info in package.infolist():
                    result.writestr(info, replacements.get(info.filename, package.read(info)))
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            Path(temporary).unlink(missing_ok=True)


def apply_plan(path, plan):
    """Execute only the removals declared before editing, on an output staging file."""
    prs = Presentation(str(path))
    # Accessing prs.slides renames slide parts in memory. Keep the actual ZIP
    # names before that happens (OfficeCLI may leave slide25.xml as page 1).
    part_names = {id(part): str(part.partname).lstrip("/") for part in prs.part.package.iter_parts()}
    if len(prs.slides) != len(plan["pages"]):
        raise AnimationError("动画计划与输出页数不一致")
    actions = {part_names[id(slide.part)]: row
               for slide, row in zip(prs.slides, plan["pages"])}
    for row in plan["shared_parts"]:
        actions[row["part"]] = {"timing": row["action"], "transition": row["action"]}
    replacements = {}
    with ZipFile(path) as package:
        for name, row in actions.items():
            if name not in package.namelist():
                raise AnimationError(f"动画计划的共享部件丢失: {name}")
            root = _xml(package.read(name))
            before = etree.tostring(root)
            if row["timing"] == "strip":
                _strip(root, "timing")
            _adapt_transition(root, row["transition"])
            if etree.tostring(root) != before:
                replacements[name] = _serialize(root)
    _rewrite(path, replacements)


def _raw_policy_problems(root, row, label):
    """Check the ZIP XML too; power_pptx may expose only one MC transition branch."""
    problems = []
    timings = _nodes(root, "timing")
    transitions = _nodes(root, "transition")
    if row.get("timing") == "strip" and timings:
        problems.append(f"{label}: 计划要求移除对象时序，但 ZIP 仍有 p:timing/p:bldLst")
    if row.get("transition") == "strip" and transitions:
        problems.append(f"{label}: 计划要求移除转场，但 ZIP 仍有 p:transition")
    if row.get("transition") == "manual":
        for transition in transitions:
            if "advTm" in transition.attrib or transition.get("advClick") != "1":
                problems.append(f"{label}: 计划要求手动翻页，但转场仍含自动推进")
            if list(transition.iter(f"{{{P}}}sndAc")):
                problems.append(f"{label}: 计划要求取消转场声音，但 ZIP 仍有声音节点")
    return problems


def verify_plan(path, plan):
    prs = Presentation(str(path))
    problems = []
    if len(prs.slides) != len(plan["pages"]):
        return ["动画计划与最终页数不一致"]
    owners = {}
    with ZipFile(path) as package:
        for slide, row in zip(prs.slides, plan["pages"]):
            part_name = str(slide.part.partname).lstrip("/")
            if part_name in package.namelist():
                problems.extend(_raw_policy_problems(_xml(package.read(part_name)), row,
                                                     f"第 {row['output_slide']} 页"))
    for slide, row in zip(prs.slides, plan["pages"]):
        if (str(slide.slide_layout.part.partname) != row["layout_part"]
                or str(slide.slide_layout.slide_master.part.partname) != row["master_part"]):
            problems.append(f"第 {row['output_slide']} 页的版式/母版关联与计划不符")
        owners[str(slide.slide_layout.part.partname).lstrip("/")] = slide.slide_layout
        owners[str(slide.slide_layout.slide_master.part.partname).lstrip("/")] = slide.slide_layout.slide_master
        for kind in ("timing", "transition"):
            try:
                if _fingerprint(slide, kind) != row[f"expected_{kind}"]:
                    problems.append(f"第 {row['output_slide']} 页 {kind} 与编辑前动画计划不符（目标、顺序或效果变化）")
            except AnimationError as exc:
                problems.append(str(exc))
    for row in plan["shared_parts"]:
        owner = owners.get(row["part"])
        with ZipFile(path) as package:
            if row["part"] in package.namelist():
                problems.extend(_raw_policy_problems(_xml(package.read(row["part"])),
                                                     {"timing": row["action"], "transition": row["action"]},
                                                     row["part"]))
        if owner is None:
            problems.append(f"共享动画部件不再关联预期页面: {row['part']}")
            continue
        for kind in ("timing", "transition"):
            try:
                if _fingerprint(owner, kind) != row[f"expected_{kind}"]:
                    problems.append(f"{row['part']}: 共享 {kind} 与计划不符")
            except AnimationError as exc:
                problems.append(str(exc))
    return problems


def main():
    parser = argparse.ArgumentParser(description="在公式、图表和动效等最后一次后处理后复核模板动画计划")
    parser.add_argument("pptx")
    parser.add_argument("--plan", required=True)
    args = parser.parse_args()
    try:
        plan = json.loads(Path(args.plan).read_text(encoding="utf-8"))
        problems = package_problems(args.pptx) + verify_plan(args.pptx, plan)
        if problems:
            print("\n".join(problems))
            return 1
    except (OSError, ValueError, KeyError, PackageError) as exc:
        print(f"动画核对失败: {exc}")
        return 1
    print("动画计划核对通过；放映播放尚需单独验证")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
