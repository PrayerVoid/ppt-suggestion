#!/usr/bin/env python3
"""Read-only OOXML animation inventory; this does not verify playback.

``inventory(prs)`` accepts a power_pptx Presentation and returns JSON-compatible
masters/layouts/slides/summary. IDs are strings in their source part's cNvPr
namespace; slide numbers are one-based, paragraph ranges are zero-based OOXML
values. Keys follow analyze_template's group/name#ordinal convention. All MC
branches are inspected, without selecting the branch a player would use.

Run: python -I scripts/animation_inventory.py template.pptx
"""
from __future__ import annotations

from collections import Counter

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
SHAPES = {"sp", "grpSp", "pic", "cxnSp", "graphicFrame", "contentPart"}
# Known structural/effect tags. Extension tags are retained as unknown features.
TIMING_TAGS = set("""
    timing tnLst par seq excl cTn stCondLst endCondLst endSync cond rtn tn
    childTnLst subTnLst prevCondLst nextCondLst iterate tmAbs tmPct
    anim animClr animEffect animMotion animRot animScale cmd set audio video
    cBhvr cMediaNode tgtEl spTgt sndTgt sldTgt inkTgt txEl pRg charRg
    subSp graphicEl chart bldLst bldP bldDgm bldOleChart bldGraphic bldSub
    bldAsOne bldAsOneSp bldAsOneDgm bldAsOneChart bldDgmChart bldChart
    tmplLst tmpl attrNameLst attrName tavLst tav val strVal boolVal intVal
    fltVal clrVal from to by rCtr ptsTypes pt rgb hsl progress
""".split())


def _tag(node):
    return node.tag.rsplit("}", 1)[-1] if isinstance(node.tag, str) else ""


def _p(node, name):
    return node.tag == f"{{{P}}}{name}"


def _number(value):
    try:
        return int(value) if value is not None else None
    except ValueError:
        return value  # Preserve malformed values for the caller's policy checks.


def _bool(value, default=False):
    return default if value is None else value in ("1", "true", "on")


def _unique(values):
    return list(dict.fromkeys(values))


def _branches(parent):
    """Flatten MC wrappers while keeping all alternatives visible."""
    for child in parent:
        if isinstance(child.tag, str) and child.tag.startswith(f"{{{MC}}}"):
            yield from _branches(child)
        else:
            yield child


def _shape_index(root):
    tree = root.find(f"{{{P}}}cSld/{{{P}}}spTree")
    index = {}
    if tree is None:
        return index
    # Includes group/root cNvPr, which share this part's ID domain.
    for node in tree.iter(f"{{{P}}}cNvPr"):
        if node.get("id") is not None:
            index.setdefault(node.get("id"), []).append(
                {"name": node.get("name", ""), "key": None})

    def walk(parent, prefix=""):
        siblings = [n for n in _branches(parent)
                    if isinstance(n.tag, str) and n.tag.startswith(f"{{{P}}}")
                    and _tag(n) in SHAPES]
        named = []
        for shape in siblings:
            nv = next((n for n in shape if _tag(n).startswith("nv")), None)
            cnv = nv.find(f"{{{P}}}cNvPr") if nv is not None else None
            if cnv is not None:
                named.append((shape, cnv))
        counts = Counter(n.get("name", "") for _, n in named)
        seen = Counter()
        for shape, cnv in named:
            name = cnv.get("name", "")
            seen[name] += 1
            label = f"{name}#{seen[name]}" if counts[name] > 1 else name
            key = f"{prefix}/{label}" if prefix else label
            hits = index.get(cnv.get("id"), [])
            # Pair duplicate IDs in XML order; ambiguous targets stay unresolved.
            hit = next((h for h in hits if h["key"] is None and h["name"] == name), None)
            if hit is not None:
                hit["key"] = key
            if _p(shape, "grpSp"):
                walk(shape, key)
    walk(tree)
    return index


def _target_ids(nodes):
    return _unique(n.get("spid") for n in nodes if n.get("spid") is not None)


def _timing(root, shapes):
    timings = list(root.iter(f"{{{P}}}timing"))
    builds = list(root.iter(f"{{{P}}}bldLst"))
    # A bldLst is usually inside timing, but standalone/MC builds matter too.
    nodes = [n for n in root.iter() if isinstance(n.tag, str) and
             (n in timings or n in builds or any(a in timings or a in builds
                                                for a in n.iterancestors()))]
    target_ids = _target_ids(nodes)
    features, triggers, paragraphs, chars, effects, behaviors = [], [], [], [], [], []
    sequences = {"main": [], "interactive": []}
    for node in nodes:
        tag = _tag(node)
        if _p(node, "cTn"):
            attrs = dict(node.attrib)
            kind = node.get("nodeType")
            if kind in ("mainSeq", "interactiveSeq"):
                sequences["main" if kind == "mainSeq" else "interactive"].append(attrs)
            if kind == "interactiveSeq":
                features.append("interactive_sequence")
            if node.get("presetClass"):
                effects.append(attrs)
        if _p(node, "cond"):
            triggers.append({"event": node.get("evt"), "delay": node.get("delay"),
                             "condition_list": _tag(node.getparent()),
                             "target_ids": _target_ids(node.iter()),
                             "time_node_ids": [n.get("val") for n in node.iter(f"{{{P}}}tn")]})
        if _p(node, "pRg") or _p(node, "charRg"):
            target = next((a for a in node.iterancestors() if _p(a, "spTgt")), None)
            item = {"spid": target.get("spid") if target is not None else None,
                    "start": _number(node.get("st")), "end": _number(node.get("end"))}
            (paragraphs if tag == "pRg" else chars).append(item)
        if tag in {"animMotion", "audio", "video", "cmd", "sndTgt"} and node.tag.startswith(f"{{{P}}}"):
            behaviors.append({"type": tag, "attributes": dict(node.attrib),
                              "target_ids": _target_ids(node.iter())})
        if _p(node, "animMotion"):
            features.append("motion")
        if _p(node, "audio") or _p(node, "video") or _p(node, "sndTgt"):
            features.append("media")
        if _p(node, "cmd"):
            features.append("command")
        if _p(node, "animClr") or _p(node, "animRot") or _p(node, "animScale"):
            features.append(tag)
        if _p(node, "bldDgm") or _p(node, "bldOleChart") or _p(node, "bldGraphic"):
            features.append(tag)
        if node.tag.startswith(f"{{{MC}}}"):
            features.append("alternate_content")
        elif not (node.tag.startswith(f"{{{P}}}") and tag in TIMING_TAGS):
            # DrawingML colors are understood values rather than new behaviors.
            if not node.tag.startswith("{http://schemas.openxmlformats.org/drawingml/2006/main}"):
                features.append(f"unknown:{node.tag}")
    targets = []
    for spid in target_ids:
        hits = shapes.get(spid, [])
        resolved = hits[0] if len(hits) == 1 else {}
        targets.append({"spid": spid, "name": resolved.get("name"),
                        "key": resolved.get("key"), "resolved": len(hits) == 1,
                        "ambiguous": len(hits) > 1})
    return {"present": bool(timings), "count": len(timings),
            "has_builds": bool(builds), "build_list_count": len(builds),
            "builds": [{"type": _tag(n), "attributes": dict(n.attrib)}
                       for b in builds for n in b.iter() if n is not b and _tag(n).startswith("bld")],
            "target_ids": target_ids, "targets": targets,
            "dangling_ids": [i for i in target_ids if i not in shapes],
            "ambiguous_ids": [i for i in target_ids if len(shapes.get(i, [])) > 1],
            "events": _unique(t["event"] for t in triggers if t["event"]),
            "triggers": triggers, "paragraph_ranges": paragraphs,
            "character_ranges": chars, "sequences": sequences, "effects": effects,
            "behaviors": behaviors, "advanced_features": _unique(features)}


def _transition(root):
    variants = []
    for node in root.iter(f"{{{P}}}transition"):
        types = []
        descendants = list(node.iter())
        for child in _branches(node):
            if _tag(child) not in {"sndAc", "extLst"}:
                types.append(_tag(child))
        # Morph can be an extension nested below extLst.
        if any(_tag(n).lower() == "morph" for n in descendants):
            types.append("morph")
        types = _unique(types) or ["none"]  # An explicit kind-less transition.
        duration = node.get(f"{{{P14}}}dur", node.get("dur"))
        variant = {"type": types[0] if len(types) == 1 else "mixed",
                         "types": types, "duration_ms": _number(duration),
                         "speed": node.get("spd"),
                         "advance_on_click": _bool(node.get("advClick"), True),
                         "advance_after_ms": _number(node.get("advTm")),
                         "automatic": node.get("advTm") is not None,
                         "has_sound": any(_tag(n) in {"sndAc", "stSnd", "endSnd", "snd"} for n in descendants),
                         "is_morph": "morph" in types,
                         "sound_refs": [{"relationship_id": n.get(f"{{{R}}}embed") or n.get(f"{{{R}}}link"),
                                         "name": n.get("name")}
                                        for n in descendants if _tag(n) == "snd"]}
        # power_pptx may retain both a working copy and its original MC choice.
        if variant not in variants:
            variants.append(variant)
    types = _unique(t for v in variants for t in v["types"])
    first = variants[0] if variants else {}
    return {"present": bool(variants), "type": types[0] if len(types) == 1 else ("mixed" if types else None),
            "types": types, "duration_ms": first.get("duration_ms"),
            "speed": first.get("speed"), "advance_on_click": first.get("advance_on_click"),
            "advance_after_ms": first.get("advance_after_ms"),
            "automatic": any(v["automatic"] for v in variants),
            "has_sound": any(v["has_sound"] for v in variants),
            "is_morph": any(v["is_morph"] for v in variants), "variants": variants}


def _record(obj):
    root = obj._element
    shapes = _shape_index(root)
    timing, transition = _timing(root, shapes), _transition(root)
    warnings = []
    if timing["dangling_ids"]:
        warnings.append("dangling_shape_targets")
    if timing["ambiguous_ids"]:
        warnings.append("ambiguous_shape_targets")
    if timing["advanced_features"]:
        warnings.append("advanced_timing_features")
    if any(n.tag == f"{{{MC}}}AlternateContent" for n in root.iter()):
        warnings.append("alternate_content_all_branches_inspected")
    cslide = root.find(f"{{{P}}}cSld")
    return {"part": str(obj.part.partname),
            "name": cslide.get("name", "") if cslide is not None else "",
            "timing": timing, "transition": transition, "warnings": warnings}


def _has_timing(record):
    return record["timing"]["present"] or record["timing"]["has_builds"]


def inventory(prs) -> dict:
    """Inventory source XML only, without modifying prs or inferring playback.

    has_inherited_timing means a linked layout/master contains timing or builds;
    inherited_timing_parts identifies those sources, not effective player state.
    Transitions are also local XML; variants preserve mutually exclusive MC
    branches, and top-level flags mean 'present in any branch'.
    """
    masters, layouts, slides = [], [], []
    for master in prs.slide_masters:
        record = _record(master)
        record["layout_parts"] = [str(lo.part.partname) for lo in master.slide_layouts]
        masters.append(record)
        for layout in master.slide_layouts:
            item = _record(layout)
            item["name"] = layout.name or item["name"]
            item["master_part"] = record["part"]
            layouts.append(item)
    sources = {r["part"]: r for r in masters + layouts}
    for number, slide in enumerate(prs.slides, 1):
        item = _record(slide)
        layout = slide.slide_layout
        item.update(slide=number, layout_part=str(layout.part.partname),
                    master_part=str(layout.slide_master.part.partname))
        inherited = [part for part in (item["layout_part"], item["master_part"])
                     if part in sources and _has_timing(sources[part])]
        item["inherited_timing_parts"] = inherited
        item["has_inherited_timing"] = bool(inherited)
        if inherited:
            item["warnings"].append("inherited_timing_present")
        slides.append(item)
    all_parts = masters + layouts + slides
    summary = {"masters": len(masters), "layouts": len(layouts), "slides": len(slides)}
    for label, records in (("masters", masters), ("layouts", layouts), ("slides", slides)):
        summary[f"{label}_with_timing"] = sum(r["timing"]["present"] for r in records)
        summary[f"{label}_with_builds"] = sum(r["timing"]["has_builds"] for r in records)
        summary[f"{label}_with_transition"] = sum(r["transition"]["present"] for r in records)
    summary.update(
        slides_with_inherited_timing=sum(s["has_inherited_timing"] for s in slides),
        slides_with_morph=sum(s["transition"]["is_morph"] for s in slides),
        slides_with_automatic_transition=sum(s["transition"]["automatic"] for s in slides),
        slides_with_transition_sound=sum(s["transition"]["has_sound"] for s in slides),
        parts_with_advanced_timing=sum(bool(r["timing"]["advanced_features"]) for r in all_parts),
        dangling_target_count=sum(len(r["timing"]["dangling_ids"]) for r in all_parts),
        ambiguous_target_count=sum(len(r["timing"]["ambiguous_ids"]) for r in all_parts),
        warning_count=sum(len(r["warnings"]) for r in all_parts))
    return {"masters": masters, "layouts": layouts, "slides": slides,
            "summary": summary, "inspection": "static_xml_only", "playback_verified": False}


if __name__ == "__main__":
    import argparse
    import json
    from power_pptx import Presentation

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx")
    args = parser.parse_args()
    print(json.dumps(inventory(Presentation(args.pptx)), ensure_ascii=False, indent=2))
