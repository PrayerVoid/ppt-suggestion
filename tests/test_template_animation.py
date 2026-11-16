"""Animation policies must fail before edits, and detect valid-but-wrong targets."""
import contextlib
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from power_pptx import Presentation
from power_pptx.util import Inches
import template_fill
from template_animation import (AnimationError, apply_plan, build_plan, verify_plan,
                                _rewrite)
from pptx_package import P, PackageError, package_problems, repair_package


def timing(owner, target, *, build="allAtOnce", paragraph=False):
    text = '<p:txEl><p:pRg st="0" end="1"/></p:txEl>' if paragraph else ""
    node = etree.fromstring((f'<p:timing xmlns:p="{P}"><p:tnLst><p:par>'
        '<p:cTn id="1" nodeType="tmRoot"><p:childTnLst><p:seq>'
        '<p:cTn id="2" nodeType="mainSeq"><p:childTnLst><p:animEffect filter="fade" transition="in">'
        '<p:cBhvr><p:cTn id="3" dur="500"/><p:tgtEl>'
        f'<p:spTgt spid="{target}">{text}</p:spTgt></p:tgtEl></p:cBhvr>'
        '</p:animEffect></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
        f'<p:bldLst><p:bldP spid="{target}" grpId="0" build="{build}"/></p:bldLst></p:timing>').encode())
    owner._element.append(node)
    return node


def transition(owner, *, morph=False):
    p159 = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"
    owner._element.append(etree.fromstring((f'<p:transition xmlns:p="{P}" xmlns:p159="{p159}" advTm="7000">'
        + ('<p159:morph/>' if morph else '<p:fade/>') + '</p:transition>').encode()))


class AnimationPolicyTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "source.pptx"
        self.output = Path(self.temp.name) / "out.pptx"
        self.prs = Presentation()
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.a = self.slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
        self.a.name, self.a.text = "标题", "原内容"
        self.b = self.slide.shapes.add_textbox(Inches(1), Inches(3), Inches(3), Inches(1))
        self.b.name, self.b.text = "正文", "原正文"

    def plan(self, items=None, policy="auto"):
        self.prs.save(self.path)
        self.output.write_bytes(self.path.read_bytes())
        return build_plan(self.path, items or [{"template_slide": 1}],
                          template_fill.scan_template(str(self.path)), policy)

    def test_text_edit_preserves_whole_shape_animation_but_resets_auto_advance(self):
        timing(self.slide, self.a.shape_id)
        transition(self.slide)
        plan = self.plan([{"template_slide": 1, "fields": {"标题": "新内容"}}])
        self.assertEqual(plan["pages"][0]["timing"], "preserve")
        self.assertEqual(plan["pages"][0]["transition"], "manual")
        apply_plan(self.output, plan)
        self.assertEqual(verify_plan(self.output, plan), [])
        with ZipFile(self.output) as package:
            root = etree.fromstring(package.read("ppt/slides/slide1.xml"))
            node = root.find(f"{{{P}}}transition")
            self.assertIsNone(node.get("advTm"))
            self.assertEqual(node.get("advClick"), "1")
        self.assertNotEqual(self.output.read_bytes(), self.path.read_bytes())

    def test_id_remapping_passes_but_wrong_existing_target_fails(self):
        timing(self.slide, self.a.shape_id)
        plan = self.plan(policy="preserve")
        prs = Presentation(self.output)
        root = prs.slides[0]._element
        for node in root.iter(f"{{{P}}}cNvPr"):
            if node.get("id") == str(self.a.shape_id):
                node.set("id", "200")
        for node in root.iter():
            if node.get("spid") == str(self.a.shape_id):
                node.set("spid", "200")
        prs.save(self.output)
        self.assertEqual(verify_plan(self.output, plan), [])
        for node in root.iter():
            if node.get("spid") == "200":
                node.set("spid", str(self.b.shape_id))
        prs.save(self.output)
        repair_package(self.output, repair_animations=False)
        self.assertEqual(package_problems(self.output), [])
        self.assertTrue(verify_plan(self.output, plan))

    def test_deletion_of_animated_shape_is_planned_before_editing(self):
        timing(self.slide, self.a.shape_id)
        item = {"template_slide": 1, "remove": ["标题"]}
        plan = self.plan([item])
        self.assertEqual(plan["pages"][0]["timing"], "strip")
        apply_plan(self.output, plan)
        self.assertEqual(verify_plan(self.output, plan), [])
        with self.assertRaises(AnimationError):
            self.plan([item], "preserve")

    def test_paragraph_build_edit_requires_replanning(self):
        self.a.text = "第一段\n第二段"
        timing(self.slide, self.a.shape_id, build="p", paragraph=True)
        item = {"template_slide": 1, "fields": {"标题": "只剩一段"}}
        plan = self.plan([item])
        self.assertEqual(plan["pages"][0]["timing"], "strip")
        with self.assertRaises(AnimationError):
            self.plan([item], "preserve")

    def test_unanimated_deletion_keeps_shifted_animation_target(self):
        timing(self.slide, self.b.shape_id)
        plan = self.plan([{"template_slide": 1, "remove": ["标题"]}], "preserve")
        prs = Presentation(self.output)
        shape = prs.slides[0].shapes[0]._element
        shape.getparent().remove(shape)
        prs.save(self.output)
        self.assertEqual(verify_plan(self.output, plan), [])

    def test_clone_of_animated_page_is_not_assumed_to_clone_effects(self):
        timing(self.slide, self.a.shape_id)
        group = self.slide.shapes.add_group_shape()
        group.name = "卡片"
        plan = self.plan([{"template_slide": 1, "clone_groups": [{"source": "卡片"}]}])
        self.assertEqual(plan["pages"][0]["timing"], "strip")

    def test_morph_reorder_strips_and_preserve_rejects(self):
        transition(self.slide, morph=True)
        plan = self.plan()
        self.assertEqual(plan["pages"][0]["transition"], "strip")
        with self.assertRaises(AnimationError):
            self.plan(policy="preserve")

    def test_shared_timing_edits_require_explicit_consistent_strip(self):
        master = self.slide.slide_layout.slide_master
        target = next(master._element.iter(f"{{{P}}}cNvPr")).get("id")
        # Use a real master shape as the target, not the root shape-tree ID.
        target = master.shapes[0].shape_id
        timing(master, target)
        item = {"template_slide": 1, "fields": {"标题": "新内容"}}
        with self.assertRaises(AnimationError):
            self.plan([item])
        plan = self.plan([item], "strip")
        apply_plan(self.output, plan)
        self.assertEqual(verify_plan(self.output, plan), [])
        with self.assertRaises(AnimationError):
            self.plan([{"template_slide": 1, "animation_policy": "strip"},
                       {"template_slide": 1, "animation_policy": "preserve"}])

    def test_alternate_content_timing_and_transitions_are_all_stripped(self):
        timing(self.slide, self.a.shape_id)
        transition(self.slide)
        plan = self.plan(policy="strip")
        with ZipFile(self.output) as package:
            root = etree.fromstring(package.read("ppt/slides/slide1.xml"))
        mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
        wrapper = etree.SubElement(root, f"{{{mc}}}AlternateContent")
        choice = etree.SubElement(wrapper, f"{{{mc}}}Choice", Requires="p14")
        for tag in ("timing", "transition"):
            node = root.find(f"{{{P}}}{tag}")
            root.remove(node)
            choice.append(node)
        _rewrite(self.output, {"ppt/slides/slide1.xml": etree.tostring(root)})
        apply_plan(self.output, plan)
        with ZipFile(self.output) as package:
            xml = package.read("ppt/slides/slide1.xml")
            self.assertNotIn(b"p:timing", xml)
            self.assertNotIn(b"p:transition", xml)
            self.assertNotIn(b"AlternateContent", xml)
        self.assertEqual(verify_plan(self.output, plan), [])

    def test_no_silent_animation_repair_in_strict_generation_path(self):
        timing(self.slide, "999")
        self.plan(policy="strip")
        before = self.output.read_bytes()
        with self.assertRaises(PackageError):
            repair_package(self.output, repair_animations=False)
        self.assertEqual(before, self.output.read_bytes())

    def test_time_dependency_and_inherited_dangling_targets_are_hard(self):
        node = timing(self.slide, self.a.shape_id)
        condition = etree.SubElement(node, f"{{{P}}}tn", val="999")
        timing(self.slide.slide_layout, "999")
        self.prs.save(self.output)
        errors = package_problems(self.output)
        self.assertTrue(any("时间依赖" in err for err in errors))
        self.assertTrue(any("slideLayout" in err and "悬空" in err for err in errors))
        node.remove(condition)

    def test_preserve_rejection_precedes_officecli_and_keeps_previous_output(self):
        timing(self.slide, self.a.shape_id)
        self.prs.save(self.path)
        self.output.write_bytes(b"previous good output")
        mapping = Path(self.temp.name) / "map.json"
        deck = Path(self.temp.name) / "deck.json"
        mapping.write_text(json.dumps({"slides": [{"template_slide": 1, "remove": ["标题"]}]}))
        deck.write_text(json.dumps({"slides": [{"title": "新内容"}]}))
        argv = ["template_fill.py", "--template", str(self.path), "--map", str(mapping),
                "--deck", str(deck), "--output", str(self.output), "--animation-policy", "preserve"]
        with patch.object(sys, "argv", argv), patch.object(template_fill.ocli, "batch") as batch, \
                contextlib.redirect_stderr(io.StringIO()):
            self.assertEqual(template_fill.main(), 2)
        batch.assert_not_called()
        self.assertEqual(self.output.read_bytes(), b"previous good output")


if __name__ == "__main__":
    unittest.main()
