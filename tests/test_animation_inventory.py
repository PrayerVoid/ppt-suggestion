"""Static animation inventory, including inheritance and MC alternatives."""
import io
import json
import sys
import unittest
import zipfile
from pathlib import Path

from lxml import etree
from power_pptx import Presentation
from power_pptx.util import Inches

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from animation_inventory import inventory, P, MC, P14, R

P159 = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"


def xml(text):
    return etree.fromstring((f'<wrapper xmlns:p="{P}" xmlns:mc="{MC}" '
                             f'xmlns:p14="{P14}" xmlns:p159="{P159}" '
                             f'xmlns:r="{R}" xmlns:custom="urn:custom">'
                             f'{text}</wrapper>').encode())[0]


def add_timing(obj, spid, extra=""):
    obj._element.append(xml(f'''
        <p:timing><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot">
          <p:childTnLst><p:seq><p:cTn id="2" nodeType="mainSeq">
            <p:childTnLst><p:animEffect filter="fade">
              <p:cBhvr><p:cTn id="3" presetClass="entr" nodeType="clickEffect"/>
                <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
              </p:cBhvr>
            </p:animEffect></p:childTnLst>
          </p:cTn><p:nextCondLst><p:cond evt="onNext" delay="0">
            <p:tgtEl><p:sldTgt/></p:tgtEl>
          </p:cond></p:nextCondLst></p:seq></p:childTnLst>
        </p:cTn></p:par></p:tnLst>{extra}
        <p:bldLst><p:bldP spid="{spid}" grpId="0"/></p:bldLst></p:timing>'''))


class AnimationInventoryTests(unittest.TestCase):
    def setUp(self):
        self.prs = Presentation()
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

    def textbox(self, parent=None, name="内容"):
        if parent is None:
            parent = self.slide
        shape = parent.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
        shape.name = name
        shape.text = "第一段\n第二段\n第三段"
        return shape

    def test_no_animation_is_json_serializable_and_does_not_mutate(self):
        self.textbox()
        parts = [self.slide, *self.prs.slide_masters, *self.prs.slide_layouts]
        before = [etree.tostring(p._element) for p in parts]
        report = inventory(self.prs)
        self.assertEqual(before, [etree.tostring(p._element) for p in parts])
        self.assertEqual(json.loads(json.dumps(report)), report)
        self.assertFalse(report["playback_verified"])
        self.assertEqual(report["inspection"], "static_xml_only")
        self.assertEqual(report["summary"]["slides_with_timing"], 0)
        self.assertEqual(report["summary"]["warning_count"], 0)
        page = report["slides"][0]
        self.assertEqual(page["slide"], 1)
        self.assertEqual(page["part"], str(self.slide.part.partname))
        self.assertEqual(page["layout_part"], str(self.slide.slide_layout.part.partname))
        self.assertFalse(page["timing"]["present"])
        self.assertEqual(page["timing"]["target_ids"], [])
        self.assertFalse(page["transition"]["present"])
        self.assertFalse(page["has_inherited_timing"])

    def test_targets_builds_sequences_and_group_keys(self):
        group = self.slide.shapes.add_group_shape()
        group.name = "卡片"
        first = self.textbox(group, "正文")
        second = self.textbox(group, "正文")
        add_timing(self.slide, first.shape_id)
        self.slide._element.append(xml(f'<p:bldLst><p:bldP spid="{second.shape_id}"/></p:bldLst>'))
        timing = inventory(self.prs)["slides"][0]["timing"]
        self.assertEqual(timing["target_ids"], [str(first.shape_id), str(second.shape_id)])
        self.assertEqual([t["key"] for t in timing["targets"]], ["卡片/正文#1", "卡片/正文#2"])
        self.assertTrue(all(t["resolved"] for t in timing["targets"]))
        self.assertEqual(timing["dangling_ids"], [])
        self.assertEqual(timing["build_list_count"], 2)
        self.assertEqual(timing["sequences"]["main"][0]["id"], "2")
        self.assertEqual(timing["effects"][0]["nodeType"], "clickEffect")
        self.assertEqual(timing["events"], ["onNext"])
        self.assertEqual(timing["advanced_features"], [])

    def test_master_layout_inheritance_uses_separate_id_domains(self):
        layout = self.slide.slide_layout
        master = layout.slide_master
        # Use an ID which is present only in the master. It must still dangle
        # in the slide, even though that slide inherits the master timing.
        cnv = next(master._element.iter(f"{{{P}}}cNvPr"))
        cnv.set("id", "777")
        add_timing(master, "777")
        layout_id = next(layout._element.iter(f"{{{P}}}cNvPr")).get("id")
        add_timing(layout, layout_id)
        add_timing(self.slide, "777")
        # An unused layout is still inventoried.
        unused = self.prs.slide_layouts[0]
        add_timing(unused, "999")
        report = inventory(self.prs)
        self.assertEqual(len(report["layouts"]), len(self.prs.slide_layouts))
        self.assertEqual(report["masters"][0]["timing"]["dangling_ids"], [])
        own_layout = next(r for r in report["layouts"] if r["part"] == str(layout.part.partname))
        self.assertEqual(own_layout["timing"]["dangling_ids"], [])
        self.assertEqual(own_layout["master_part"], str(master.part.partname))
        page = report["slides"][0]
        self.assertTrue(page["has_inherited_timing"])
        self.assertEqual(page["timing"]["dangling_ids"], ["777"])
        self.assertEqual(page["inherited_timing_parts"], [str(layout.part.partname), str(master.part.partname)])
        self.assertEqual(report["summary"]["layouts_with_timing"], 2)
        self.assertEqual(report["summary"]["dangling_target_count"], 2)
        # Inherited presence is reported even on a slide without local timing.
        second = self.prs.slides.add_slide(layout)
        page = inventory(self.prs)["slides"][1]
        self.assertEqual(page["part"], str(second.part.partname))
        self.assertFalse(page["timing"]["present"])
        self.assertTrue(page["has_inherited_timing"])

    def test_alternate_content_timing_builds_paragraphs_and_events(self):
        shape = self.textbox()
        self.slide._element.append(xml(f'''
          <mc:AlternateContent><mc:Choice Requires="p14"><p:timing>
            <p:tnLst><p:seq><p:cTn id="5" nodeType="interactiveSeq">
              <p:stCondLst><p:cond evt="onClick" delay="0">
                <p:tgtEl><p:spTgt spid="{shape.shape_id}"/></p:tgtEl>
              </p:cond></p:stCondLst><p:childTnLst><p:anim>
                <p:cBhvr><p:tgtEl><p:spTgt spid="{shape.shape_id}">
                  <p:txEl><p:pRg st="0" end="2"/></p:txEl>
                </p:spTgt></p:tgtEl></p:cBhvr>
              </p:anim></p:childTnLst>
            </p:cTn></p:seq></p:tnLst>
            <mc:AlternateContent><mc:Choice Requires="p14"><p:bldLst>
              <p:bldP spid="{shape.shape_id}" build="p"/>
            </p:bldLst></mc:Choice><mc:Fallback/></mc:AlternateContent>
          </p:timing></mc:Choice><mc:Fallback><p:bldLst>
            <p:bldP spid="998"/>
          </p:bldLst></mc:Fallback></mc:AlternateContent>'''))
        page = inventory(self.prs)["slides"][0]
        timing = page["timing"]
        self.assertTrue(timing["present"])
        self.assertEqual(timing["build_list_count"], 2)
        self.assertEqual(timing["dangling_ids"], ["998"])
        self.assertEqual(timing["paragraph_ranges"], [{"spid": str(shape.shape_id), "start": 0, "end": 2}])
        self.assertEqual(timing["events"], ["onClick"])
        self.assertEqual(timing["triggers"][0]["target_ids"], [str(shape.shape_id)])
        self.assertEqual(len(timing["sequences"]["interactive"]), 1)
        self.assertIn("alternate_content", timing["advanced_features"])
        self.assertIn("alternate_content_all_branches_inspected", page["warnings"])

    def test_morph_automatic_duration_and_sound_in_mc_branches(self):
        self.slide._element.append(xml('''
          <mc:AlternateContent><mc:Choice Requires="p159">
            <p:transition p14:dur="1250" advClick="0" advTm="3500">
              <p159:morph option="byObject"/>
              <p:sndAc><p:stSnd loop="1"><p:snd r:embed="rIdSound" name="铃声"/></p:stSnd></p:sndAc>
            </p:transition>
          </mc:Choice><mc:Fallback>
            <p:transition spd="slow"><p:fade/></p:transition>
          </mc:Fallback></mc:AlternateContent>'''))
        # Store the exact authoring XML in the test package. power_pptx's save
        # normalizer treats a manually injected MC wrapper as stale without a
        # corresponding direct working copy, so save alone would erase it.
        source_xml = etree.tostring(self.slide._element)
        stream, package = io.BytesIO(), io.BytesIO()
        self.prs.save(stream)
        stream.seek(0)
        with zipfile.ZipFile(stream) as src, zipfile.ZipFile(package, "w") as dst:
            for entry in src.infolist():
                data = source_xml if entry.filename == "ppt/slides/slide1.xml" else src.read(entry)
                dst.writestr(entry, data)
        package.seek(0)
        loaded = Presentation(package)
        before = etree.tostring(loaded.slides[0]._element)
        report = inventory(loaded)
        self.assertEqual(etree.tostring(loaded.slides[0]._element), before)
        transition = report["slides"][0]["transition"]
        self.assertEqual(transition["types"], ["morph", "fade"])
        self.assertEqual(len(transition["variants"]), 2)
        self.assertEqual(transition["type"], "mixed")
        self.assertEqual(transition["duration_ms"], 1250)
        self.assertEqual(transition["advance_after_ms"], 3500)
        self.assertFalse(transition["advance_on_click"])
        self.assertTrue(transition["automatic"])
        self.assertTrue(transition["has_sound"])
        self.assertTrue(transition["is_morph"])
        self.assertEqual(transition["variants"][0]["sound_refs"][0]["relationship_id"], "rIdSound")
        self.assertFalse(transition["variants"][1]["automatic"])
        self.assertEqual(report["summary"]["slides_with_morph"], 1)
        self.assertEqual(report["summary"]["slides_with_automatic_transition"], 1)
        self.assertEqual(report["summary"]["slides_with_transition_sound"], 1)

    def test_advanced_features_and_build_only_inheritance(self):
        shape = self.textbox()
        add_timing(self.slide, shape.shape_id, f'''
          <p:tnLst><p:animMotion path="M 0 0 L 1 1"><p:cBhvr><p:tgtEl>
            <p:spTgt spid="{shape.shape_id}"/>
          </p:tgtEl></p:cBhvr></p:animMotion><p:audio><p:cMediaNode>
            <p:tgtEl><p:sndTgt r:embed="rIdAudio"/></p:tgtEl>
          </p:cMediaNode></p:audio><p:video/><p:cmd cmd="play"/>
          <custom:futureEffect/></p:tnLst>''')
        layout = self.slide.slide_layout
        layout._element.append(xml('<p:bldLst><p:bldP spid="1"/></p:bldLst>'))
        page = inventory(self.prs)["slides"][0]
        for feature in ("motion", "media", "command", "unknown:{urn:custom}futureEffect"):
            self.assertIn(feature, page["timing"]["advanced_features"])
        self.assertTrue(page["has_inherited_timing"])
        self.assertFalse(page["transition"]["has_sound"])

    def test_duplicate_local_id_is_ambiguous_not_dangling(self):
        one = self.textbox(name="一")
        two = self.textbox(name="二")
        next(two._element.iter(f"{{{P}}}cNvPr")).set("id", str(one.shape_id))
        add_timing(self.slide, one.shape_id)
        timing = inventory(self.prs)["slides"][0]["timing"]
        self.assertEqual(timing["ambiguous_ids"], [str(one.shape_id)])
        self.assertEqual(timing["dangling_ids"], [])
        self.assertIsNone(timing["targets"][0]["key"])
        self.assertFalse(timing["targets"][0]["resolved"])


if __name__ == "__main__":
    unittest.main()
