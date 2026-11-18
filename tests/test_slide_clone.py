"""Regression for owned tags, relationship type collisions and independent notes."""
import sys
from pathlib import Path
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]/'scripts'))
from lxml import etree
from power_pptx import Presentation
from power_pptx.opc.package import Part
from power_pptx.opc.packuri import PackURI
from power_pptx.util import Inches
from pptx_package import P, R, package_problems, repair_package
from slide_clone import clone_slide, remove_original_slides


def tagged_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    sh = s.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
    sh.text = 'Original content'
    part = Part.load(PackURI('/ppt/tags/tag1.xml'),
                     'application/vnd.openxmlformats-officedocument.presentationml.tags+xml',
                     prs.part.package, f'<p:tagLst xmlns:p="{P}"><p:tag name="vendor" val="data"/></p:tagLst>'.encode())
    rid = s.part.rels.get_or_add(R+'/tags', part)
    data = etree.SubElement(sh._element.find('.//{'+P+'}nvPr'), '{'+P+'}custDataLst')
    etree.SubElement(data, '{'+P+'}tags', {'{'+R+'}id': rid})
    s.notes_slide.notes_text_frame.text = 'Original notes'
    return s


class CloneRegression(unittest.TestCase):
    def test_two_clones_have_independent_tags_notes_and_reopenable_shapes(self):
        with tempfile.TemporaryDirectory() as d:
            p = Presentation(); origin = tagged_slide(p)
            clones = [clone_slide(p, origin) for _ in range(2)]
            clones[0].shapes[0].text = 'First edit'
            clones[0].notes_slide.notes_text_frame.text = 'First notes'
            self.assertEqual(origin.shapes[0].text, 'Original content')
            self.assertEqual(clones[1].shapes[0].text, 'Original content')
            remove_original_slides(p, 1)
            f = Path(d)/'out.pptx'; p.save(f); repair_package(f)
            self.assertEqual(package_problems(f), [])
            r = Presentation(f)
            self.assertEqual([s.shapes[0].text for s in r.slides], ['First edit', 'Original content'])
            self.assertEqual([s.notes_slide.notes_text_frame.text for s in r.slides], ['First notes', 'Original notes'])
            tag_parts = [next(rel.target_part for rel in s.part.rels.values() if rel.reltype == R+'/tags') for s in r.slides]
            self.assertNotEqual(tag_parts[0].partname, tag_parts[1].partname)
            self.assertEqual(tag_parts[0].blob, tag_parts[1].blob)
            from power_pptx._schema import iter_schema_violations
            self.assertEqual(list(iter_schema_violations(f)), [])

    def test_isolated_body_preserves_structural_layout_and_reopens(self):
        from slide_clone import isolate_layout
        from power_pptx._schema import iter_schema_violations
        with tempfile.TemporaryDirectory() as d:
            p = Presentation(); structural = tagged_slide(p)
            body = clone_slide(p, structural)
            master = structural.slide_layout.slide_master
            old_layout = structural.slide_layout.part
            old_master = master.part
            for number,owner in enumerate((old_layout,old_master),2):
                tags=Part.load(PackURI(f'/ppt/tags/tag{number}.xml'),
                    'application/vnd.openxmlformats-officedocument.presentationml.tags+xml',
                    p.part.package,f'<p:tagLst xmlns:p="{P}"><p:tag name="test" val="data"/></p:tagLst>'.encode())
                owner.rels.get_or_add(R+'/tags',tags)
            before = (structural.slide_layout._element.xml, master._element.xml)
            cache = {}
            isolate_layout(p, body, cache)
            self.assertIsNot(body.slide_layout.part, old_layout)
            self.assertIsNot(body.slide_layout.slide_master.part, old_master)
            for owner in (body.slide_layout, body.slide_layout.slide_master):
                for shape in list(owner.shapes):
                    shape._element.getparent().remove(shape._element)
            self.assertEqual(before, (structural.slide_layout._element.xml, master._element.xml))
            another = clone_slide(p, structural)
            isolate_layout(p, another, cache)
            self.assertIs(body.slide_layout.part, another.slide_layout.part)
            f = Path(d)/'isolated.pptx'; p.save(f); repair_package(f)
            self.assertEqual(package_problems(f), [])
            self.assertEqual(list(iter_schema_violations(f)), [])
            reopened = Presentation(f)
            self.assertEqual(len(reopened.slides), 3)
            self.assertNotEqual(reopened.slides[0].slide_layout.part.partname,
                                reopened.slides[1].slide_layout.part.partname)
            self.assertEqual(len(reopened.slides[1].slide_layout.slide_master.shapes), 0)
            self.assertGreater(len(reopened.slides[0].slide_layout.slide_master.shapes), 0)

    def test_prepared_edits_keep_nested_clone_sources_and_original_ops(self):
        from slide_clone import prepared_edit_ops
        ops = [{'command':'add','parent':'/','from':'/slide[2]'},
               {'command':'set','path':'/slide[4]/group[1]/shape[2]','props':{'text':'kept'}},
               {'command':'add','parent':'/slide[4]','from':'/slide[4]/group[1]'},
               {'command':'remove','path':'/slide[3]'}]
        edits=prepared_edit_ops(ops,3)
        self.assertEqual(len(edits),2)
        self.assertEqual(edits[0]['path'],'/slide[1]/group[1]/shape[2]')
        self.assertEqual(edits[1]['from'],'/slide[1]/group[1]')
        self.assertEqual(ops[1]['path'],'/slide[4]/group[1]/shape[2]')
        with self.assertRaises(ValueError):
            prepared_edit_ops([{'command':'set','path':'/slide[1]/shape[1]'}],3)

    def test_clone_geometry_ignores_new_metadata_but_detects_movement(self):
        from template_contract import fingerprint
        p=Presentation();origin=tagged_slide(p);clone=clone_slide(p,origin)
        self.assertEqual(fingerprint(origin),fingerprint(clone))
        clone.shapes[0].left+=1000
        self.assertNotEqual(fingerprint(origin),fingerprint(clone))

    def test_present_but_wrong_type_rid_is_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p = Presentation(); s = tagged_slide(p)
            nrid = next(r.rId for r in s.part.rels.values() if r.reltype == R+'/notesSlide')
            s._element.find('.//{'+P+'}tags').set('{'+R+'}id', nrid)
            f = Path(d)/'wrong.pptx'; p.save(f)
            self.assertTrue(any('关系类型错误' in issue for issue in package_problems(f)))

    def test_shared_tag_part_is_rejected_even_when_ids_and_targets_exist(self):
        with tempfile.TemporaryDirectory() as d:
            p = Presentation(); origin = tagged_slide(p); clone = clone_slide(p, origin)
            old = next(r for r in clone.part.rels.values() if r.reltype == R+'/tags')
            target = next(r.target_part for r in origin.part.rels.values() if r.reltype == R+'/tags')
            clone.part.drop_rel(old.rId)
            rid = clone.part.rels.get_or_add(R+'/tags', target)
            clone._element.find('.//{'+P+'}tags').set('{'+R+'}id', rid)
            f = Path(d)/'shared.pptx'; p.save(f)
            self.assertTrue(any('tags 部件被多个源共享' in issue for issue in package_problems(f)))

if __name__ == '__main__':
    unittest.main()
