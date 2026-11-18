import unittest
import tempfile
from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from power_pptx import Presentation
from power_pptx.util import Inches
from template_canvas import clean, layers, inject, generate
from template_contract import fingerprint,validate_map_mode,density_evidence

class CanvasContractTests(unittest.TestCase):
    def setup_slide(self):
        p=Presentation();s=p.slides.add_slide(p.slide_layouts[6]);s.shapes.add_textbox(Inches(1),Inches(1),Inches(4),Inches(3)).name='OldBody'
        spec={'name':'body','template_slide':1,'content_region':[500000,500000,7000000,5000000],
              'style':{'font':'Arial','accent':'1E5550','text':'404040'},'layers':{}}
        for k,owner in layers(s).items():spec['layers'][k]={'keep':[],'remove':[sh.name for sh in owner.shapes],'keep_reasons':{}}
        return p,s,spec
    def test_body_and_inherited_content_removed(self):
        p,s,spec=self.setup_slide();e=clean(s,spec)
        self.assertFalse(list(s.shapes));self.assertFalse(list(s.slide_layout.shapes));self.assertFalse(list(s.slide_layout.slide_master.shapes));self.assertIn('master',{r['layer'] for r in e})
    def test_unclassified_object_rejected(self):
        p,s,spec=self.setup_slide();spec['layers']['slide']['remove']=[]
        with self.assertRaises(ValueError):clean(s,spec)
    def test_conflicting_keep_remove_rejected(self):
        p,s,spec=self.setup_slide();spec['layers']['slide']['keep']=['OldBody']
        with self.assertRaises(ValueError):clean(s,spec)
    def test_preserve_needs_reason(self):
        p,s,spec=self.setup_slide();spec['layers']['slide'].update(keep=['OldBody'],remove=[])
        with self.assertRaises(ValueError):clean(s,spec)
    def test_intruding_decoration_needs_review(self):
        p,s,spec=self.setup_slide();spec['layers']['slide'].update(keep=['OldBody'],remove=[],keep_reasons={'OldBody':'decoration'})
        with self.assertRaises(ValueError):clean(s,spec)
    def test_explicit_notes_required(self):
        p,s,spec=self.setup_slide();clean(s,spec)
        with self.assertRaises(ValueError):inject(s,spec,{'items':[],'source':'doc'})
    def test_capacity_not_shrunk(self):
        p,s,spec=self.setup_slide();clean(s,spec)
        with self.assertRaises(ValueError):inject(s,spec,{'notes':'talk','source':'doc','title':'Title','items':[{'title':'A','body':'word '*5000}]*2})
    def test_default_notes_exclude_production_sources(self):
        from speaker_notes import generated_script
        script=generated_script({'title':'协议边界','blocks':[{'kind':'paragraph','text':'请求与响应分别加密。'}],'source':'参考资料.html 第三部分','source_mapping':{'text':'内部制作溯源'}})
        self.assertIn('请求与响应',script)
        self.assertNotIn('参考资料',script)
        self.assertNotIn('内部制作溯源',script)
    def test_quote_only_renders_explicit_external_attribution(self):
        import render_pptx
        p,s,spec=self.setup_slide()
        render_pptx._quote(s,1,1,8,{'text':'真实引文','source':'参考资料.html'}, {})
        visible='\n'.join(sh.text for sh in s.shapes if sh.has_text_frame)
        self.assertNotIn('参考资料.html',visible)
        render_pptx._quote(s,1,3,8,{'text':'另一引文','external_citation':'作者，《著作》，第12页'}, {})
        self.assertIn('作者，《著作》，第12页','\n'.join(sh.text for sh in s.shapes if sh.has_text_frame))
    def test_notes_opt_out_preserves_template_and_default_preserves_labeled_notes(self):
        with tempfile.TemporaryDirectory() as d:
            p,s,spec=self.setup_slide();s.notes_slide.notes_text_frame.text='Template remarks'
            source=Path(d)/'source.pptx';p.save(source)
            page={'mother':'body','title':'A real mechanism','source':'internal mapping',
                  'items':[{'title':'A','body':'Evidence and conditions'},{'title':'B','body':'Limits and assumptions'}]}
            out=Path(d)/'out.pptx'
            generate(source,{'mothers':[spec]},out,{'notes_enabled':False,'slides':[page]})
            self.assertEqual(Presentation(out).slides[0].notes_slide.notes_text_frame.text,'Template remarks')
            page['notes']='A natural speaker script.'
            generate(source,{'mothers':[spec]},out,{'slides':[page]})
            notes=Presentation(out).slides[0].notes_slide.notes_text_frame.text
            self.assertIn('【本次讲稿】\nA natural speaker script.',notes)
            self.assertIn('【模板原备注（供参考）】\nTemplate remarks',notes)

    def test_mixed_structure_and_core_body_preserve_geometry_and_inheritance(self):
        from pptx_package import package_problems
        from power_pptx._schema import iter_schema_violations
        from template_animation import verify_plan
        import json
        with tempfile.TemporaryDirectory() as d:
            p,s,spec=self.setup_slide(); source=Path(d)/'source.pptx'; p.save(source)
            original=fingerprint(s)
            structure={'template_slide':1,'fields':{'OldBody':'Opening context'},
                       'title':'Opening mechanism','source':'evidence','notes':'Opening speech.'}
            body={'mother':'body','title':'Protocol mechanism','source':'evidence','notes':'Body speech.',
                  'blocks':[{'kind':'table','headers':['Mechanism','Boundary'],
                             'rows':[['Encryption','Client to endpoint'],['Trusted compute','Attested runtime']]}]}
            out=Path(d)/'mixed.pptx'
            generate(source,{'mothers':[spec]},out,{'slides':[structure,body]})
            result=Presentation(out)
            self.assertEqual(fingerprint(result.slides[0]),original)
            self.assertEqual(result.slides[0].shapes[0].text,'Opening context')
            self.assertNotEqual(result.slides[0].slide_layout.part.partname,result.slides[1].slide_layout.part.partname)
            self.assertGreater(len(result.slides[0].slide_layout.slide_master.shapes),0)
            self.assertEqual(len(result.slides[1].slide_layout.slide_master.shapes),0)
            self.assertTrue(any(sh.has_table for sh in result.slides[1].shapes))
            self.assertEqual(package_problems(out),[])
            self.assertEqual(list(iter_schema_violations(out)),[])
            self.assertEqual(verify_plan(out,json.loads(Path(str(out)+'.animation-plan.json').read_text())),[])

    def test_native_core_region_keeps_table_fonts_and_restores_style(self):
        import render_pptx as renderer
        p,s,spec=self.setup_slide(); clean(s,spec)
        previous=(renderer.FONT,renderer.ACCENT,renderer.CONTENT_BOTTOM)
        page={'title':'A real mechanism','source':'evidence','notes':'A natural explanation.',
              'blocks':[{'kind':'table','headers':['A','B'],'rows':[['Input','Encrypted']]}]}
        inject(s,spec,page)
        self.assertEqual((renderer.FONT,renderer.ACCENT,renderer.CONTENT_BOTTOM),previous)
        table=next(sh for sh in s.shapes if sh.has_table)
        self.assertGreaterEqual(table.left,spec['content_region'][0])
        self.assertEqual(table.table.cell(1,0).text_frame.paragraphs[0].runs[0].font.size.pt,12)
        self.assertEqual(table.table.cell(1,0).text_frame.paragraphs[0].runs[0].font.name,'Arial')
        page['blocks']=[{'kind':'unknown'}]
        with self.assertRaises(ValueError): inject(s,spec,page)
        self.assertEqual((renderer.FONT,renderer.ACCENT,renderer.CONTENT_BOTTOM),previous)

    def test_strict_object_additions_require_custom(self):
        for key in ('extra_shapes','formulas'):
            with self.subTest(key=key),self.assertRaises(ValueError):
                validate_map_mode({'slides':[{key:[{}]}]})
        with self.assertRaises(ValueError):
            validate_map_mode({'slides':[{'fields':{'Picture':{'chart':{}}}}]})
        self.assertEqual(validate_map_mode({'mode':'custom','layout_change_plan':'audit whole layout','slides':[{'extra_shapes':[{}]}]}),'custom')

    def test_strict_delete_rejected(self):
        with self.assertRaises(ValueError):validate_map_mode({'slides':[{'remove':['card']} ]})
    def test_strict_clone_rejected(self):
        with self.assertRaises(ValueError):validate_map_mode({'slides':[{'clone_groups':[{}]}]})
    def test_custom_requires_whole_layout_plan(self):
        with self.assertRaises(ValueError):validate_map_mode({'mode':'custom','slides':[]})
    def test_density_slot_mismatch_rejected(self):
        with self.assertRaises(ValueError):density_evidence({'items':[{'title':'real','body':'evidence'}],'notes':'n','source':'s'}, {}, 4)
    def test_geometry_fingerprint_detects_connector_or_container_change(self):
        p,s,spec=self.setup_slide();before=fingerprint(s);s.shapes[0].left+=1000;self.assertNotEqual(before,fingerprint(s))
    def test_duplicate_slide_creation_identity_is_blocked(self):
        from lxml import etree
        from pptx_package import package_problems
        with tempfile.TemporaryDirectory() as d:
            p=Presentation()
            ns='http://schemas.openxmlformats.org/presentationml/2006/main'
            for _ in range(2):
                s=p.slides.add_slide(p.slide_layouts[6]);cs=s._element.find('{'+ns+'}cSld')
                ext=etree.SubElement(cs,'{'+ns+'}extLst');child=etree.SubElement(ext,'{'+ns+'}ext',uri='{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}')
                etree.SubElement(child,'{http://schemas.microsoft.com/office/powerpoint/2010/main}creationId',val='123')
            f=Path(d)/'duplicate.pptx';p.save(f)
            self.assertTrue(any('slide creationId 123' in issue for issue in package_problems(f)))
    def test_multiple_output_slides_keep_unique_creation_ids_and_notes(self):
        from lxml import etree
        with tempfile.TemporaryDirectory() as d:
            p,s,spec=self.setup_slide();cs=s._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}cSld');ext=etree.SubElement(cs,'{http://schemas.openxmlformats.org/presentationml/2006/main}extLst');ex=etree.SubElement(ext,'{http://schemas.openxmlformats.org/presentationml/2006/main}ext',uri='{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}');etree.SubElement(ex,'{http://schemas.microsoft.com/office/powerpoint/2010/main}creationId',val='123')
            source=Path(d)/'source.pptx';out=Path(d)/'out.pptx';p.save(source)
            page={'mother':'body','title':'Mechanism','notes':'Explicit speech','source':'source doc','items':[{'title':'A','body':'Conditions and evidence'},{'title':'B','body':'Boundaries and sources'}]}
            generate(source,{'mothers':[spec]},out,{'slides':[page,page]});r=Presentation(out)
            self.assertEqual(len(r.slides),2);ids=[sl._element.xpath('.//*[local-name()="creationId"]')[0].get('val') for sl in r.slides];self.assertEqual(len(set(ids)),2)
            self.assertTrue(all('Explicit speech' in sl.notes_slide.notes_text_frame.text for sl in r.slides))

if __name__=='__main__':unittest.main()
