"""Strict template geometry/slot contract and measurable density evidence."""
import argparse
import json
from pathlib import Path
from power_pptx import Presentation


def fingerprint(slide):
    result=[]
    def visit(shapes,path=()):
        for i,s in enumerate(shapes):
            key=path+(i,)
            # Geometry XML includes nested group transforms, lines, rotations and preset shapes.
            geometry=[]
            for node in s._element:
                if node.tag.rsplit('}',1)[-1] in {'spPr','grpSpPr','nvSpPr','nvCxnSpPr','nvGrpSpPr','nvPicPr','style'}:
                    from copy import deepcopy
                    from lxml import etree
                    normalized = deepcopy(node)
                    # Owned metadata/extension identities are independent on each clone;
                    # they do not change geometry, grouping, locks or placeholder slots.
                    for metadata in normalized.xpath('.//*[local-name()="cNvPr" or local-name()="custDataLst"]'):
                        metadata.getparent().remove(metadata)
                    # Clones receive fresh relationship IDs. Relationship types
                    # and targets are checked by the package validator; the
                    # strict geometry contract compares the shape's geometry
                    # and content, so an rId is identity metadata here.
                    for node in normalized.iter():
                        for attr in list(node.attrib):
                            if attr.startswith('{'+
                                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships}'):
                                node.set(attr, '__RELATIONSHIP_ID__')
                    geometry.append(etree.tostring(normalized, method='c14n', exclusive=True).decode())
            result.append((key,s.name,int(s.shape_type),int(s.left),int(s.top),int(s.width),int(s.height),geometry))
            if s.shape_type==6:visit(s.shapes,key)
    visit(slide.shapes)
    return result


def check_contract(template,output,source_pages):
    t=Presentation(str(template));o=Presentation(str(output))
    if len(o.slides)!=len(source_pages):raise ValueError('strict page count mismatch')
    for i,(slide,page) in enumerate(zip(o.slides,source_pages),1):
        if fingerprint(slide)!=fingerprint(t.slides[page-1]):
            raise ValueError(f'page {i}: strict template geometry/object count changed; choose another page or split content')
    return True


def validate_map_mode(mapping):
    # Older custom maps remain callable only as an explicitly separate custom interface.
    mode=mapping.get('mode','strict')
    if mode not in {'strict','custom'}:raise ValueError('mode must be strict or custom; semi-autonomous uses template_canvas.py')
    if mode=='strict':
        for i,page in enumerate(mapping.get('slides',[]),1):
            if any(page.get(key) for key in ('remove', 'clone_groups', 'add_texts', 'extra_shapes', 'formulas')):
                raise ValueError(f'page {i}: strict forbids deleting/cloning/adding layout objects; fill all semantic slots or select another template page')
            if any(isinstance(value, dict) and 'chart' in value
                   for value in (page.get('fields') or {}).values()):
                raise ValueError(f'page {i}: chart object replacement requires custom mode and layout_change_plan; strict preserves existing objects')
    elif not mapping.get('layout_change_plan'):
        raise ValueError('custom layout changes require layout_change_plan covering containers, connectors, numbering, backplates and recentering; custom is not strict')
    return mode


def density_evidence(page,fields,expected_slots):
    items=page.get('items',[])
    if len(items)!=expected_slots:raise ValueError('semantic item count must equal effective template slot count')
    if not page.get('notes','').strip():raise ValueError('explicit per-page notes are required')
    if not page.get('source','').strip():raise ValueError('source is required for density evidence')
    for item in items:
        if not item.get('title','').strip() or not item.get('body','').strip():
            raise ValueError('empty semantic slot')
    # Counts are evidence, not a substitute for semantic review or an arbitrary word quota.
    return {'title':page['title'],'body_characters':sum(len(x['body']) for x in items),
            'expected_slots':expected_slots,'effective_slots':len(items),'slot_coverage':len(items)/expected_slots,
            'filled_fields':len(fields),'source':page['source'],'explicit_notes':True,
            'semantic_review_required':['evidence','conditions','mechanisms','boundaries','source'],
            'content_coverage':[{'item':x['title'],'body_characters':len(x['body'])} for x in items]}


def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--template',required=True);ap.add_argument('--output',required=True);ap.add_argument('--map',required=True);a=ap.parse_args()
    m=json.loads(Path(a.map).read_text());validate_map_mode(m);check_contract(a.template,a.output,[s['template_slide'] for s in m['slides']]);print('Strict geometry and object count: PASS')

if __name__=='__main__':main()
