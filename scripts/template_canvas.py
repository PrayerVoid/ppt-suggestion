"""Audited clean canvas and autonomous body injection; no template card cloning.

python template_canvas.py --template T.pptx --manifest M.json --output mothers.pptx
python template_canvas.py --template T.pptx --manifest M.json --deck D.json --output final.pptx
Every slide/layout/master top-level shape must be explicitly kept or removed.
Shared inherited layers must use identical decisions. Backgrounds remain inherited.
"""
import argparse
import json
import os
from pathlib import Path
import tempfile
from power_pptx import Presentation
from power_pptx.dml.color import RGBColor
from pptx_package import package_problems, repair_package
from template_capacity import check_box


def layers(slide):
    return {'slide': slide, 'layout': slide.slide_layout,
            'master': slide.slide_layout.slide_master}


def clean(slide, spec):
    region = spec['content_region']
    if len(region) != 4 or min(region[2:]) <= 0:
        raise ValueError('content_region must be [x,y,w,h] in EMU')
    evidence = []
    for layer, owner in layers(slide).items():
        policy = spec['layers'][layer]
        if layer != 'slide' and owner._element.xpath('.//*[local-name()="timing"]'):
            raise ValueError(f'{layer}: inherited animation requires a separate audited cleanup plan')
        keep, remove = set(policy['keep']), set(policy['remove'])
        names = [s.name for s in owner.shapes]
        if len(names) != len(set(names)):
            raise ValueError(f'{layer}: duplicate shape names require an unambiguous source page')
        if layer != 'slide' and set(names) == keep and not keep & remove:
            continue  # shared inherited layer was already audited and cleaned in this output
        if keep & remove or keep | remove != set(names):
            raise ValueError(f'{layer}: every shape must have exactly one keep/remove decision')
        reasons = policy.get('keep_reasons', {})
        if any(not reasons.get(n) for n in keep):
            raise ValueError(f'{layer}: every kept object needs a decoration/header/footer reason')
        for shape in list(owner.shapes):
            text = getattr(shape, 'text', '')
            evidence.append({'layer': layer, 'name': shape.name, 'action': 'keep' if shape.name in keep else 'remove',
                             'text': text, 'box': [shape.left, shape.top, shape.width, shape.height],
                             'reason': reasons.get(shape.name, 'body content/layout')})
            if shape.name in remove:
                shape._element.getparent().remove(shape._element)
            else:
                x,y,w,h = region
                # Canvas intrusion must be explicitly reviewed, including inherited decorations.
                overlap = max(0,min(x+w,shape.left+shape.width)-max(x,shape.left))*max(0,min(y+h,shape.top+shape.height)-max(y,shape.top))
                if overlap and shape.name not in policy.get('canvas_overlap_reviewed', []):
                    raise ValueError(f'{layer}/{shape.name}: retained object intersects canvas; review explicitly')
    # Body pages have new objects and navigation; inherited layers are isolated.
    from template_animation import _strip
    for owner in layers(slide).values():
        _strip(owner._element,'timing')
        _strip(owner._element,'transition')
    return evidence


def style_evidence(slide):
    runs=[]
    def visit(shapes):
        for s in shapes:
            if s.shape_type == 6:
                visit(s.shapes)
            if getattr(s,'has_text_frame',False):
                for p in s.text_frame.paragraphs:
                    for r in p.runs:
                        runs.append({'font':r.font.name,'size':r.font.size.pt if r.font.size else None,
                                     'rPr': r._r.rPr.xml if r._r.rPr is not None else None})
    for owner in layers(slide).values(): visit(owner.shapes)
    theme=next((r.target_part for r in slide.slide_layout.slide_master.part.rels.values() if r.reltype.endswith('/theme')),None)
    return {'runs':runs, 'master_part':str(slide.slide_layout.slide_master.part.partname),
            'theme_part':str(theme.partname) if theme else None,
            'theme_xml':theme.blob.decode('utf-8') if theme else None}


def inject(slide, spec, page, *, notes_enabled=True, asset_base=None):
    if notes_enabled and not page.get('notes','').strip():
        raise ValueError('explicit page notes are required')
    if not page.get('source','').strip():
        raise ValueError('page source is required')
    import render_pptx as renderer
    from copy import deepcopy
    style=spec['style']; x,y,w,h=[value/914400 for value in spec['content_region']]
    blocks=page.get('blocks')
    if blocks is None:
        items=page.get('items',[])
        if not items or any(not it.get('title') or not it.get('body') for it in items):
            raise ValueError('body requires core blocks or nonempty semantic items')
        blocks=[{'kind':'rows','items':[{'tag':it['title'],'text':it['body']} for it in items]}]
    if not isinstance(blocks,list) or not blocks:
        raise ValueError('body requires nonempty core blocks')
    for layer, name in spec.get('header', []):
        if layer != 'slide':
            raise ValueError('variable headers must be slide-local, not shared layout/master objects')
        shape=next(s for s in slide.shapes if s.name==name)
        replace_text(shape,page.get('section','方案汇报'))
    def mix(color, other, fraction):
        return RGBColor(*(round(a*(1-fraction)+b*fraction) for a,b in zip(color,other)))
    accent=RGBColor.from_hex(style['accent']); ink=RGBColor.from_hex(style['text'])
    white=RGBColor(255,255,255)
    values={'FONT':style['font'],'TEXT':ink,'SUBTLE':mix(ink,white,.2),
            'FAINT':mix(ink,white,.35),'ACCENT':accent,'ACCENT_DK':accent,
            'ACCENT_LT':mix(accent,white,.9),'LINE':mix(accent,white,.8),
            'PANEL':mix(accent,white,.96),'WHITE':white}
    if asset_base is not None: values['BASE_DIR']=Path(asset_base).resolve()
    previous={key:getattr(renderer,key) for key in values}
    formulas=deepcopy(renderer.FORMULA_SRC)
    first=len(slide.shapes)
    try:
        for key,value in values.items(): setattr(renderer,key,value)
        # Optional opaque reading surface occupies only the audited body area.
        if spec.get('body_panel'):
            panel=renderer.add_rect(slide,x,y,w,h,white)
            panel.name='CanvasReadingSurface'
        padding=.18
        problems=check_box(page['title'],int((w-2*padding)*914400),600000,26,'CanvasTitle')
        if problems: raise ValueError('; '.join(problems))
        renderer.add_text(slide,x+padding,y+.12,w-2*padding,.58,page['title'],size=26,bold=True,color=accent).name='CanvasTitle'
        citation=page.get('external_citation','').strip()
        bottom=.4 if citation else .12
        renderer.render_body(slide,{'title':page['title'],'blocks':blocks},style,
                             (x+padding,y+.88,w-2*padding,h-.88-bottom),enlarge=False)
        if citation:
            renderer.add_text(slide,x+padding,y+h-.35,w-2*padding,.28,citation,size=10,color=ink).name='CanvasCitation'
        boxes=[]
        for shape in list(slide.shapes)[first:]:
            row={'name':shape.name,'text':getattr(shape,'text',''),
                 'box':[shape.left,shape.top,shape.width,shape.height]}
            if shape.name in renderer.FORMULA_SRC:
                latex,display=renderer.FORMULA_SRC[shape.name]
                row['formula']={'latex':latex,'display':display}
            boxes.append(row)
        return boxes
    except RuntimeError as exc:
        raise ValueError(str(exc)) from exc
    finally:
        for key,value in previous.items(): setattr(renderer,key,value)
        renderer.FORMULA_SRC.clear(); renderer.FORMULA_SRC.update(formulas)


def replace_text(shape, value):
    # Preserve paragraph/run properties and geometry of the source slot.
    tf=shape.text_frame
    p=tf.paragraphs[0]
    if p.runs:
        p.runs[0].text=value
        for r in list(p.runs)[1:]: p._p.remove(r._r)
    else: p.text=value
    for other in list(tf.paragraphs)[1:]: tf._txBody.remove(other._p)


def fill_structure(slide, page, inventory):
    """Fill audited original text slots; leave all geometry and groups intact."""
    from template_fill import resolve_key, walk_shapes
    from template_capacity import check_text
    from template_contract import fingerprint
    before=fingerprint(slide)
    fields=page.get('fields',{})
    if not fields: raise ValueError('structural page requires explicit fields')
    supplied={resolve_key(key,inventory,page['template_slide'])['key']:value for key,value in fields.items()}
    kept={resolve_key(key,inventory,page['template_slide'])['key']:reason for key,reason in page.get('keep_text',{}).items()}
    slots=[rec for rec in inventory['keys'].values() if rec['kind']=='text' and rec['shape'].text.strip()]
    if any(rec['key'] not in supplied and not kept.get(rec['key']) for rec in slots):
        raise ValueError('structural page has unclassified original text slots')
    targets={rec['steps']:rec['shape'] for rec in walk_shapes(list(slide.shapes))}
    evidence=[]
    for key,value in supplied.items():
        rec=inventory['keys'][key]
        if rec['kind']!='text' or not isinstance(value,str) or not value.strip():
            raise ValueError('structural fields must contain nonempty text for existing slots')
        problems=check_text(rec['shape'],value,key,width=rec.get('display_width'),height=rec.get('display_height'))
        if problems: raise ValueError('; '.join(problems))
        replace_text(targets[rec['steps']],value)
        evidence.append({'key':key,'text':value})
    if fingerprint(slide)!=before: raise ValueError('structural fill changed geometry')
    return evidence


def generate(template, manifest, output, deck=None, *, asset_base=None):
    specs=manifest['mothers']; pages=deck['slides'] if deck else [{'mother':s['name']} for s in specs]
    output=Path(output)
    if Path(template).resolve()==output.resolve(): raise ValueError('original template is immutable')
    if not isinstance(specs,list) or not specs or len({s['name'] for s in specs})!=len(specs):
        raise ValueError('mothers must be nonempty and have unique names')
    # Validate original inherited decisions before shared layers are edited.
    prs=Presentation(str(template)); source=prs; original_count=len(prs.slides); evidence=[]
    inherited_policies={}
    for spec in specs:
        if not isinstance(spec['template_slide'],int) or not 1 <= spec['template_slide'] <= original_count:
            raise ValueError('template_slide is out of range')
        x,y,w,h=spec['content_region']
        if min(x,y)<0 or min(w,h)<=0 or x+w>prs.slide_width or y+h>prs.slide_height:
            raise ValueError('content_region is outside the slide')
        origin=prs.slides[spec['template_slide']-1]
        for layer,owner in layers(origin).items():
            policy=spec['layers'][layer]; keep=set(policy['keep']);remove=set(policy['remove'])
            if keep & remove or keep | remove != {s.name for s in owner.shapes}:
                raise ValueError(f'{layer}: original inventory does not match explicit decisions')
            if layer=='slide': continue
            key=str(owner.part.partname); policy=spec['layers'][layer]
            if key in inherited_policies and inherited_policies[key] != policy:
                raise ValueError('conflicting decisions for shared inherited layer '+key)
            inherited_policies[key]=policy
    from slide_clone import clone_slide, isolate_layout
    from template_fill import scan_template
    from template_animation import build_plan, apply_plan, verify_plan, write_plan
    inventory=scan_template(str(template))
    specs_by_name={s['name']:s for s in specs}
    animation_items=[]
    for page in pages:
        if 'mother' in page:
            spec=specs_by_name[page['mother']]
            animation_items.append({'template_slide':spec['template_slide'],'extra_shapes':[{}],'animation_policy':'strip'})
        else:
            animation_items.append({'template_slide':page['template_slide'],'fields':page.get('fields',{})})
    plan=build_plan(template,animation_items,inventory)
    layout_cache={}
    for index,page in enumerate(pages,1):
        selected=animation_items[index-1]['template_slide']
        origin=source.slides[selected-1]
        slide=clone_slide(prs,origin)
        before=style_evidence(slide)
        if 'mother' in page:
            spec=specs_by_name[page['mother']]
            isolate_layout(prs,slide,layout_cache)
            plan['pages'][index-1]['layout_part']=str(slide.slide_layout.part.partname)
            plan['pages'][index-1]['master_part']=str(slide.slide_layout.slide_master.part.partname)
            audit=clean(slide,spec)
            boxes=inject(slide,spec,page,notes_enabled=deck.get('notes_enabled',True),asset_base=asset_base) if deck else []
            row={'mother':spec['name'],'decisions':audit,'injected':boxes}
        else:
            row={'structural':True,'fields':fill_structure(slide,page,inventory[selected-1])}
        row.update(template_slide=selected,output_slide=index,style_extraction=before,
                   remaining_text=[(s.name,getattr(s,'text','')) for s in slide.shapes])
        if deck:
            items=page.get('items',[])
            row['content']={'source':page.get('source'),'effective_items':len(items),
                'block_kinds':[b['kind'] for b in page.get('blocks',[])],
                'body_characters':sum(len(item['body']) for item in items),
                'content_coverage':[{'title':item['title'],'body_characters':len(item['body'])} for item in items]}
        evidence.append(row)
    for _ in range(original_count):
        entry=prs.slides._sldIdLst[0];prs.part.drop_rel(entry.rId);prs.slides._sldIdLst.remove(entry)
    if deck:
        from speaker_notes import apply_notes, verify_notes
        notes_report=apply_notes(prs,deck,preserve_template=True)
        for row,note in zip(evidence,notes_report): row['notes']=note
        problems=verify_notes(prs,deck)
        if problems: raise ValueError('\n'.join(problems))
    output.parent.mkdir(parents=True,exist_ok=True)
    fd,tmp=tempfile.mkstemp(suffix='.pptx',dir=output.parent);os.close(fd)
    try:
        prs.save(tmp)
        apply_plan(tmp,plan)
        repair_package(tmp,repair_animations=False)
        problems=package_problems(tmp)+verify_plan(tmp,plan)
        if problems: raise ValueError('\n'.join(problems))
        os.replace(tmp,output)
    finally:
        Path(tmp).unlink(missing_ok=True)
    Path(str(output)+'.canvas-evidence.json').write_text(json.dumps(evidence,ensure_ascii=False,indent=2))
    write_plan(plan,output)
    formulas=[dict(slide=row['output_slide'],shape=shape['name'],left=shape['box'][0],top=shape['box'][1],
                   width=shape['box'][2],height=shape['box'][3],**shape['formula'])
              for row in evidence for shape in row.get('injected',[]) if 'formula' in shape]
    sidecar=Path(str(output)+'.formulas.json')
    if formulas: sidecar.write_text(json.dumps(formulas,ensure_ascii=False,indent=2))
    else: sidecar.unlink(missing_ok=True)
    return evidence


def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--template',required=True);ap.add_argument('--manifest',required=True);ap.add_argument('--deck');ap.add_argument('--output',required=True);a=ap.parse_args()
    generate(a.template,json.loads(Path(a.manifest).read_text()),a.output,json.loads(Path(a.deck).read_text()) if a.deck else None,
             asset_base=Path(a.deck).resolve().parent if a.deck else None)

if __name__=='__main__': main()
