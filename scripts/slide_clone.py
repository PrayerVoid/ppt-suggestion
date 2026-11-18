"""Clone a slide within its package with complete, typed relationship remapping.

Layouts/themes/media are shared. Tags belong to a slide and are copied per clone.
Notes are created independently by the notes writer. The library's cached slide
root and shape tree remain intact; source XML child order is retained.
"""
from copy import deepcopy
import uuid
import re

from power_pptx.opc.package import Part

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def clone_slide(prs, origin):
    slide = prs.slides.add_slide(origin.slide_layout)
    root = slide._element
    cs = root.find('{'+P+'}cSld')
    tree = slide.shapes._spTree
    source_cs = origin._element.find('{'+P+'}cSld')
    source_tree = origin.shapes._spTree
    tree.clear()
    tree.attrib.update(source_tree.attrib)
    for child in source_tree:
        tree.append(deepcopy(child))
    cs.clear()
    cs.attrib.update(source_cs.attrib)
    for child in source_cs:
        cs.append(tree if child is source_tree else deepcopy(child))
    root.attrib.clear()
    root.attrib.update(origin._element.attrib)
    for child in list(root):
        root.remove(child)
    for child in origin._element:
        root.append(cs if child is source_cs else deepcopy(child))
    for node in root.xpath('.//*[local-name()="creationId"]'):
        if 'val' in node.attrib:
            node.set('val', str(uuid.uuid4().int % (2**32-1)))
        if 'id' in node.attrib:
            node.set('id', '{'+str(uuid.uuid4()).upper()+'}')
    mapping = {}
    for element in root.iter():
        for attr, rid in list(element.attrib.items()):
            if not attr.startswith('{'+R+'}'):
                continue
            if rid not in mapping:
                rel = origin.part.rels[rid]
                if rel.is_external:
                    mapped = slide.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
                else:
                    target = rel.target_part
                    if rel.reltype == R+'/tags':
                        if target.rels:
                            raise ValueError('tag part with child relationships is unsupported')
                        target = Part.load(prs.part.package.next_partname('/ppt/tags/tag%d.xml'),
                                           target.content_type, prs.part.package, target.blob)
                    mapped = slide.part.rels.get_or_add(rel.reltype, target)
                mapping[rid] = mapped
            element.set(attr, mapping[rid])
    if origin.has_notes_slide:
        slide.notes_slide.notes_text_frame.text = origin.notes_slide.notes_text_frame.text
    return slide


def replace_text(shape, value):
    """Replace slot content while retaining paragraph and first-run formatting."""
    frame=shape.text_frame
    prototype=deepcopy(frame.paragraphs[0]._p)
    for paragraph in list(frame.paragraphs):
        frame._txBody.remove(paragraph._p)
    from power_pptx.text.text import _Paragraph
    for line in str(value).split('\n'):
        element=deepcopy(prototype)
        frame._txBody.append(element)
        paragraph=_Paragraph(element,frame)
        run=paragraph.runs[0] if paragraph.runs else paragraph.add_run()
        for child in list(element):
            if child.tag.rsplit('}',1)[-1] in ('r','br','fld') and child is not run._r:
                element.remove(child)
        run.text=line


def remove_original_slides(prs, count):
    for _ in range(count):
        entry = prs.slides._sldIdLst[0]
        prs.part.drop_rel(entry.rId)
        prs.slides._sldIdLst.remove(entry)


def isolate_layout(prs, slide, cache):
    """Give cleaned body pages their own layout/master, leaving structural pages intact."""
    from lxml import etree
    layout = slide.slide_layout.part
    key = str(layout.partname)
    if key not in cache:
        master = slide.slide_layout.slide_master.part
        package = prs.part.package
        from power_pptx.opc.packuri import PackURI
        reserved = {str(part.partname) for part in package.iter_parts()}

        def next_name(pattern):
            index = 1
            while pattern % index in reserved:
                index += 1
            name = pattern % index
            reserved.add(name)
            return PackURI(name)

        def copy_part(source, pattern, excluded):
            result = type(source).load(next_name(pattern), source.content_type, package, source.blob)
            mapping = {}
            for rel in source.rels.values():
                if rel.reltype in excluded:
                    continue
                if rel.is_external:
                    rid = result.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
                else:
                    target = rel.target_part
                    if rel.reltype == R+'/tags':
                        if target.rels: raise ValueError('tag part with child relationships is unsupported')
                        target = Part.load(next_name('/ppt/tags/tag%d.xml'), target.content_type, package, target.blob)
                    rid = result.rels.get_or_add(rel.reltype, target)
                mapping[rel.rId] = rid
            for element in result._element.iter():
                for attr, rid in list(element.attrib.items()):
                    if attr.startswith('{'+R+'}') and rid in mapping:
                        element.set(attr, mapping[rid])
            return result

        new_master = copy_part(master, '/ppt/slideMasters/slideMaster%d.xml', {R+'/slideLayout'})
        new_layout = copy_part(layout, '/ppt/slideLayouts/slideLayout%d.xml', {R+'/slideMaster'})
        new_layout.rels.get_or_add(R+'/slideMaster', new_master)
        layout_rid = new_master.rels.get_or_add(R+'/slideLayout', new_layout)
        layout_list = new_master._element.find('{'+P+'}sldLayoutIdLst')
        if layout_list is None:
            raise ValueError('slide master lacks layout registration')
        layout_list.clear()
        etree.SubElement(layout_list, '{'+P+'}sldLayoutId', {'id':'2147483649', '{'+R+'}id':layout_rid})
        master_rid = prs.part.rels.get_or_add(R+'/slideMaster', new_master)
        master_list = prs.part._element.find('{'+P+'}sldMasterIdLst')
        next_id = max([2147483647]+[int(n.get('id')) for n in master_list])+1
        etree.SubElement(master_list, '{'+P+'}sldMasterId', {'id':str(next_id), '{'+R+'}id':master_rid})
        cache[key] = new_layout
    old = next(rel for rel in slide.part.rels.values() if rel.reltype == R+'/slideLayout')
    slide.part.drop_rel(old.rId)
    slide.part.rels.get_or_add(R+'/slideLayout', cache[key])


def prepared_edit_ops(ops, original_count):
    """Retarget edits from appended clones to an already prepared slide package."""
    edits = []
    for operation in ops:
        if operation.get('command') == 'add' and operation.get('parent') == '/' and operation.get('from'):
            continue
        if operation.get('command') == 'remove' and re.fullmatch(r'/slide\[\d+\]', operation.get('path', '')):
            continue
        operation = dict(operation)
        for field in ('path', 'parent', 'from'):
            if field in operation:
                def remap(match):
                    index = int(match[1]) - original_count
                    if index < 1:
                        raise ValueError('edit unexpectedly targets an original template slide')
                    return f'/slide[{index}]'
                operation[field] = re.sub(r'^/slide\[(\d+)\]', remap, operation[field])
        edits.append(operation)
    return edits


def clone_selected(template, selected, output):
    from power_pptx import Presentation
    prs = Presentation(str(template))
    originals = list(prs.slides)
    for index in selected:
        clone_slide(prs, originals[index-1])
    remove_original_slides(prs, len(originals))
    prs.save(str(output))
