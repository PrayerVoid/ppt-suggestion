"""Native notesSlide speaker scripts derived only from supplied slide content.

notes_enabled defaults to true. Explicit slide.notes replaces the generated script.
Template notes are retained after a labeled boundary; opt-out never edits them.
"""
from __future__ import annotations

GENERATED_MARKER = "【本次讲稿】"
TEMPLATE_MARKER = "【模板原备注（供参考）】"


def generated_script(data):
    from deck_check import collect_text
    title = str(data.get("title", "")).strip()
    parts = [title] if title else []
    for field in ("subtitle", "kicker"):
        value = str(data.get(field, "")).strip()
        if value and value not in parts:
            parts.append(value)
    # Production provenance is internal evidence, never an audience-facing script.
    def audience_content(value):
        internal = {'source', 'sources', 'source_note', 'source_mapping', 'provenance',
                    'evidence', 'material_mapping', 'notes', 'notes_enabled'}
        if isinstance(value, dict):
            return {k: audience_content(v) for k,v in value.items() if k not in internal}
        if isinstance(value, list):
            return [audience_content(v) for v in value]
        return value
    for value in collect_text(audience_content(data)):
        text = str(value).strip()
        if text and text not in parts:
            parts.append(text)
    # This is deliberately extractive: no ungrounded claims, estimates or transitions.
    return "\n\n".join(parts)


def apply_notes(prs, deck, *, preserve_template=False):
    if deck.get("notes_enabled", True) is False:
        return []
    report = []
    for number, (slide, data) in enumerate(zip(prs.slides, deck.get("slides", [])), 1):
        explicit = data.get("notes")
        script = str(explicit).strip() if explicit is not None else generated_script(data)
        if not script:
            raise ValueError(f"第 {number} 页缺少可生成讲稿的内容")
        tf = slide.notes_slide.notes_text_frame
        existing = tf.text.strip()
        if GENERATED_MARKER in existing:
            existing = existing.split(TEMPLATE_MARKER, 1)[1].strip() if TEMPLATE_MARKER in existing else ""
        text = GENERATED_MARKER + "\n" + script
        if preserve_template and existing:
            text += "\n\n" + TEMPLATE_MARKER + "\n" + existing
        tf.text = text
        report.append({"slide": number, "mode": "explicit" if explicit is not None else "extractive",
                       "characters": len(script), "template_preserved": bool(preserve_template and existing)})
    return report


def verify_notes(prs, deck):
    if deck.get("notes_enabled", True) is False:
        return []
    errors = []
    for number, (slide, data) in enumerate(zip(prs.slides, deck.get("slides", [])), 1):
        text = slide.notes_slide.notes_text_frame.text
        expected = str(data["notes"]).strip() if data.get("notes") is not None else generated_script(data)
        if GENERATED_MARKER + "\n" + expected not in text:
            errors.append(f"第 {number} 页讲稿未写入原生备注")
    return errors
