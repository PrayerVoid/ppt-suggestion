"""Regression coverage for notes, injected palette and nonmutating capacity."""
import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from power_pptx import Presentation
from power_pptx.util import Inches, Pt
from speaker_notes import apply_notes, verify_notes
from template_capacity import check_text
from render_pptx import apply_style, add_text, add_hline


class NotesAndStyleTests(unittest.TestCase):
    def setUp(self):
        self.prs = Presentation()
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.deck = {"slides": [{"type": "cover", "title": "原始资料标题", "subtitle": "来源中的副标题"}]}

    def test_default_extracts_script_and_native_roundtrip(self):
        report = apply_notes(self.prs, self.deck)
        self.assertEqual(report[0]["mode"], "extractive")
        self.assertIn("原始资料标题", self.slide.notes_slide.notes_text_frame.text)
        with tempfile.TemporaryDirectory() as directory:
            p = Path(directory) / "notes.pptx"
            self.prs.save(p)
            self.assertEqual(verify_notes(Presentation(p), self.deck), [])

    def test_explicit_script_preserves_old_notes_without_repeating(self):
        self.slide.notes_slide.notes_text_frame.text = "模板旧备注"
        self.deck["slides"][0]["notes"] = "本页介绍原始资料中的范围。"
        apply_notes(self.prs, self.deck, preserve_template=True)
        first = self.slide.notes_slide.notes_text_frame.text
        self.assertIn("模板旧备注", first)
        self.assertIn("本页介绍", first)
        apply_notes(self.prs, self.deck, preserve_template=True)
        self.assertEqual(first, self.slide.notes_slide.notes_text_frame.text)

    def test_opt_out_preserves_existing(self):
        self.slide.notes_slide.notes_text_frame.text = "既有备注"
        self.deck["notes_enabled"] = False
        self.assertEqual(apply_notes(self.prs, self.deck, preserve_template=True), [])
        self.assertEqual(self.slide.notes_slide.notes_text_frame.text, "既有备注")

    def test_prepared_runtime_runs_without_source_directory(self):
        import subprocess
        from prepare_renderer import prepare
        from power_pptx import Presentation
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            runtime = prepare(root / "renderer")
            deck = root / "deck.json"
            deck.write_text(json.dumps(self.deck))
            out = root / "out.pptx"
            result = subprocess.run([sys.executable, str(runtime / "render_pptx.py"),
                "--input", str(deck), "--output", str(out)], cwd=root, capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(verify_notes(Presentation(out), self.deck), [])
            self.assertTrue(Path(str(out) + ".render.json").exists())

    def test_capacity_preserves_xml_and_uses_real_size(self):
        shape = self.slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(.5))
        run = shape.text_frame.paragraphs[0].add_run()
        run.text = "原文"
        run.font.size = Pt(48)
        run.font.bold = True
        xml = shape._element.xml
        self.assertTrue(check_text(shape, "容量检查不能丢失字号", "test"))
        self.assertEqual(xml, shape._element.xml)

    def test_template_clone_color_repair_uses_theme_roles(self):
        from template_fill import MapError, build_ops, scan_template
        group = self.slide.shapes.add_group_shape()
        group.name = "原子"
        shape = group.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
        shape.name = "正文"
        shape.text = "原文"
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            src = root / "template.pptx"
            self.prs.save(src)
            item = {"template_slide": 1, "clone_groups": [{"source": "原子", "fields": {
                "原子/正文": {"text": "替换正文", "color": "tx1"}}}]}
            ops, _, _ = build_ops([item], scan_template(src), root)
            self.assertTrue(any(op.get("props", {}).get("color") == "tx1" for op in ops))
            item["clone_groups"][0]["fields"]["原子/正文"]["color"] = "8B4265"
            with self.assertRaises(MapError):
                build_ops([item], scan_template(src), root)

    def test_style_injects_default_font_and_line_colors(self):
        import render_pptx as renderer
        keys = ("WHITE", "TEXT", "SUBTLE", "FAINT", "ACCENT", "ACCENT_DK", "ACCENT_LT", "LINE", "PANEL")
        old = {key: getattr(renderer, key) for key in keys}
        try:
            with tempfile.TemporaryDirectory() as directory:
                p = Path(directory) / "style.json"
                palette = {key: str(old[key]) for key in keys}
                palette.update(TEXT="242424", LINE="DFD4D9", ACCENT="8B4265")
                p.write_text(json.dumps({"palette": palette}))
                apply_style(p, self.deck)
                t = add_text(self.slide, 0, 0, 2, 1, "测试")
                line = add_hline(self.slide, 0, 1, 1)
                self.assertEqual(str(t.text_frame.paragraphs[0].runs[0].font.color.rgb), "242424")
                self.assertEqual(str(line.line.color.rgb), "DFD4D9")
                p.write_text(json.dumps({"palette": palette, "page_types": ["content"]}))
                with self.assertRaises(ValueError):
                    apply_style(p, self.deck)
        finally:
            for key, value in old.items():
                setattr(renderer, key, value)


if __name__ == "__main__":
    unittest.main()
