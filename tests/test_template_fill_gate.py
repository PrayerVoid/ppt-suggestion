"""Generation must reject overflow without replacing a previous good output."""
import contextlib
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from power_pptx import Presentation
from power_pptx.util import Inches, Pt
import template_fill


class PublicationGateTests(unittest.TestCase):
    def test_layout_failure_preserves_existing_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            template, out = root / "template.pptx", root / "out.pptx"
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            group = slide.shapes.add_group_shape()
            shape = group.shapes.add_textbox(Inches(1), Inches(1), Inches(1), Inches(0.3))
            shape.text = "这个组合里的长标题必须被拦住"
            shape.text_frame.word_wrap = True
            shape.text_frame.paragraphs[0].runs[0].font.size = Pt(24)
            prs.save(template)
            out.write_bytes(b"previous successful output")
            deck, mapping = root / "deck.json", root / "map.json"
            deck.write_text(json.dumps({"slides": [{"title": "测试"}]}))
            mapping.write_text(json.dumps({"slides": [{"template_slide": 1}]}))
            argv = ["template_fill.py", "--template", str(template), "--deck", str(deck),
                    "--map", str(mapping), "--output", str(out)]
            errors = io.StringIO()
            with patch.object(sys, "argv", argv), patch.object(template_fill.ocli, "batch"), \
                    patch.object(template_fill.ocli, "close"), contextlib.redirect_stderr(errors):
                self.assertEqual(template_fill.main(), 1)
            self.assertIn("文本框溢出", errors.getvalue())
            self.assertEqual(out.read_bytes(), b"previous successful output")
            self.assertFalse(list(root.glob(".template-fill-*")))

    def test_clone_uses_group_dimension_properties_and_readback(self):
        with tempfile.TemporaryDirectory() as tmp:
            template = Path(tmp) / "template.pptx"
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            group = slide.shapes.add_group_shape()
            group.name = "卡片"
            shape = group.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
            shape.text = "文字"
            prs.save(template)
            ops, checks, _ = template_fill.build_ops(
                [{"template_slide": 1, "remove": ["卡片"], "clone_groups": [
                    {"source": "卡片", "width": "3in", "height": "2in"}]}],
                template_fill.scan_template(str(template)), Path(tmp))
            resize = next(op for op in ops if op["command"] == "set")
            self.assertEqual(resize["props"], {"width": "2743200emu", "height": "1828800emu"})
            self.assertTrue(any(check[0] == "geometry" for check in checks))

    def test_chart_categories_cannot_change_delimited_data_shape(self):
        with self.assertRaises(template_fill.MapError):
            template_fill.chart_props({"categories": ["一,二"], "series": [
                {"name": "数量", "values": [1]}], "source": "测试"}, (0, 0, 1, 1), 1)


if __name__ == "__main__":
    unittest.main()
