"""Regression checks for template groups, text insets and baseline exemptions."""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from power_pptx import Presentation
from power_pptx.util import Inches, Pt
from power_pptx.oxml.xmlchemy import OxmlElement
from validate_pptx import (bbox, check_slide, collect_baseline, estimate_text_height,
                           overlap_key, positioned_shapes)


class TemplateTextTests(unittest.TestCase):
    def setUp(self):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(13.3333), Inches(7.5)
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

    def textbox(self, shapes, text, x=1, y=1, w=3, h=0.5, size=22):
        shape = shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        shape.text = text
        shape.text_frame.word_wrap = True
        shape.text_frame.paragraphs[0].runs[0].font.size = Pt(size)
        return shape

    def issues(self, baseline=frozenset()):
        hard, soft = [], []
        check_slide(self.slide, 1, None, hard, soft, 13.3333, 7.5, baseline, [])
        return hard

    def test_group_title_overflow_is_not_baseline_exempt(self):
        group = self.slide.shapes.add_group_shape()
        title = self.textbox(group.shapes, "短标题")
        baseline = collect_baseline(self.prs)
        title.text_frame.paragraphs[0].runs[0].text = "四项能力：模板组合扩展的长标题"
        self.assertTrue(any("文本框溢出" in error for error in self.issues(baseline)))

    def test_nested_group_coordinates(self):
        outer = self.slide.shapes.add_group_shape()
        inner = outer.shapes.add_group_shape()
        self.textbox(inner.shapes, "文字", x=1, y=2, w=2, h=1)
        # Set explicit local transforms after group auto-bounds have settled.
        for group, off, ext, child_off, child_ext in (
            (inner, (3, 4), (4, 2), (1, 2), (2, 1)),
            (outer, (1, 1), (6, 4), (3, 4), (4, 2)),
        ):
            xf = group._element.grpSpPr.xfrm
            for name, attrs, values in (("off", ("x", "y"), off),
                                        ("ext", ("cx", "cy"), ext),
                                        ("chOff", ("x", "y"), child_off),
                                        ("chExt", ("cx", "cy"), child_ext)):
                node = getattr(xf, name)
                for key, value in zip(attrs, values):
                    node.set(key, str(int(Inches(value))))
        self.assertEqual(len(list(positioned_shapes(self.slide.shapes))), 1)
        actual = bbox(next(positioned_shapes(self.slide.shapes)))
        for value, expected in zip(actual, (1, 1, 6, 4)):
            self.assertAlmostEqual(value, expected)

    def test_nowrap_width_includes_tracking(self):
        shape = self.textbox(self.slide.shapes, "测试", w=0.85, h=1)
        shape.text_frame.word_wrap = False
        run = shape.text_frame.paragraphs[0].runs[0]
        run._r.get_or_add_rPr().set("spc", "1200")
        self.assertTrue(any("不换行文字横向溢出" in error for error in self.issues()))

    def test_margins_and_paragraph_spacing_count(self):
        shape = self.textbox(self.slide.shapes, "测试", w=3, h=0.5)
        tf = shape.text_frame
        before = estimate_text_height(tf, 3)
        tf.margin_top = Inches(0.2)
        tf.paragraphs[0].space_after = Pt(12)
        self.assertGreater(estimate_text_height(tf, 3), before + 0.3)
        self.assertTrue(any("文本框溢出" in error for error in self.issues()))

    def test_paragraph_default_font_size_is_used(self):
        shape = self.textbox(self.slide.shapes, "测试", w=3, h=0.4)
        paragraph = shape.text_frame.paragraphs[0]
        paragraph.runs[0].font.size = None
        ppr = paragraph._p.get_or_add_pPr()
        default = OxmlElement("a:defRPr")
        default.set("sz", "3600")
        ppr.append(default)
        self.assertTrue(any("文本框溢出" in error for error in self.issues()))

    def test_decoration_pattern_survives_translation(self):
        a = self.textbox(self.slide.shapes, "甲", x=1, y=1, w=2, h=1)
        b = self.textbox(self.slide.shapes, "乙", x=1.15, y=1.03, w=2, h=1)
        key = overlap_key(a, b)
        for shape in (a, b):
            shape.left += Inches(6.02)
            shape.top += Inches(0.001)
        self.assertEqual(overlap_key(a, b), key)
        b.text = "新的文字"
        self.assertNotEqual(overlap_key(a, b), key)

    def test_text_containment_is_not_a_panel_exemption(self):
        self.textbox(self.slide.shapes, "完全重合的文字", x=1, y=1, w=3, h=1)
        self.textbox(self.slide.shapes, "覆盖", x=1, y=1, w=3, h=1)
        self.assertTrue(any("形状重叠" in error for error in self.issues()))

    def test_changed_text_collision_stays_hard(self):
        a = self.textbox(self.slide.shapes, "原有标签", x=1, y=1, w=2, h=1, size=18)
        b = self.textbox(self.slide.shapes, "原有正文", x=1.3, y=1.1, w=2, h=1, size=18)
        baseline = collect_baseline(self.prs)
        a.text_frame.paragraphs[0].runs[0].text = "新的标签"
        self.assertTrue(any("形状重叠" in error for error in self.issues(baseline)))


if __name__ == "__main__":
    unittest.main()
