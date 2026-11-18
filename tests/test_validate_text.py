"""Regression checks for template groups, text insets and baseline exemptions."""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from power_pptx import Presentation
from power_pptx.util import Inches, Pt
from power_pptx.oxml.xmlchemy import OxmlElement
from validate_pptx import (audited_template_page, bbox, check_slide, collect_baseline,
                           estimate_text_height, overlap_key, positioned_shapes)


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

    def audited_check(self, output):
        context, baseline = audited_template_page(self.slide, output, 13.3333, 7.5)
        hard, soft, info = [], [], []
        check_slide(output, 1, context, hard, soft, 13.3333, 7.5, baseline, info)
        return hard, info

    def cloned(self):
        from slide_clone import clone_slide
        return clone_slide(self.prs, self.slide)

    def test_only_unchanged_source_overflow_is_info(self):
        self.textbox(self.slide.shapes, '原样艺术字', w=.3, h=.2, size=44)
        output = self.cloned()
        hard, info = self.audited_check(output)
        self.assertFalse(hard)
        self.assertTrue(any('文本框溢出' in item for item in info))
        output.shapes[0].text = '新增正文必须失败'
        hard, _ = self.audited_check(output)
        self.assertTrue(any('文本框溢出' in item for item in hard))

    def test_one_changed_collision_member_stays_hard(self):
        self.textbox(self.slide.shapes, '原有标签', x=1, y=1, w=2, h=1, size=18)
        self.textbox(self.slide.shapes, '原有正文', x=1.3, y=1.1, w=2, h=1, size=18)
        output = self.cloned()
        self.assertFalse(self.audited_check(output)[0])
        output.shapes[0].text_frame.paragraphs[0].runs[0].text = '新的标签'
        self.assertTrue(any('形状重叠' in item for item in self.audited_check(output)[0]))

    def test_style_change_invalidates_existing_overflow(self):
        self.textbox(self.slide.shapes, '相同文字', w=.3, h=.2, size=44)
        output = self.cloned()
        output.shapes[0].text_frame.paragraphs[0].runs[0].font.size = Pt(50)
        self.assertTrue(any('文本框溢出' in item for item in self.audited_check(output)[0]))

    def test_added_oob_and_zero_size_stay_hard(self):
        self.textbox(self.slide.shapes, '原有越界', x=14, h=1)
        output = self.cloned()
        self.assertFalse(self.audited_check(output)[0])
        self.textbox(output.shapes, '新增越界', x=14, y=4, h=1)
        self.textbox(output.shapes, '新增零尺寸', x=3, y=4, w=0, h=0)
        hard, _ = self.audited_check(output)
        self.assertTrue(any('元素越界' in item for item in hard))
        self.assertTrue(any('零尺寸' in item for item in hard))

    def test_source_page_does_not_exempt_another_page(self):
        self.textbox(self.slide.shapes, '原样艺术字', w=.3, h=.2, size=44)
        output = self.cloned()
        other = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        context, baseline = audited_template_page(other, output, 13.3333, 7.5)
        hard = []
        check_slide(output, 1, context, hard, [], 13.3333, 7.5, baseline, [])
        self.assertTrue(any('文本框溢出' in item for item in hard))

    def test_cli_requires_source_map_and_rejects_changed_text(self):
        import contextlib
        import io
        import json
        import tempfile
        from unittest.mock import patch
        import validate_pptx
        self.textbox(self.slide.shapes, '原样艺术字', w=.3, h=.2, size=44)
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            template, output, mapping = root/'template.pptx', root/'output.pptx', root/'map.json'
            self.prs.save(template)
            self.prs.save(output)
            mapping.write_text(json.dumps({'slides': [{'template_slide': 1}]}))
            argv = ['validate_pptx.py', str(output), '--baseline', str(template)]
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                with patch.object(sys, 'argv', argv):
                    self.assertEqual(validate_pptx.main(), 2)
                with patch.object(sys, 'argv', argv + ['--baseline-map', str(mapping)]):
                    self.assertEqual(validate_pptx.main(), 0)
                changed = Presentation(output)
                changed.slides[0].shapes[0].text = '新增文字'
                changed.save(output)
                with patch.object(sys, 'argv', argv + ['--baseline-map', str(mapping)]):
                    self.assertEqual(validate_pptx.main(), 1)

    def test_mixed_run_width_and_single_glyph_height(self):
        from validate_pptx import paragraph_width
        shape = self.textbox(self.slide.shapes, 'A', w=3, h=1, size=40)
        run = shape.text_frame.paragraphs[0].add_run()
        run.text = 'small'; run.font.size = Pt(10)
        mixed = paragraph_width(shape.text_frame.paragraphs[0], 'Asmall')
        run.font.size = Pt(40)
        self.assertLess(mixed, paragraph_width(shape.text_frame.paragraphs[0], 'Asmall'))
        shape.text = '可'
        shape.text_frame.paragraphs[0].runs[0].font.size = Pt(100)
        self.assertAlmostEqual(estimate_text_height(shape.text_frame, .1),
                               estimate_text_height(shape.text_frame, 3))


if __name__ == "__main__":
    unittest.main()
