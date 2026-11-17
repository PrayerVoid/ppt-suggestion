import ast
import random
import sys
import tempfile
import unittest
from pathlib import Path
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import render_pptx as renderer
import deck_check
import sample_style

# Match the documented copy-in workflow, including the renderer's dependencies.
LAYOUT = dict(vars(renderer))
exec(compile((ROOT / "extensions/special-layout/blocks.py").read_text(),
             "special-layout/blocks.py", "exec"), LAYOUT)


class SpecialLayoutChecks(unittest.TestCase):
    def setUp(self):
        self.prs, blank = renderer.new_presentation()
        self.slide = self.prs.slides.add_slide(blank)

    def test_core_and_extension_registration_are_disjoint(self):
        core = {"bignum", "asym", "quote", "marginalia", "quadrant"}
        special = {"diagonal", "radial", "ring", "concentric"}
        self.assertTrue(core <= renderer.BLOCK_RENDERERS.keys())
        self.assertTrue(core <= deck_check.BUILTIN_KINDS)
        self.assertTrue(special.isdisjoint(renderer.BLOCK_RENDERERS))
        # agenda is rendered by render_catalog rather than render_block.
        self.assertEqual(set(renderer.BLOCK_RENDERERS) | {"agenda"}, deck_check.BUILTIN_KINDS)
        tree = ast.parse((ROOT / "extensions/special-layout/blocks.py").read_text())
        functions = {node.name for node in tree.body if isinstance(node, ast.FunctionDef)}
        for kind in special:
            self.assertIn("_" + kind, functions)
        for kind in core:
            self.assertNotIn("_" + kind, functions)
        self.assertNotIn("_layout_items", functions)

    def test_extra_kinds_accepts_core_and_extension_names(self):
        kinds = {"bignum", "asym", "quote", "diagonal", "radial", "ring",
                 "marginalia", "concentric", "quadrant"}
        deck = {"slides": [{"type": "content", "title": "核心与扩展名称可以共同校验", "blocks": [
            {"kind": kind} for kind in sorted(kinds)]}]}
        self.assertEqual(deck_check.validate_deck(deck, extra_kinds=kinds, template_mode=True), [])
        self.assertEqual(deck_check.validate_deck(deck, extra_kinds={"diagonal", "radial", "ring", "concentric"},
                                                  template_mode=True), [])

    def test_core_kinds_require_density_and_diversity(self):
        deck = {"slides": [{"type": "content", "title": "短引文不能绕过正文密度", "blocks": [
            {"kind": "quote", "text": "原文", "source": "可追溯来源"}]}]}
        errors = deck_check.validate_deck(deck)
        self.assertTrue(any("110" in error for error in errors))
        self.assertTrue(any("版式种类过少" in error for error in errors))
        self.assertFalse(any("未知" in error for error in errors))

    def test_sample_candidates_use_core_only_and_do_not_request_quotes_or_values(self):
        for recipe in sample_style.RECIPES:
            for kind in recipe.split("+"):
                self.assertIn(kind, renderer.BLOCK_RENDERERS)
                self.assertNotIn(kind, {"bignum", "quote", "asym"})
        self.assertEqual(sample_style.block_spec(random.Random(0), "quadrant")["items"], 4)
        self.assertIn("marginalia", sample_style.ASYM)

    def test_excess_items_are_not_silently_discarded(self):
        for kind, field, limit in [("bignum", "items", 3), ("radial", "items", 6),
                                   ("ring", "items", 3), ("marginalia", "notes", 3),
                                   ("concentric", "rings", 3), ("quadrant", "items", 4)]:
            with self.subTest(kind=kind), self.assertRaises(ValueError):
                LAYOUT["_" + kind](self.slide, 0.7, 1.7, 12, {
                    field: [{}] * (limit + 1), "value": "placeholder"}, {})
        with self.assertRaises(ValueError):
            LAYOUT["_asym"](self.slide, 0.7, 1.7, 12,
                            {"right": {"items": ["text"] * 5}}, {})

    def test_radial_nodes_stay_inside_reported_height(self):
        for count in range(3, 7):
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
            y = 1.7
            height = LAYOUT["_radial"](slide, 0.7, y, 12,
                {"center": {"title": "中心"}, "items": [{"title": "节点"}] * count}, {})
            for shape in slide.shapes:
                self.assertGreaterEqual(shape.top / 914400, y - 0.001)
                self.assertLessEqual((shape.top + shape.height) / 914400, y + height + 0.001)

    def test_concentric_labels_map_outer_to_inner_and_height_is_not_capped(self):
        height = LAYOUT["_concentric"](self.slide, 0.7, 1.7, 12,
            {"rings": [{"label": label, "text": "解释内容。" * 70}
                       for label in ("外层", "中层", "内层")]}, {})
        self.assertGreater(height, 4.1)
        texts = [shape.text for shape in self.slide.shapes if shape.has_text_frame]
        for index, label in enumerate(("外层", "中层", "内层"), 1):
            self.assertIn(str(index), texts)
            self.assertIn(f"{index}  {label}", texts)

    def test_figure_reserves_long_side_text_and_caption(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "image.png"
            Image.new("RGB", (1000, 300), "white").save(src)
            height = renderer._figure(self.slide, 0.7, 1.7, 12,
                {"src": str(src), "height": 0.6, "caption": "图注",
                 "side": {"items": ["侧栏说明文字需要完整保留。" * 12]}}, {})
            self.assertGreater(height, 0.92)
            for shape in self.slide.shapes:
                self.assertLessEqual((shape.top + shape.height) / 914400,
                                     1.7 + height + 0.001)


if __name__ == "__main__":
    unittest.main()
