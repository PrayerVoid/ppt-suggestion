import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import deck_check


class DeckContentChecks(unittest.TestCase):
    def test_single_point_chart_is_rejected(self):
        deck = {"slides": [{"type": "content", "title": "单点数据应使用核心数字",
                            "blocks": [{"kind": "chart", "chart_type": "column_clustered",
                                        "categories": ["时延"], "series": [{"name": "P99", "values": [0.5786]}],
                                        "source": "资料中的单点测量值"}]}]}
        errors = deck_check.validate_deck(deck, template_mode=True)
        self.assertTrue(any("只有一个数据点" in error for error in errors))

    def test_humanize_phrasing_is_not_a_structural_error(self):
        deck = {"slides": [{"type": "content", "title": "能力边界先决定接口范围",
                            "blocks": [{"kind": "lead", "text": "先把资料整理成表格，把结果放进页面。"}]}]}
        errors = deck_check.validate_deck(deck, template_mode=True)
        self.assertEqual(errors, [])


if __name__ == "__main__":
    unittest.main()
