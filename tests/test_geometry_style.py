"""The decoration pass must use the actual renderer palette and own its shapes."""
import contextlib
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "extensions" / "geometry-deco"))
from power_pptx import Presentation
import add_deco


class GeometryStyleTests(unittest.TestCase):
    def test_palette_conflict_rejects_before_writing(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            prs = Presentation(); prs.slides.add_slide(prs.slide_layouts[6])
            src, out, deck = root / "src.pptx", root / "out.pptx", root / "deck.json"
            prs.save(src)
            deck.write_text(json.dumps({"slides": [{"type": "cover"}]}))
            (root / "style.json").write_text(json.dumps({"palette": {"ACCENT": "8B4265"},
                "enabled_extensions": ["geometry-deco"], "geometry_motif": "circles"}))
            argv = ["add_deco.py", str(src), "--deck", str(deck), "--out", str(out), "--accent", "4F6082"]
            with patch.object(sys, "argv", argv), contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(add_deco.main(), 1)
            self.assertFalse(out.exists())

    def test_style_is_applied_and_owned_decorations_are_idempotent(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            prs = Presentation(); prs.slides.add_slide(prs.slide_layouts[6])
            src, out, deck = root / "src.pptx", root / "out.pptx", root / "deck.json"
            prs.save(src)
            deck.write_text(json.dumps({"slides": [{"type": "cover"}]}))
            (root / "style.json").write_text(json.dumps({"palette": {"ACCENT": "8B4265"},
                "enabled_extensions": ["geometry-deco"], "geometry_motif": "circles"}))
            argv = ["add_deco.py", str(src), "--deck", str(deck), "--out", str(out)]
            with patch.object(sys, "argv", argv), contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(add_deco.main(), 0)
            first = len(Presentation(out).slides[0].shapes)
            self.assertGreater(first, 0)
            evidence = json.loads(Path(str(out) + ".render.json").read_text())
            self.assertEqual(evidence["decoration"]["accent"], "8B4265")
            self.assertEqual(evidence["decoration"]["owner"], "ppt-suggestion.geometry-deco")
            argv[1] = str(out)
            with patch.object(sys, "argv", argv), contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(add_deco.main(), 0)
            self.assertEqual(len(Presentation(out).slides[0].shapes), first)


if __name__ == "__main__":
    unittest.main()
