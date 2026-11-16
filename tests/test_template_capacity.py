import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from power_pptx import Presentation
from power_pptx.util import Inches, Pt
from template_capacity import check_image, check_text


class TemplateCapacityTests(unittest.TestCase):
    def test_small_text_box_is_rejected_before_edit(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(1), Inches(0.25))
        shape.text = "模板标题"
        shape.text_frame.paragraphs[0].runs[0].font.size = Pt(24)
        self.assertTrue(check_text(shape, "这是一段明显装不下的长标题", "标题"))

    def test_replacement_image_checks_native_pixels(self):
        from PIL import Image
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "tiny.png"
            Image.new("RGB", (20, 20), "white").save(source)
            placeholder = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(3))
            self.assertTrue(check_image(source, placeholder, "图片"))


if __name__ == "__main__":
    unittest.main()
