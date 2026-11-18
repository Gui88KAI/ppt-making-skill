"""Exercise title verification against actual temporary PPTX objects."""
import tempfile
import unittest
from pathlib import Path

from pptx import Presentation
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches

from title_editability_check import inspect_title


class TitleEditabilityTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "case.pptx"
        self.deck = Presentation()
        self.slide = self.deck.slides.add_slide(self.deck.slide_layouts[6])

    def add_character(self, index, value, top=1):
        shape = self.slide.shapes.add_textbox(Inches(index), Inches(top), Inches(1), Inches(1))
        shape.name = f"标题-{index}-{value}"
        shape.text = value
        return shape

    def save_and_check(self, title, **kwargs):
        self.deck.save(self.path)
        return inspect_title(self.path, title, **kwargs)

    def test_separate_native_textboxes_are_accepted(self):
        self.add_character(1, "感", top=1)
        self.add_character(2, "恩", top=1.2)
        result = self.save_and_check("感恩", only_title_text=True)
        self.assertEqual(result["status"], "passed")
        self.assertEqual([row["text"] for row in result["characters"]], ["感", "恩"])
        self.assertNotEqual(result["characters"][0]["y_emu"], result["characters"][1]["y_emu"])

    def test_one_box_with_multiple_characters_is_rejected(self):
        self.add_character(1, "感恩")
        self.assertEqual(self.save_and_check("感恩")["status"], "failed")

    def test_title_character_order_is_checked_by_index(self):
        self.add_character(2, "恩")
        self.add_character(1, "感")
        self.assertEqual(self.save_and_check("感恩")["status"], "passed")
        self.assertEqual(self.save_and_check("恩感")["status"], "failed")

    def test_unrelated_cover_text_can_be_rejected_when_requested(self):
        self.add_character(1, "感")
        self.add_character(2, "恩")
        extra = self.slide.shapes.add_textbox(Inches(1), Inches(3), Inches(3), Inches(1))
        extra.text = "不需要的说明"
        self.assertEqual(self.save_and_check("感恩", only_title_text=True)["status"], "failed")
        self.assertEqual(self.save_and_check("感恩", only_title_text=False)["status"], "passed")

    def test_locked_textbox_is_rejected(self):
        shape = self.add_character(1, "感")
        locks = OxmlElement("a:spLocks")
        locks.set("noMove", "1")
        shape._element.nvSpPr.cNvSpPr.append(locks)
        self.assertEqual(self.save_and_check("感")["status"], "failed")

    def test_missing_file_and_out_of_range_slide_fail(self):
        self.assertEqual(inspect_title(self.path, "感")["status"], "failed")
        self.add_character(1, "感")
        self.assertEqual(self.save_and_check("感", slide_number=2)["status"], "failed")


if __name__ == "__main__":
    unittest.main(verbosity=2)
