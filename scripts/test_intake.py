"""Behavioral checks for fresh-agent intake routing; no presentation is authored."""
import copy
import unittest
from pathlib import Path

from intake import TYPES, assess


class IntakeTests(unittest.TestCase):
    def complete(self):
        return dict(type="05", topic="感恩教育", audience="大学生", occasion="主题班会", duration_minutes=15, page_count=12)

    def test_topic_alone_does_not_author_or_default_to_competition(self):
        result = assess({"topic": "感恩教育"})
        self.assertFalse(result["can_author"])
        self.assertEqual(result["status"], "needs_input")
        self.assertIsNone(result["type"])
        self.assertIn("type", result["missing"])
        self.assertIn("page_count", result["missing"])
        self.assertEqual(len(next(q for q in result["questions"] if q["field"] == "type")["options"]), 5)

    def test_known_details_are_not_asked_again(self):
        brief = self.complete()
        del brief["type"]
        del brief["page_count"]
        self.assertEqual(assess(brief)["missing"], ["type", "page_count"])

    def test_complete_brief_routes_without_extra_approval(self):
        result = assess(self.complete())
        self.assertTrue(result["can_author"])
        self.assertEqual(result["questions"], [])
        self.assertEqual(result["read_next"], "types/05-课程汇报.md")

    def test_all_five_routes_point_to_existing_files(self):
        root = Path(__file__).resolve().parents[1]
        for code, (_, route, _) in TYPES.items():
            brief = dict(self.complete(), type=code)
            result = assess(brief)
            self.assertTrue(result["can_author"])
            self.assertTrue((root / route).is_file())

    def test_duration_is_not_page_delegation(self):
        brief = self.complete()
        del brief["page_count"]
        self.assertIn("page_count", assess(brief)["missing"])

    def test_explicit_page_count_outside_suggested_range_is_respected(self):
        result = assess(dict(self.complete(), type="03", page_count=8))
        self.assertTrue(result["can_author"])
        self.assertEqual(result["page_count"], {"mode": "fixed", "value": 8})

    def test_maximum_range_and_delegation_are_accepted(self):
        for pages in ({"mode": "maximum", "value": 9}, {"mode": "range", "min": 8, "max": 12}, {"mode": "delegate"}):
            result = assess(dict(self.complete(), page_count=pages))
            self.assertTrue(result["can_author"])
            self.assertEqual(result["page_count"], pages)

    def test_invalid_page_limits_cannot_author(self):
        for pages in (0, -1, True, "auto", {"mode": "range", "min": 15, "max": 10}):
            self.assertFalse(assess(dict(self.complete(), page_count=pages))["can_author"])

    def test_invalid_type_does_not_fall_back(self):
        result = assess(dict(self.complete(), type="unknown"))
        self.assertFalse(result["can_author"])
        self.assertNotIn("read_next", result)

    def test_input_is_not_modified_by_assessment(self):
        brief = self.complete()
        original = copy.deepcopy(brief)
        assess(brief)
        self.assertEqual(brief, original)


if __name__ == "__main__":
    unittest.main(verbosity=2)
