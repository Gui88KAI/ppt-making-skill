"""Behavioral tests of structured intake and routing; these do not simulate an AI."""
import copy
import math
import unittest
from pathlib import Path

from intake import TYPES, assess


class IntakeTests(unittest.TestCase):
    def complete(self):
        return dict(type="05", topic="感恩教育", audience="大学生", occasion="主题班会",
                    duration_minutes=15, page_count=12)

    def challenge(self):
        return dict(type="01", topic="智能巡检", audience="评委", occasion="挑战杯答辩",
                    duration_minutes=8, page_count=12, competition="挑战杯")

    def test_topic_alone_does_not_select_a_type_or_pages(self):
        result = assess({"topic": "感恩教育"})
        self.assertFalse(result["route_ready"])
        self.assertFalse(result["can_author"])
        self.assertIn("type", result["missing"])
        self.assertIn("page_count", result["missing"])
        self.assertEqual(len(next(q for q in result["questions"] if q["field"] == "type")["options"]), 5)

    def test_known_details_are_not_asked_again(self):
        brief = self.complete()
        del brief["type"]
        del brief["page_count"]
        result = assess(brief)
        self.assertEqual(result["missing"], ["type", "page_count"])
        self.assertEqual([q["field"] for q in result["questions"]], ["type", "page_count"])

    def test_complete_class_meeting_goes_directly_to_guide(self):
        result = assess(self.complete())
        self.assertTrue(result["route_ready"])
        self.assertTrue(result["can_author"])
        self.assertEqual(result["questions"], [])
        self.assertEqual(result["read_next"], "directions/05-课程与主题班会/GUIDE.md")

    def test_all_five_routes_are_separate_and_exist(self):
        root = Path(__file__).resolve().parents[1]
        paths = set()
        for code, (_, route, _) in TYPES.items():
            brief = dict(self.complete(), type=code)
            if code == "01":
                brief["competition"] = "校园创新赛"
            result = assess(brief)
            self.assertTrue(result["route_ready"])
            self.assertTrue((root / route).is_file())
            paths.add(result["read_next"])
        self.assertEqual(len(paths), 5)

    def test_duration_does_not_delegate_page_count(self):
        brief = self.complete()
        del brief["page_count"]
        self.assertIn("page_count", assess(brief)["missing"])

    def test_explicit_page_count_is_respected_even_if_small(self):
        result = assess(dict(self.complete(), type="03", page_count=8))
        self.assertTrue(result["can_author"])
        self.assertEqual(result["page_count"], {"mode": "fixed", "value": 8})

    def test_maximum_range_and_delegation_are_accepted(self):
        for pages in ({"mode": "maximum", "value": 9}, {"mode": "range", "min": 8, "max": 12}, {"mode": "delegate"}):
            with self.subTest(pages=pages):
                result = assess(dict(self.complete(), page_count=pages))
                self.assertTrue(result["can_author"])
                self.assertEqual(result["page_count"], pages)

    def test_invalid_page_limits_do_not_author(self):
        for pages in (0, -1, True, "auto", {"mode": "range", "min": 15, "max": 10}):
            with self.subTest(pages=pages):
                self.assertFalse(assess(dict(self.complete(), page_count=pages))["can_author"])

    def test_invalid_duration_is_not_ready(self):
        for duration in (0, -1, True, math.inf, math.nan, "15"):
            with self.subTest(duration=duration):
                self.assertFalse(assess(dict(self.complete(), duration_minutes=duration))["route_ready"])

    def test_invalid_type_never_falls_back(self):
        result = assess(dict(self.complete(), type="unknown"))
        self.assertFalse(result["can_author"])
        self.assertNotIn("read_next", result)

    def test_competition_without_name_requests_it(self):
        result = assess(dict(self.complete(), type="01"))
        self.assertTrue(result["route_ready"])
        self.assertFalse(result["can_author"])
        self.assertIn("competition", result["specialized_missing"])

    def test_challenge_missing_structure_does_not_claim_ready(self):
        result = assess(self.challenge())
        self.assertTrue(result["route_ready"])
        self.assertFalse(result["can_author"])
        self.assertEqual(result["status"], "needs_specialized_input")
        self.assertEqual(set(result["specialized_missing"]), {"competition_year", "competition_stage", "challenge_kind"})
        self.assertIsNone(result["content_ready"])
        self.assertEqual(result["delivery_status"], "not_checked")

    def test_academic_challenge_requires_category_not_funding(self):
        brief = dict(self.challenge(), competition_year=2025, competition_stage="国赛",
                     challenge_kind="academic")
        result = assess(brief)
        self.assertEqual(result["specialized_missing"], ["project_category"])
        brief["project_category"] = "自然科学论文"
        result = assess(brief)
        self.assertTrue(result["can_author"])
        self.assertNotIn("track", result["specialized_missing"])

    def test_entrepreneurship_challenge_requires_track(self):
        brief = dict(self.challenge(), competition_year=2026, competition_stage="省赛",
                     challenge_kind="创业计划竞赛")
        self.assertEqual(assess(brief)["specialized_missing"], ["track"])
        brief["track"] = "乡村振兴"
        self.assertTrue(assess(brief)["can_author"])

    def test_explicit_unknown_is_recorded_without_reasking(self):
        brief = dict(self.challenge(), competition_year="不知道", competition_stage="省赛",
                     challenge_kind="不知道")
        result = assess(brief)
        self.assertFalse(result["can_author"])
        self.assertEqual(set(result["specialized_unresolved"]), {"competition_year", "challenge_kind"})
        self.assertEqual(result["questions"], [])
        self.assertEqual(result["next_action"], "read_selected_guide_and_continue_draft")

    def test_unrecognized_challenge_kind_cannot_silently_pass(self):
        brief = dict(self.challenge(), competition_year=2025, competition_stage="国赛",
                     challenge_kind="某比赛")
        result = assess(brief)
        self.assertFalse(result["can_author"])
        self.assertEqual(result["specialized_missing"], ["challenge_kind"])

    def test_explicit_dachallenge_name_does_not_reask_kind(self):
        brief = dict(self.challenge(), competition="挑战杯大挑", competition_year=2025,
                     competition_stage="国赛", project_category="科技发明")
        result = assess(brief)
        self.assertTrue(result["can_author"])
        self.assertEqual(result["specialized_missing"], [])

    def test_separate_projects_keep_requested_directions(self):
        reading = assess(dict(self.complete(), type="02", occasion="独立阅读"))
        pitch = assess(dict(self.complete(), type="03", occasion="融资路演"))
        self.assertNotEqual(reading["read_next"], pitch["read_next"])
        self.assertTrue(reading["can_author"] and pitch["can_author"])

    def test_script_never_certifies_content_or_delivery(self):
        for code in ("02", "03", "04", "05"):
            result = assess(dict(self.complete(), type=code))
            self.assertIsNone(result["content_ready"])
            self.assertEqual(result["delivery_status"], "not_checked")

    def test_invalid_root_is_reported(self):
        result = assess(["not", "an", "object"])
        self.assertEqual(result["status"], "invalid")
        self.assertFalse(result["route_ready"])

    def test_input_is_not_modified_by_assessment(self):
        brief = self.challenge()
        original = copy.deepcopy(brief)
        assess(brief)
        self.assertEqual(brief, original)


if __name__ == "__main__":
    unittest.main(verbosity=2)
