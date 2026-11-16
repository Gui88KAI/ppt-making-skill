"""Exercise real PPTX bindings and failure modes, not documentation wording."""
import copy
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches

from deck_lint import lint, recommended_image_slots
from delivery_check import inspect
from entry_motion import click_fades


class DeliveryCheckTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        Image.new("RGB", (40, 30), "navy").save(self.root / "scene.png")
        Image.new("RGB", (40, 30), "gold").save(self.root / "wrong.png")
        self.prs = Presentation()
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.picture = self.slide.shapes.add_picture(str(self.root / "scene.png"), 0, 0)
        self.texts = [self.slide.shapes.add_textbox(Inches(i), Inches(2), Inches(1), Inches(1)) for i in range(2)]
        for i, shape in enumerate(self.texts):
            shape.text = f"提示{i + 1}"
        self.groups = [[shape.shape_id] for shape in self.texts]
        click_fades(self.slide, self.groups)
        self.path = self.root / "deck.pptx"
        self.plan = {"pages": [{"page": 1, "images_required": True,
                               "asset_bindings": [{"file": "scene.png", "shape_id": self.picture.shape_id}],
                               "motion": {"entry_groups": [{"purpose": "依次揭示"}]},
                               "written_motion": {"effect": "fade", "groups": self.groups}}]}
        self.brief = {"page_count": {"mode": "fixed", "value": 1}}

    def check(self):
        self.prs.save(self.path)
        return inspect(self.path, self.plan, self.brief, self.root)

    def test_real_picture_and_ordered_click_groups_match_without_claiming_playback(self):
        self.plan["native_playback"] = "passed"
        result = self.check()
        self.assertEqual(result["static_status"], "passed")
        self.assertEqual(result["pages"][0]["image_bindings_checked"], 1)
        self.assertEqual(result["pages"][0]["motion_comparison"], "matched_file_only")
        self.assertEqual(result["native_playback"], "not_performed_by_this_tool")
        self.assertEqual(result["visual_review"], "not_performed_by_this_tool")

    def test_same_filename_binding_must_match_actual_embedded_bytes(self):
        self.prs.save(self.path)
        (self.root / "scene.png").write_bytes((self.root / "wrong.png").read_bytes())
        result = inspect(self.path, self.plan, self.brief, self.root)
        self.assertEqual(result["static_status"], "failed")
        self.assertEqual(result["pages"][0]["image_bindings_checked"], 0)

    def test_required_image_cannot_be_satisfied_by_prompt_or_candidate_count(self):
        page = self.plan["pages"][0]
        page.update(asset_bindings=[], image_candidates=3, prompt="a fine illustration")
        self.assertTrue(self.check()["errors"])

    def test_textbox_id_is_not_accepted_as_a_picture(self):
        self.plan["pages"][0]["asset_bindings"][0]["shape_id"] = self.texts[0].shape_id
        self.assertTrue(self.check()["errors"])

    def test_file_and_page_constraints_are_checked(self):
        for constraint, fail in [(2, True), ({"mode": "maximum", "value": 2}, False),
                                 ({"mode": "range", "min": 2, "max": 3}, True),
                                 ({"mode": "delegate"}, False)]:
            with self.subTest(constraint=constraint):
                self.brief["page_count"] = constraint
                self.assertEqual(bool(self.check()["errors"]), fail)
        self.brief["page_count"] = 1
        self.plan["pages"][0]["asset_bindings"][0]["file"] = "missing.png"
        self.assertTrue(self.check()["errors"])

    def test_plans_or_receipts_cannot_substitute_for_written_animation(self):
        timing = self.slide._element.find(qn("p:timing"))
        self.slide._element.remove(timing)
        self.assertTrue(self.check()["errors"])
        self.plan["pages"][0]["written_motion"]["groups"] = []
        self.assertTrue(self.check()["errors"])
        self.plan["pages"][0].pop("written_motion")
        self.assertTrue(self.check()["errors"])

    def test_reversed_click_sequence_is_detected(self):
        self.plan["pages"][0]["written_motion"]["groups"] = list(reversed(self.groups))
        self.assertEqual(self.check()["pages"][0]["motion_comparison"], "failed")

    def test_unsupported_effect_is_explicitly_unverified(self):
        self.slide._element.find(".//" + qn("p:animEffect")).set("filter", "wipe")
        result = self.check()
        self.assertEqual(result["pages"][0]["motion_comparison"], "unsupported_needs_manual_check")
        self.assertEqual(result["static_status"], "passed_with_warnings")

    def test_broken_animation_target_fails_even_without_plan(self):
        self.slide._element.find(".//" + qn("p:spTgt")).set("spid", "999")
        self.prs.save(self.path)
        self.assertTrue(inspect(self.path)["errors"])

    def test_grouped_picture_binding_is_checked(self):
        group = self.slide.shapes.add_group_shape()
        picture = group.shapes.add_picture(str(self.root / "wrong.png"), 0, Inches(3))
        self.plan["pages"][0]["asset_bindings"].append({"file": "wrong.png", "shape_id": picture.shape_id})
        result = self.check()
        self.assertFalse(result["errors"])
        self.assertEqual(result["pages"][0]["image_bindings_checked"], 2)

    def test_partial_plan_is_not_a_full_deck_verification(self):
        self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.assertTrue(self.check()["errors"])

    def test_cli_report_cannot_overwrite_inputs(self):
        self.prs.save(self.path)
        before = self.path.read_bytes()
        run = subprocess.run([sys.executable, str(Path(__file__).with_name("delivery_check.py")),
                              str(self.path), "--report", str(self.path)], capture_output=True)
        self.assertNotEqual(run.returncode, 0)
        self.assertEqual(self.path.read_bytes(), before)

    def test_image_led_density_uses_meaningful_roles_and_recommended_coverage(self):
        self.plan["visual_mode"] = "image-led"
        self.plan["pages"][0]["asset_bindings"][0]["image_role"] = "scene"
        result = self.check()
        self.assertEqual(result["visual_density"]["status"], "insufficient")
        self.assertEqual(result["visual_density"]["meaningful_image_slots"], 1)
        self.assertEqual(result["visual_density"]["recommended_slots"], 2)
        self.assertEqual(result["static_status"], "passed_with_warnings")


class PlanningPolicyTests(unittest.TestCase):
    def test_visual_coverage_targets_scale_with_page_count(self):
        self.assertEqual(recommended_image_slots(1), 2)
        self.assertEqual(recommended_image_slots(8), 4)
        self.assertEqual(recommended_image_slots(12), 6)
        self.assertEqual(recommended_image_slots(15), 8)
        self.assertGreater(recommended_image_slots(18), 8)

    def test_image_led_plan_warns_when_theme_images_are_too_few(self):
        page = {
            "purpose": "具体情境进入讨论",
            "claim": "从一次解锁看见选择入口",
            "visual_anchor": "情境场景",
            "logic_role": "hook",
            "reading_order": ["问题", "场景", "提示"],
            "layout_family": "editorial",
            "asset_bindings": [{"image_role": "scene", "file": "scene.png", "shape_id": 1}],
        }
        pages = [copy.deepcopy(page) for _ in range(8)]
        for extra in pages[2:]:
            extra["asset_bindings"] = []
        errors, warnings, _ = lint({
            "visual_mode": "image-led",
            "pages": pages,
        })
        self.assertFalse(errors)
        self.assertTrue(any("默认建议至少 4 个" in item for item in warnings))

    def test_theme_role_is_counted_but_texture_is_not(self):
        page = {
            "purpose": "比较两种边界",
            "claim": "时间与空间各自解决一种负担",
            "visual_anchor": "左右对照",
            "logic_role": "comparison",
            "reading_order": ["左侧", "右侧", "结论"],
            "layout_family": "comparison",
            "asset_bindings": [
                {"image_role": "situation", "file": "scene.png", "shape_id": 1},
                {"image_role": "texture", "file": "paper.png", "shape_id": 2},
            ],
        }
        _, warnings, _ = lint({"visual_mode": "image-led", "pages": [page]})
        self.assertTrue(any("只有 1 个已记录主题图位" in item for item in warnings))

    def test_small_deck_may_use_consistent_layout_and_one_selected_image(self):
        page = {"purpose": "比较同口径数据", "visual_anchor": "chart", "layout_family": "grid",
                "images_required": True, "image_candidates": 1, "motion": {"transition": "none"}}
        errors, warnings, counts = lint({"pages": [copy.deepcopy(page), copy.deepcopy(page)]})
        self.assertEqual((errors, warnings), ([], []))
        self.assertEqual(counts["grid"], 2)

    def test_four_planned_steps_are_not_rejected_by_a_helper_limit(self):
        page = {"purpose": "解释四步过程", "visual_anchor": "path",
                "motion": {"entry_groups": [{"purpose": str(i)} for i in range(4)]}}
        errors, warnings, _ = lint({"pages": [page]})
        self.assertFalse(errors)
        self.assertTrue(warnings)  # A plan is still not a file-level implementation.

    def test_prompt_compiler_records_requested_variants_without_fabricating_images(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            src, out = root / "source.json", root / "prompts.json"
            src.write_text(json.dumps({"visual_bible": {"style": "watercolor"},
                                       "slides": [{"number": 1, "image": {"key": "scene", "prompt": "a letter"}}]}), encoding="utf-8")
            cmd = [sys.executable, str(Path(__file__).with_name("compile_prompts.py")), str(src), "-o", str(out)]
            result = subprocess.run(cmd, capture_output=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            data = json.loads(out.read_text(encoding="utf-8"))
            self.assertEqual(len(data["images"]), 1)
            self.assertEqual(data["selection_rule"]["actual_images_generated"], 0)
            self.assertEqual(data["images"][0]["style"], "watercolor")


if __name__ == "__main__":
    unittest.main()
