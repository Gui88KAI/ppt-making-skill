"""Observable PPTX regression checks; run directly with the skill dependencies."""
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches

import deck_builder
from deck_lint import lint
from motion import MC, P159, apply_slide_motion, transition
from wps_compat import NS, explicit_fonts, inspect, prepare


class WpsCompatibilityTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name)
        self.prs = Presentation()
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.shape = self.slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
        self.shape.text = "知恩于心，感恩于行"

    def save(self, name="sample.pptx"):
        path = self.path / name
        self.prs.save(path)
        return path

    def test_wps_default_and_explicit_morph_use_safe_fallback(self):
        for spec, expected in [({}, "fade"), ({"motion": {"transition": "morph"}}, "fade"),
                               ({"motion": {"transition": "morph", "fallback": "none"}}, "none")]:
            self.assertEqual(apply_slide_motion(self.slide, spec, "wps"), expected)
            explicit_fonts(self.prs)
            report = inspect(self.save())
            self.assertFalse(report["errors"])
            self.assertEqual(report["transitions"]["P1"], [expected])
            self.assertEqual(report["native_wps_validation"]["status"], "not_performed")

    def test_morph_uses_extension_and_fallback_before_timing(self):
        self.slide._element.append(etree.Element(qn("p:timing")))
        transition(self.slide, "morph")
        wrapper = self.slide._element.find(f"{{{MC}}}AlternateContent")
        self.assertIsNotNone(wrapper.find(f".//{{{P159}}}morph"))
        self.assertIsNone(wrapper.find(".//" + qn("p:morph")))
        self.assertIsNotNone(wrapper.find(f"{{{MC}}}Fallback/" + qn("p:transition") + "/" + qn("p:fade")))
        self.assertLess(list(self.slide._element).index(wrapper), list(self.slide._element).index(self.slide._element.find(qn("p:timing"))))
        transition(self.slide, "fade")
        self.assertIsNone(self.slide._element.find(f"{{{MC}}}AlternateContent"))
        self.assertEqual(len(self.slide._element.findall(qn("p:transition"))), 1)

    def test_font_declarations_cover_all_scripts_and_keep_editable_text(self):
        self.shape.text_frame.paragraphs[0].runs[0].font.name = "微软雅黑"
        explicit_fonts(self.prs)
        props = self.shape.text_frame.paragraphs[0].runs[0]._r.rPr
        for tag in ("latin", "ea", "cs"):
            self.assertEqual(props.find(qn("a:" + tag)).get("typeface"), "微软雅黑")
        loaded = Presentation(self.save())
        self.assertEqual(loaded.slides[0].shapes[0].text, self.shape.text)
        self.assertEqual(inspect(self.save())["embedded_font_parts"], [])

    def test_prepare_preserves_source_and_geometry_and_requires_new_path(self):
        transition(self.slide, "morph")
        source = self.save("original.pptx")
        before = source.read_bytes()
        output = self.path / "safe.pptx"
        prepare(source, output)
        self.assertEqual(before, source.read_bytes())
        loaded = Presentation(output)
        new = loaded.slides[0].shapes[0]
        self.assertEqual((new.left, new.top, new.width, new.height, new.text),
                         (self.shape.left, self.shape.top, self.shape.width, self.shape.height, self.shape.text))
        self.assertFalse(inspect(output)["errors"])
        with self.assertRaises(ValueError):
            prepare(source, source)

    def test_normalizing_fonts_keeps_distinct_latin_and_east_asian_choices(self):
        run = self.shape.text_frame.paragraphs[0].runs[0]
        run.font.name = "Arial"
        props = run._r.rPr
        east = etree.SubElement(props, qn("a:ea"), {"typeface": "微软雅黑"})
        explicit_fonts(self.prs)
        self.assertEqual(props.find(qn("a:latin")).get("typeface"), "Arial")
        self.assertEqual(props.find(qn("a:ea")).get("typeface"), "微软雅黑")

    def test_template_route_keeps_content_and_writes_safe_transitions(self):
        source = self.save("template.pptx")
        mapping = self.path / "mapping.yaml"
        mapping.write_text("slides: []\n", encoding="utf-8")
        out = self.path / "filled.pptx"
        proc = subprocess.run([sys.executable, str(Path(__file__).with_name("fill_template.py")),
                               str(source), str(mapping), "-o", str(out), "--target", "wps", "--check"], capture_output=True)
        self.assertEqual(proc.returncode, 0, proc.stderr.decode(errors="replace"))
        self.assertEqual(Presentation(out).slides[0].shapes[0].text, self.shape.text)
        self.assertFalse(inspect(out)["errors"])

    def test_inspector_and_cli_reject_external_image(self):
        source = self.save()
        unsafe = self.path / "unsafe.pptx"
        with zipfile.ZipFile(source) as src, zipfile.ZipFile(unsafe, "w") as dst:
            for name in src.namelist():
                content = src.read(name)
                if name == "ppt/slides/_rels/slide1.xml.rels":
                    root = etree.fromstring(content)
                    etree.SubElement(root, f"{{{NS['r']}}}Relationship", {
                        "Id": "rId99", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        "Target": "https://example.invalid/external.png", "TargetMode": "External"})
                    content = etree.tostring(root)
                dst.writestr(name, content)
        self.assertTrue(any("外链" in e for e in inspect(unsafe)["errors"]))
        proc = subprocess.run([sys.executable, str(Path(__file__).with_name("wps_compat.py")), str(unsafe), "--strict"], capture_output=True)
        self.assertEqual(proc.returncode, 1)

    def test_static_checker_detects_original_wrong_morph(self):
        node = etree.SubElement(self.slide._element, qn("p:transition"))
        etree.SubElement(node, qn("p:morph"))
        self.assertTrue(inspect(self.save())["errors"])

    def test_lint_rejects_empty_spec_and_does_not_require_morph_ids_for_wps(self):
        errors, _, _ = lint({"slides": []})
        self.assertTrue(errors)
        _, warnings, _ = lint({"compatibility": {"target": "wps"}, "pages": [{"type": "quote", "title": "感恩"}]})
        self.assertFalse(any("共享运动" in w or "shared_ids" in w for w in warnings))

    def test_builder_generates_safe_file_and_notes(self):
        spec = {"compatibility": {"target": "wps"}, "pages": [{"type": "quote", "title": "感恩", "speaker_notes": "讨论一分钟"}]}
        out = self.path / "built.pptx"
        deck_builder.build(spec, str(out))
        self.assertFalse(inspect(out)["errors"])
        self.assertIn("讨论一分钟", Presentation(out).slides[0].notes_slide.notes_text_frame.text)

    def test_check_command_propagates_layout_failure(self):
        spec = self.path / "overflow.json"
        spec.write_text(json.dumps({"compatibility": {"target": "wps"}, "pages": [{"type": "cards", "title": "正文溢出", "items": [{"title": "项目", "note": "这是故意构造的超长正文。" * 160}]}]}, ensure_ascii=False), encoding="utf-8")
        proc = subprocess.run([sys.executable, str(Path(__file__).with_name("deck_builder.py")), str(spec), "-o", str(self.path / "overflow.pptx"), "--check"], capture_output=True)
        self.assertNotEqual(proc.returncode, 0)
        self.assertTrue((self.path / "overflow.wps-report.json").exists())


if __name__ == "__main__":
    unittest.main(verbosity=2)
