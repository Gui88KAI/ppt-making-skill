#!/usr/bin/env python3
"""WPS preparation and file-based static inspection; never claims native validation."""
import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from pptx import Presentation
from pptx.oxml.xmlchemy import OxmlElement

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/package/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}


def target_of(spec):
    value = str((spec.get("compatibility") or {}).get("target", "powerpoint")).lower()
    if value not in ("wps", "powerpoint"):
        raise ValueError(f"未知播放目标 {value!r}；可用 wps / powerpoint")
    return value


def motion_kind(page, target="powerpoint", default=None):
    motion = page.get("motion") or {}
    kind = motion.get("transition", default or ("fade" if target == "wps" else "morph"))
    if kind not in ("morph", "fade", "none"):
        raise ValueError(f"未知转场 {kind!r}；可用 morph / fade / none")
    if target == "wps" and kind == "morph":
        fallback = motion.get("fallback", "fade")
        return fallback if fallback in ("fade", "none") else "fade"
    return kind


def explicit_fonts(prs, font=None, default="微软雅黑"):
    """Keep chosen run fonts unless an intentional replacement is requested."""
    changed = 0
    for slide in prs.slides:
        for run in slide._element.iter(f"{{{NS['a']}}}r", f"{{{NS['a']}}}fld"):
            props = run.find(f"{{{NS['a']}}}rPr")
            if props is None:
                props = OxmlElement("a:rPr")
                run.insert(0, props)
            chosen = {}
            for tag in ("latin", "ea", "cs"):
                current = props.find(f"{{{NS['a']}}}{tag}")
                name = current.get("typeface") if current is not None else None
                chosen[tag] = name if name and not name.startswith("+") else None
            fallback = chosen["latin"] or chosen["ea"] or default
            for tag in ("latin", "ea", "cs"):
                old = props.find(f"{{{NS['a']}}}{tag}")
                if old is not None:
                    props.remove(old)
            for tag in ("latin", "ea", "cs"):
                node = OxmlElement(f"a:{tag}")
                node.set("typeface", font or chosen[tag] or fallback)
                props.insert_element_before(node, "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst")
            changed += 1
    return changed


def prepare(source, output, font=None, strip_animations=False):
    if Path(source).resolve() == Path(output).resolve():
        raise ValueError("兼容副本不能覆盖原文件，请提供不同输出路径")
    from motion import transition
    prs = Presentation(source)
    runs = explicit_fonts(prs, font=font)
    removed = 0
    for slide in prs.slides:
        transition(slide, "fade")
        if strip_animations:
            for child in list(slide._element):
                if child.tag == f"{{{NS['p']}}}timing" or (
                    child.tag == f"{{{NS['mc']}}}AlternateContent"
                    and child.find(f".//{{{NS['p']}}}timing") is not None
                ):
                    slide._element.remove(child)
                    removed += 1
    Path(output).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output)
    return {"source": str(source), "output": str(output), "font_runs_normalized": runs,
            "font_replacement": font, "animation_blocks_removed": removed,
            "transitions": "fade", "layout_review_required": True}


def inspect(path):
    path = Path(path)
    errors, warnings = [], []
    fonts, transitions, missing_ea = set(), {}, 0
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        bad_crc = z.testzip()
        if bad_crc:
            errors.append(f"损坏的 ZIP 成员: {bad_crc}")
        slides = sorted((n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                        key=lambda n: int(re.search(r"slide(\d+)", n).group(1)))
        if not slides:
            errors.append("文件没有幻灯片")
        for name in names:
            low = name.lower()
            if "vbaproject" in low or low.startswith("ppt/activex/") or low.startswith("ppt/embeddings/"):
                errors.append(f"宏、ActiveX 或嵌入对象需移除或另行验证: {name}")
            if low.startswith("ppt/media/") and Path(low).suffix not in (".png", ".jpg", ".jpeg"):
                warnings.append(f"非 PNG/JPEG 媒体需在 WPS 验证或转换: {name}")
            if not name.endswith(".rels"):
                continue
            for rel in ET.fromstring(z.read(name)):
                if rel.get("TargetMode") == "External" and rel.get("Type", "").rsplit("/", 1)[-1] in ("image", "audio", "video", "media", "oleObject"):
                    errors.append(f"存在外链图片、媒体或对象: {name} -> {rel.get('Target')}")
        for index, name in enumerate(slides, 1):
            root = ET.fromstring(z.read(name))
            effects = [child.tag.rsplit("}", 1)[-1] for node in root.findall(".//p:transition", NS) for child in node]
            transitions[f"P{index}"] = effects or ["none"]
            if any(effect != "fade" for effect in effects):
                errors.append(f"P{index}: 转场 {effects} 超出 WPS 保守配置 fade/none")
            if root.find(".//p:timing", NS) is not None:
                warnings.append(f"P{index}: 包含元素级动画，需要 WPS 放映验证")
            for run in root.findall(".//a:r", NS) + root.findall(".//a:fld", NS):
                if not run.findtext("a:t", default="", namespaces=NS).strip():
                    continue
                props = run.find("a:rPr", NS)
                ea = None if props is None else props.find("a:ea", NS)
                if ea is None or not ea.get("typeface") or ea.get("typeface", "").startswith("+"):
                    missing_ea += 1
                if props is not None:
                    for tag in ("latin", "ea", "cs"):
                        f = props.find(f"a:{tag}", NS)
                        if f is not None and f.get("typeface"):
                            fonts.add(f.get("typeface"))
        if missing_ea:
            warnings.append(f"{missing_ea} 个文字 run 未显式声明具体 East Asian 字体")
        embedded = [n for n in names if n.startswith("ppt/fonts/") and not n.endswith("/")]
        font_list = ET.fromstring(z.read("ppt/presentation.xml")).find("p:embeddedFontLst", NS)
        if bool(embedded) != (font_list is not None and len(font_list) > 0):
            warnings.append("字体嵌入声明与字体资源不完整，需核对关系及目标设备结果")
    return {
        "file": str(path.resolve()), "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
        "target": "wps", "slide_count": len(slides), "fonts": sorted(fonts),
        "embedded_font_parts": embedded, "font_installation_checked": False,
        "transitions": transitions, "errors": sorted(set(errors)), "warnings": sorted(set(warnings)),
        "static_status": "failed" if errors else "passed_with_warnings" if warnings else "passed",
        "native_wps_validation": {"status": "not_performed", "version": None, "platform": None},
        "limitations": ["静态检查不验证 WPS 实际渲染、字体替换或放映", "已声明字体不等于已嵌入或在目标电脑安装"],
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx")
    parser.add_argument("--prepare", metavar="OUTPUT")
    parser.add_argument("--font", help="有意替换全部原生文字字体，仅与 --prepare 同用")
    parser.add_argument("--strip-animations", action="store_true")
    parser.add_argument("--report")
    parser.add_argument("--strict", action="store_true", help="静态 error 返回非零；不把静态通过当作实机通过")
    args = parser.parse_args()
    if not args.prepare and (args.font or args.strip_animations):
        parser.error("--font / --strip-animations 需要 --prepare 输出副本")
    changes = prepare(args.pptx, args.prepare, args.font, args.strip_animations) if args.prepare else None
    result = inspect(args.prepare or args.pptx)
    if changes:
        result["preparation"] = changes
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.report:
        Path(args.report).parent.mkdir(parents=True, exist_ok=True)
        Path(args.report).write_text(payload, encoding="utf-8")
    print(payload)
    return 1 if args.strict and result["errors"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
