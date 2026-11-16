#!/usr/bin/env python3
"""Compare final PPTX pages, picture bindings and basic fade groups with a plan.

This read-only check never certifies visual quality or native playback. It does
not download assets, modify a deck or trust self-reported review status.
"""
import argparse
import hashlib
import json
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn

from deck_lint import is_meaningful_binding, recommended_image_slots
from intake import check_pages


def sha256(data):
    return hashlib.sha256(data).hexdigest()


def positive_int(value):
    return isinstance(value, int) and not isinstance(value, bool) and value > 0


def walk_shapes(shapes):
    for shape in shapes:
        yield shape
        if hasattr(shape, "shapes"):
            yield from walk_shapes(shape.shapes)


def fade_groups(timing):
    """Return ordered target groups only for the supported basic click-fade form.

    None means an animation form requiring manual inspection, not a pass.
    """
    if timing is None:
        return []
    mains = [n for n in timing.iter(qn("p:cTn")) if n.get("nodeType") == "mainSeq"]
    if len(mains) != 1:
        return None
    steps = mains[0].find(qn("p:childTnLst"))
    if steps is None or not len(steps):
        return None
    # General motion/rotation/scale/value animations are beyond this parser.
    unsupported = ("anim", "animMotion", "animRot", "animScale", "animClr", "cmd", "audio", "video")
    if any(timing.find(".//" + qn("p:" + tag)) is not None for tag in unsupported):
        return None
    result = []
    for step in steps:
        effects = [n for n in step.iter(qn("p:cTn"))
                   if n.get("nodeType") in ("clickEffect", "withEffect", "afterEffect")]
        if not effects or effects[0].get("nodeType") != "clickEffect":
            return None
        if any(n.get("nodeType") != "withEffect" for n in effects[1:]):
            return None
        targets = []
        for effect in effects:
            animations = list(effect.iter(qn("p:animEffect")))
            if effect.get("presetClass") != "entr" or len(animations) != 1:
                return None
            animation = animations[0]
            if animation.get("filter") != "fade" or animation.get("transition") != "in":
                return None
            target = animation.find(".//" + qn("p:spTgt"))
            if target is None or not (target.get("spid") or "").isdigit():
                return None
            targets.append(int(target.get("spid")))
        result.append(sorted(targets))
    parsed = [sid for group in result for sid in group]
    all_targets = {int(n.get("spid")) for n in timing.iter(qn("p:spTgt"))
                   if (n.get("spid") or "").isdigit()}
    if len(parsed) != len(set(parsed)) or set(parsed) != all_targets:
        return None
    if sum(1 for _ in timing.iter(qn("p:animEffect"))) != len(parsed):
        return None
    return result


def inspect(path, plan=None, brief=None, base_dir=None):
    path = Path(path)
    base = Path(base_dir) if base_dir is not None else path.parent
    report = {
        "file": str(path.resolve()), "sha256": sha256(path.read_bytes()),
        "errors": [], "warnings": [], "pages": [],
        "visual_review": "not_performed_by_this_tool",
        "native_playback": "not_performed_by_this_tool",
        "scope": "Page count, exact embedded image bindings and supported click-fade targets/order only",
    }
    errors, warnings = report["errors"], report["warnings"]
    prs = Presentation(path)
    count = report["slide_count"] = len(prs.slides)
    pages = None
    visual_mode = None
    meaningful_image_slots = 0
    missing_image_roles = []
    if plan is not None:
        if not isinstance(plan, dict) or not isinstance(plan.get("pages"), list) or not plan["pages"]:
            errors.append("plan 必须包含非空 pages 数组；生图 slides 不能代替实际页序")
        else:
            pages = plan["pages"]
            visual_mode = plan.get("visual_mode") or (plan.get("visual_bible") or {}).get("visual_mode")
            if len(pages) != count:
                errors.append(f"计划 {len(pages)} 页，实际 {count} 页")
    else:
        warnings.append("未提供逐页计划，仅生成文件清单，未核对图片/动画绑定")
    if visual_mode is None and isinstance(brief, dict):
        visual_mode = brief.get("visual_mode")
    if brief is not None:
        constraint, error = check_pages(brief.get("page_count") if isinstance(brief, dict) else None)
        if error:
            errors.append(error)
        elif ((constraint["mode"] == "fixed" and count != constraint["value"])
              or (constraint["mode"] == "maximum" and count > constraint["value"])
              or (constraint["mode"] == "range" and not constraint["min"] <= count <= constraint["max"])):
            errors.append(f"实际 {count} 页不符合用户页数约束 {constraint}")
    else:
        warnings.append("未提供brief，未核对用户页数约束")

    for number, slide in enumerate(prs.slides, 1):
        shapes = {s.shape_id: s for s in walk_shapes(slide.shapes)}
        pictures = {}
        for sid, shape in shapes.items():
            if shape._element.tag == qn("p:pic"):
                try:
                    pictures[sid] = sha256(shape.image.blob)
                except (KeyError, ValueError, AttributeError):
                    errors.append(f"P{number}: 图片对象 {sid} 无法读取嵌入图片")
        timing = slide._element.find(qn("p:timing"))
        targets = []
        if timing is not None:
            for node in timing.iter(qn("p:spTgt")):
                raw = node.get("spid") or ""
                if not raw.isdigit() or int(raw) not in shapes:
                    errors.append(f"P{number}: 动画目标 {raw!r} 不存在")
                elif int(raw) not in targets:
                    targets.append(int(raw))
        actual_groups = fade_groups(timing)
        row = {"page": number, "pictures": [{"shape_id": s, "sha256": h} for s, h in pictures.items()],
               "animation_targets": targets, "basic_fade_groups": actual_groups,
               "image_bindings_checked": 0, "motion_comparison": "not_requested"}
        report["pages"].append(row)
        if pages is None or number > len(pages):
            continue
        page = pages[number - 1]
        if not isinstance(page, dict):
            errors.append(f"P{number}: 页面记录必须是对象")
            continue
        if "page" in page and (not positive_int(page["page"]) or page["page"] != number):
            errors.append(f"P{number}: page 编号必须与实际页序一致")
        bindings = page.get("asset_bindings", [])
        if not isinstance(bindings, list):
            errors.append(f"P{number}: asset_bindings 必须是数组")
            bindings = []
        if page.get("images_required") and not bindings:
            errors.append(f"P{number}: 要求主题图却没有实际图片绑定")
        bound = set()
        for binding in bindings:
            if not isinstance(binding, dict):
                errors.append(f"P{number}: 图片绑定必须是对象")
                continue
            if not (binding.get("image_role") or binding.get("role")):
                missing_image_roles.append(number)
            sid, filename = binding.get("shape_id"), binding.get("file")
            if not positive_int(sid) or sid not in pictures or sid in bound:
                errors.append(f"P{number}: 图片目标 {sid!r} 无效、非图片或重复绑定")
                continue
            bound.add(sid)
            if not isinstance(filename, str) or not filename.strip() or "://" in filename:
                errors.append(f"P{number}: 图片 {sid} 需绑定实际本地文件，不读取远程URL")
                continue
            asset = Path(filename)
            if not asset.is_absolute():
                asset = base / asset
            if not asset.is_file():
                errors.append(f"P{number}: 图片 {sid} 的绑定文件不存在: {filename}")
            elif sha256(asset.read_bytes()) != pictures[sid]:
                errors.append(f"P{number}: 图片 {sid} 与绑定文件内容不一致: {filename}")
            else:
                row["image_bindings_checked"] += 1
                if is_meaningful_binding(binding):
                    meaningful_image_slots += 1
        if set(pictures) - bound:
            warnings.append(f"P{number}: 图片对象 {sorted(set(pictures) - bound)} 未绑定，未证明其符合图位计划")

        planned = page.get("motion") or {}
        planned = planned.get("entry_groups") if isinstance(planned, dict) else None
        written = page.get("written_motion")
        if planned and written is None:
            errors.append(f"P{number}: 只有动画计划，没有实际写入的对象绑定")
        if written is None:
            if timing is not None:
                warnings.append(f"P{number}: 存在动画，未提供实际组绑定，顺序待核对")
            continue
        if not isinstance(written, dict) or not isinstance(written.get("groups"), list):
            errors.append(f"P{number}: written_motion 需要 effect 和 groups")
            continue
        groups = written["groups"]
        if planned and not groups:
            errors.append(f"P{number}: 存在页内动画计划，实际写入组却为空")
            continue
        if any(not isinstance(g, list) or not g or any(not positive_int(s) or s not in shapes for s in g) for g in groups):
            errors.append(f"P{number}: 声称写入的动画组含无效对象")
            continue
        flat = [sid for group in groups for sid in group]
        if len(flat) != len(set(flat)):
            errors.append(f"P{number}: 基本动画绑定含重复对象")
            continue
        if groups and timing is None:
            errors.append(f"P{number}: 声称写入动画，但实际文件没有时间线")
            row["motion_comparison"] = "failed"
        elif written.get("effect") != "fade" or actual_groups is None:
            warnings.append(f"P{number}: 动画形式不在基本淡入解析范围，需作者工具及实机核对组与顺序")
            row["motion_comparison"] = "unsupported_needs_manual_check"
        elif [sorted(g) for g in groups] != actual_groups:
            errors.append(f"P{number}: 实际点击淡入组或顺序与绑定不一致")
            row["motion_comparison"] = "failed"
        else:
            row["motion_comparison"] = "matched_file_only"
    if visual_mode == "image-led" and missing_image_roles:
        warnings.append("图片绑定缺少 image_role: " + ", ".join(f"P{n}" for n in sorted(set(missing_image_roles))))
    recommended_slots = recommended_image_slots(count)
    report["visual_density"] = {
        "visual_mode": visual_mode or "unspecified",
        "meaningful_image_slots": meaningful_image_slots,
        "recommended_slots": recommended_slots if visual_mode == "image-led" else None,
        "status": "not_applicable"
        if visual_mode != "image-led"
        else "pass"
        if meaningful_image_slots >= recommended_slots
        else "insufficient",
    }
    if visual_mode == "image-led" and meaningful_image_slots < recommended_slots:
        warnings.append(
            f"image-led 实际只有 {meaningful_image_slots} 个已绑定主题图位，"
            f"默认建议至少 {recommended_slots} 个；限制生图次数时请用承担叙事的原生图示补位并记录原因"
        )
    report["static_status"] = "failed" if errors else "passed_with_warnings" if warnings else "passed"
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx")
    parser.add_argument("--plan", help="最终逐页记录JSON，必须包含pages")
    parser.add_argument("--brief", help="包含用户page_count约束的brief.json")
    parser.add_argument("--report")
    args = parser.parse_args()
    def read(name):
        return json.loads(Path(name).read_text(encoding="utf-8-sig")) if name else None
    result = inspect(args.pptx, read(args.plan), read(args.brief),
                     Path(args.plan).resolve().parent if args.plan else None)
    output = json.dumps(result, ensure_ascii=False, indent=2)
    if args.report:
        report_path = Path(args.report).resolve()
        inputs = [Path(n).resolve() for n in (args.pptx, args.plan, args.brief) if n]
        if report_path in inputs:
            parser.error("--report 不能覆盖PPTX、计划或brief")
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(output, encoding="utf-8")
    print(output)
    return 1 if result["errors"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
