#!/usr/bin/env python3
"""检查 PPT 规格中的叙事、视觉、版式和动效连续性。"""
import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path
from wps_compat import motion_kind, target_of

def load(path):
    raw = Path(path).read_text(encoding="utf-8")
    return json.loads(raw) if path.endswith(".json") else __import__("yaml").safe_load(raw)

def lint(spec):
    if not isinstance(spec, dict):
        return ["规格必须是包含 pages 的对象"], [], Counter()
    pages = spec.get("pages") or []
    errors, warnings = [], []
    if not isinstance(spec.get("pages"), list) or not pages:
        errors.append("规格必须包含非空 pages 列表；visual_manifest 的 slides 不能直接作为页面规格")
        return errors, warnings, Counter()
    target = target_of(spec)
    default_transition = (spec.get("compatibility") or {}).get("preferred_transition")
    families = []
    motion_ids = []
    for i, page in enumerate(pages, 1):
        if not isinstance(page, dict):
            errors.append(f"P{i}: 页面规格必须是对象")
            continue
        family = page.get("layout_family") or page.get("layout") or page.get("type", "cards")
        families.append(family)
        required = ["visual_anchor"]
        missing = [k for k in required if not page.get(k)]
        if not (page.get("claim") or page.get("purpose")):
            missing.append("purpose/claim")
        if missing:
            warnings.append(f"P{i}: 缺少叙事字段 {', '.join(missing)}")
        motion = page.get("motion") or {}
        ids = set(motion.get("shared_ids") or [])
        motion_ids.append(ids)
        try:
            effective = motion_kind(page, target, default_transition)
        except ValueError as exc:
            errors.append(f"P{i}: {exc}")
            effective = "none"
        if effective == "morph" and not motion.get("shared_ids"):
            warnings.append(f"P{i}: 没有 motion.shared_ids，Morph 无法形成连续对象")
        if effective == "morph" and i > 1 and motion_ids[-2] and ids and not (motion_ids[-2] & ids):
            warnings.append(f"P{i}: 与上一页没有共享运动对象，页面间动画可能断开")
        if motion.get("entry_groups"):
            warnings.append(f"P{i}: entry_groups 仅是计划；实际文件用 delivery_check.py 或作者工具核对，再在目标软件放映")
        if page.get("images_required", False):
            n = page.get("image_candidates", page.get("candidates", 0))
            if not page.get("asset_bindings") and not page.get("image") and not n:
                warnings.append(f"P{i}: 需要主题图，但尚未记录素材；候选提示词和数量不证明图片已嵌入")
        claim = str(page.get("claim", ""))
        if re.search(r"\d", claim) and not page.get("sources") and not page.get("evidence"):
            warnings.append(f"P{i}: 标题含数字但没有 evidence/sources")
    counts = Counter(families)
    # Names and counts cannot establish visual variety; compare actual page previews.
    return errors, warnings, counts

def main():
    ap = argparse.ArgumentParser(description="检查 PPT 规格的叙事与设计连续性")
    ap.add_argument("spec")
    ap.add_argument("--strict", action="store_true")
    args = ap.parse_args()
    errors, warnings, counts = lint(load(args.spec))
    for item in errors:
        print("ERROR", item)
    for item in warnings:
        print("WARN ", item)
    print("LAYOUTS", dict(counts))
    if errors or (args.strict and warnings):
        sys.exit(1)

if __name__ == "__main__":
    main()

