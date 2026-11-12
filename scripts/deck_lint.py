#!/usr/bin/env python3
"""检查 PPT 规格中的叙事、视觉、版式和动效连续性。"""
import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path

def load(path):
    raw = Path(path).read_text(encoding="utf-8")
    return json.loads(raw) if path.endswith(".json") else __import__("yaml").safe_load(raw)

def lint(spec):
    pages = spec.get("pages") or []
    errors, warnings = [], []
    families = []
    motion_ids = []
    for i, page in enumerate(pages, 1):
        family = page.get("layout_family") or page.get("layout") or page.get("type", "cards")
        families.append(family)
        required = ["claim", "visual_anchor", "bridge_from_previous", "bridge_to_next"]
        missing = [k for k in required if not page.get(k)]
        if missing:
            warnings.append(f"P{i}: 缺少叙事字段 {', '.join(missing)}")
        motion = page.get("motion") or {}
        ids = set(motion.get("shared_ids") or [])
        motion_ids.append(ids)
        if not motion.get("shared_ids"):
            warnings.append(f"P{i}: 没有 motion.shared_ids，Morph 无法形成连续对象")
        if i > 1 and motion_ids[-2] and ids and not (motion_ids[-2] & ids):
            warnings.append(f"P{i}: 与上一页没有共享运动对象，页面间动画可能断开")
        if len(motion.get("entry_groups") or []) > 3:
            errors.append(f"P{i}: entry_groups 超过 3 组")
        if page.get("images_required", False):
            n = page.get("image_candidates", page.get("candidates", 0))
            if n < 3:
                warnings.append(f"P{i}: 图片候选少于 3 张")
        claim = str(page.get("claim", ""))
        if re.search(r"\d", claim) and not page.get("sources") and not page.get("evidence"):
            warnings.append(f"P{i}: 标题含数字但没有 evidence/sources")
        if i > 1 and family == families[-2]:
            warnings.append(f"P{i}: 与上一页重复版式 {family}")
    counts = Counter(families)
    if len(counts) < min(5, len(pages)):
        warnings.append(f"整套只使用 {len(counts)} 种版式家族，建议至少 5 种")
    for family, count in counts.items():
        if count > 3:
            warnings.append(f"版式 {family} 使用 {count} 次，可能形成模板感")
    if pages and pages[0].get("layout_family") not in ("cover", "quote", "editorial", None):
        warnings.append("P1 建议使用 cover/quote/editorial 作为开场锚点")
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
