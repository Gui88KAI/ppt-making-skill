#!/usr/bin/env python3
"""把视觉圣经和逐页视觉规格编译成可批量生图的 prompts.json。"""
import argparse
import hashlib
import json
from pathlib import Path

VARIATIONS = [
    "documentary and restrained, natural imperfections, believable scale",
    "premium editorial photograph, controlled highlights, refined composition",
    "cinematic but physically plausible, strong depth and a clear visual hierarchy",
]

def stable_seed(base, slide_no, candidate):
    raw = f"{base}:{slide_no}:{candidate}".encode("utf-8")
    return int(hashlib.sha1(raw).hexdigest()[:8], 16)

def compile_prompt(bible, slide, image, candidate):
    b = bible or {}
    parts = [
        "Create one polished presentation asset, not a poster and not an infographic.",
        f"story claim: {slide.get('claim', '')}",
        f"visual anchor: {slide.get('visual_anchor', '')}",
        f"subject identity: {b.get('subject_identity', '')}",
        f"visual metaphor: {b.get('visual_metaphor', '')}",
        f"recurring motif: {b.get('recurring_motif', '')}",
        f"camera language: {b.get('camera_language', '')}",
        f"material and lighting language: {b.get('material_language', '')}",
        f"overall style: {b.get('style', '')}",
        f"palette: {b.get('palette', '')}",
        f"continuity from previous page: {slide.get('continuity_from_previous', '')}",
        f"continuity toward next page: {slide.get('continuity_to_next', '')}",
        f"page-specific scene: {image.get('prompt', '')}",
        f"candidate direction: {VARIATIONS[candidate % len(VARIATIONS)]}",
        "clean space reserved for the text area: " + str(image.get("text_area", "none")),
        "no readable text, no letters, no logo, no watermark, no fake UI, no random people, no unrelated objects, no style drift, no duplicated subject",
    ]
    return ", ".join(x for x in parts if x and not x.endswith(": "))

def main():
    ap = argparse.ArgumentParser(description="编译有上下文的 PPT 配图提示词")
    ap.add_argument("manifest")
    ap.add_argument("-o", "--out", default="prompts.json")
    ap.add_argument("--candidates", type=int, default=3)
    args = ap.parse_args()
    raw = json.loads(Path(args.manifest).read_text(encoding="utf-8"))
    bible = raw.get("visual_bible") or {}
    base_seed = bible.get("seed", bible.get("thesis", "ppt"))
    images = []
    for slide in raw.get("slides", []):
        image = slide.get("image") or {}
        for candidate in range(max(1, args.candidates)):
            spec = dict(image)
            key = spec.get("key", f"p{int(slide.get('number', 0)):02d}_asset")
            spec["key"] = f"{key}_c{candidate + 1}"
            spec["slide"] = slide.get("number")
            spec["claim"] = slide.get("claim", "")
            spec["visual_anchor"] = slide.get("visual_anchor", "")
            spec["continuity_from_previous"] = slide.get("continuity_from_previous", "")
            spec["continuity_to_next"] = slide.get("continuity_to_next", "")
            spec["prompt"] = compile_prompt(bible, slide, image, candidate)
            spec.setdefault("style", image.get("style", bible.get("style", "photo")))
            spec["seed"] = stable_seed(base_seed, slide.get("number", 0), candidate + 1)
            spec["candidate"] = candidate + 1
            images.append(spec)
    out = {
        "theme": bible.get("thesis", ""),
        "palette": bible.get("palette", ""),
        "style": bible.get("style", "photo"),
        "visual_bible": bible,
        "images": images,
        "selection_rule": {
            "minimum_candidates_per_slide": max(1, args.candidates),
            "score_threshold": 78,
            "score_dimensions": ["topic_fit", "subject_consistency", "composition", "lighting", "craft"],
        },
    }
    Path(args.out).write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已生成 {args.out}：{len(images)} 张候选图规格")

if __name__ == "__main__":
    main()
