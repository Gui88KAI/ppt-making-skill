#!/usr/bin/env python3
"""template_map.py — 打印一个 PPTX 的结构（每页每个元素的编号、位置、文字、字号）。

用来搞清楚「这套漂亮模板的哪一页、哪个元素，对应我要填的内容」。
配合 fill_template.py 使用：先 map → 写映射 → 填内容。

用法:
    python3 template_map.py 模板.pptx
    python3 template_map.py 模板.pptx --slide 3
    python3 template_map.py 模板.pptx --json > map.json
    python3 template_map.py 模板.pptx --find "示例标题"
"""
import argparse
import json
import sys

from pptx import Presentation
from pptx.util import Emu


def cm(v):
    try:
        return round(Emu(v).cm, 2)
    except Exception:
        return 0


def shape_info(sh):
    info = {
        "idx": None,
        "type": str(sh.shape_type).split(" ")[0],
        "name": sh.name,
        "x": cm(sh.left), "y": cm(sh.top),
        "w": cm(sh.width), "h": cm(sh.height),
        "text": "",
        "sizes": [],
    }
    if sh.has_text_frame:
        txt = sh.text_frame.text.strip()
        info["text"] = txt
        info["sizes"] = sorted({r.font.size.pt for p in sh.text_frame.paragraphs
                                for r in p.runs if r.font.size})
    if sh.shape_type == 13:
        try:
            info["image"] = f"{sh.image.size[0]}x{sh.image.size[1]}"
        except Exception:
            pass
    return info


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--slide", type=int, default=0)
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--find", default="")
    a = ap.parse_args()

    prs = Presentation(a.pptx)
    out = []
    for si, slide in enumerate(prs.slides, 1):
        if a.slide and si != a.slide:
            continue
        shapes = []
        for idx, sh in enumerate(slide.shapes):
            info = shape_info(sh)
            info["idx"] = idx
            shapes.append(info)
        out.append({"slide": si, "layout": slide.slide_layout.name, "shapes": shapes})

    if a.find:
        for s in out:
            for sh in s["shapes"]:
                if a.find in sh["text"]:
                    print(f"P{s['slide']} 元素#{sh['idx']} ({sh['x']},{sh['y']} {sh['w']}x{sh['h']}cm) "
                          f"{sh['sizes']} 「{sh['text'][:40]}」")
        return

    if a.json:
        json.dump(out, sys.stdout, ensure_ascii=False, indent=1)
        return

    for s in out:
        print(f"\n=== 第 {s['slide']} 页（{s['layout']}）===")
        for sh in s["shapes"]:
            tag = ""
            if sh["text"]:
                tag = f"「{sh['text'].replace(chr(10), ' / ')[:36]}」 {sh['sizes']}"
            elif sh.get("image"):
                tag = f"[图片 {sh['image']}]"
            print(f"  #{sh['idx']:<2} {sh['type']:<10} ({sh['x']:>5},{sh['y']:>5}) "
                  f"{sh['w']:>5}x{sh['h']:<5}cm {tag}")


if __name__ == "__main__":
    main()
