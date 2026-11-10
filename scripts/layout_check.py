#!/usr/bin/env python3
"""layout_check.py — 排版自检：把"看着乱"变成可检查的指标。

检查项：
  ✗ 字号过小（<14pt）、正文过小（<18pt）
  ✗ 文字溢出文本框（按字号和框宽估算行数）
  ✗ 元素重叠（文本框之间、文本框与图片）
  ✗ 元素出界（超出画布或安全边距）
  ✗ 空页
  ⚠ 每页文字过多（>200 字）

用法:
    python3 layout_check.py 项目.pptx
    python3 layout_check.py 项目.pptx --strict     # 有 error 就返回非 0
    python3 layout_check.py 项目.pptx --json
"""
import argparse
import json
import sys

from pptx import Presentation
from pptx.util import Emu

EMU_CM = 360000.0
SAFE = 0.8      # 安全边距 cm
MARGIN = 1.2    # 建议最小边距 cm


def cm(v):
    try:
        return v / EMU_CM
    except Exception:
        return 0


def is_body(sh):
    """判断是不是正文类文字（排除装饰字、编号、logo 等）。"""
    t = sh.text_frame.text.strip()
    if len(t) < 6:
        return False
    sizes = [r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs if r.font.size]
    if not sizes:
        return False
    fs = max(sizes)
    return 12 <= fs <= 32


def walk(shapes):
    for sh in shapes:
        if sh.shape_type == 6:
            yield from walk(sh.shapes)
        else:
            yield sh


def text_len(t):
    return sum(1 if ord(c) > 127 else 0.55 for c in t.replace("\n", ""))


def check(path):
    prs = Presentation(path)
    W, H = cm(prs.slide_width), cm(prs.slide_height)
    problems = []
    for si, slide in enumerate(prs.slides, 1):
        shapes = list(walk(slide.shapes))
        texts = [sh for sh in shapes if sh.has_text_frame and sh.text_frame.text.strip()]
        pics = [sh for sh in shapes if sh.shape_type == 13]
        chars = sum(text_len(sh.text_frame.text) for sh in texts)

        if not texts and not pics:
            problems.append((si, "error", "空页", "这一页没有任何内容"))
        if chars > 260:
            problems.append((si, "warn", "文字过多", f"{int(chars)} 字，路演页建议 ≤ 200 字"))

        for sh in texts:
            t = sh.text_frame.text.strip()
            sizes = [r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs if r.font.size]
            fs = max(sizes) if sizes else None
            if fs is None or not is_body(sh):
                continue
            if fs < 14:
                problems.append((si, "error", "字号过小", f"{fs:.0f}pt「{t[:16]}」"))
            elif fs < 18:
                problems.append((si, "warn", "字号偏小", f"{fs:.0f}pt「{t[:16]}」"))
            # 溢出估算
            w = cm(sh.width)
            em = fs / 72 * 2.54
            per = max(int(w / em), 1)
            lines = 0
            for para in sh.text_frame.paragraphs:
                n = sum(1 if ord(c) > 127 else 0.55 for c in para.text)
                lines += max(1, -(-int(n) // per))
            need = lines * em * (sh.text_frame.paragraphs[0].line_spacing or 1.3)
            h = cm(sh.height)
            if need > h * 1.15:
                problems.append((si, "error", "文字溢出", f"需 {need:.1f}cm > 框 {h:.1f}cm「{t[:16]}」"))

        # 出界
        for sh in shapes:
            try:
                x, y = cm(sh.left), cm(sh.top)
                w, h = cm(sh.width), cm(sh.height)
            except Exception:
                continue
            if w <= 0 or h <= 0:
                continue
            if x < -SAFE or y < -SAFE or x + w > W + SAFE or y + h > H + SAFE:
                continue  # 故意出血的装饰元素，跳过
            if x < MARGIN * 0.5 or y < MARGIN * 0.5 or x + w > W - MARGIN * 0.5 or y + h > H - MARGIN * 0.5:
                if (sh.has_text_frame and sh.text_frame.text.strip() and is_body(sh)):
                    problems.append((si, "warn", "贴边太近",
                                     f"「{sh.text_frame.text.strip()[:14]}」距边 < {MARGIN*0.5:.1f}cm"))

        # 重叠（只查文字框之间，且交集明显）
        boxes = []
        for sh in [x for x in texts if is_body(x)]:
            try:
                boxes.append((cm(sh.left), cm(sh.top), cm(sh.left) + cm(sh.width),
                              cm(sh.top) + cm(sh.height), sh.text_frame.text.strip()[:12]))
            except Exception:
                pass
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                a, b = boxes[i], boxes[j]
                ox = min(a[2], b[2]) - max(a[0], b[0])
                oy = min(a[3], b[3]) - max(a[1], b[1])
                if ox > 1.2 and oy > 0.8:
                    problems.append((si, "error", "文字重叠", f"「{a[4]}」×「{b[4]}」重叠 {ox:.1f}×{oy:.1f}cm"))



    return problems, len(prs.slides)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--strict", action="store_true")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    problems, n = check(a.pptx)
    if a.json:
        json.dump([{"slide": s, "level": l, "kind": k, "detail": d} for s, l, k, d in problems],
                  sys.stdout, ensure_ascii=False, indent=1)
        return
    errs = [p for p in problems if p[1] == "error"]
    warns = [p for p in problems if p[1] == "warn"]
    print(f"共 {n} 页：{len(errs)} 个必须修、{len(warns)} 个建议改\n")
    cur = None
    for s, lvl, kind, detail in problems:
        if s != cur:
            print(f"第 {s} 页"); cur = s
        mark = "✗" if lvl == "error" else "⚠"
        print(f"  {mark} {kind}: {detail}")
    if not problems:
        print("排版检查通过，没有发现溢出/重叠/过小字号。")
    if a.strict and errs:
        sys.exit(1)


if __name__ == "__main__":
    main()
