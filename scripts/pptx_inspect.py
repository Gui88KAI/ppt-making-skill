#!/usr/bin/env python3
"""pptx_inspect.py — 快速体检一个 PPTX：页数、版式、文字、字体、字号。

用法:
    python3 pptx_inspect.py 模板.pptx
    python3 pptx_inspect.py 模板.pptx --text      # 打印每页完整文字
    python3 pptx_inspect.py 模板.pptx --fonts     # 只列用到的字体

依赖: pip install python-pptx
"""
import sys
import collections

try:
    from pptx import Presentation
    from pptx.util import Emu
except ImportError:
    print("缺少 python-pptx，先执行: pip install python-pptx")
    sys.exit(1)

PLACEHOLDER_HINTS = ("点击输入", "单击此处", "Lorem ipsum", "Click to add", "文本框")


def shape_texts(slide):
    out = []
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            out.append(sh.text_frame.text.strip())
        if getattr(sh, "has_table", False) and sh.has_table:
            for row in sh.table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        out.append(cell.text.strip())
    return out


def collect_fonts(prs):
    fonts = collections.Counter()
    sizes = collections.Counter()
    for slide in prs.slides:
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                for run in p.runs:
                    if run.font.name:
                        fonts[run.font.name] += 1
                    if run.font.size:
                        sizes[round(Emu(run.font.size).pt)] += 1
    return fonts, sizes


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    path = sys.argv[1]
    show_text = "--text" in sys.argv
    only_fonts = "--fonts" in sys.argv

    prs = Presentation(path)
    w_emu, h_emu = prs.slide_width, prs.slide_height
    ratio = "16:9" if abs(w_emu / h_emu - 16 / 9) < 0.02 else "4:3" if abs(w_emu / h_emu - 4 / 3) < 0.02 else f"{w_emu/h_emu:.2f}:1"
    print(f"文件: {path}")
    print(f"尺寸: {Emu(w_emu).cm:.1f} x {Emu(h_emu).cm:.1f} cm ({ratio})")
    print(f"页数: {len(prs.slides)}")

    fonts, sizes = collect_fonts(prs)
    if fonts:
        print("\n字体使用次数:")
        for name, cnt in fonts.most_common():
            print(f"  {name}: {cnt}")
    if sizes:
        print("\n字号分布(pt):")
        for size, cnt in sorted(sizes.items()):
            print(f"  {size}pt: {cnt} 处")
    tiny = [s for s in sizes if s < 14]
    small = [s for s in sizes if 14 <= s < 18]
    if tiny:
        print(f"\n⚠ 有 {sum(sizes[s] for s in tiny)} 处字号小于 14pt，投影几乎看不清")
    if small:
        print(f"提示: {sum(sizes[s] for s in small)} 处 14–17pt，用于注释/来源是正常的，正文应 ≥18pt")
    if only_fonts:
        return

    warnings = []
    print("\n逐页:")
    for idx, slide in enumerate(prs.slides, 1):
        layout = slide.slide_layout.name
        texts = shape_texts(slide)
        pics = sum(1 for sh in slide.shapes if sh.shape_type == 13)  # PICTURE
        total_chars = sum(len(t) for t in texts)
        preview = " / ".join(t.replace("\n", " ")[:24] for t in texts[:2])
        print(f"  P{idx:<3} [{layout}] 文字{total_chars:>4}字 图{pics}张  {preview}")
        if total_chars > 200:
            warnings.append(f"P{idx} 文字 {total_chars} 字，路演页偏多，考虑拆页")
        for t in texts:
            if any(h in t for h in PLACEHOLDER_HINTS):
                warnings.append(f"P{idx} 残留占位符: {t[:20]}")
        if show_text:
            for t in texts:
                print(f"        | {t}")

    if warnings:
        print("\n检查提示:")
        for w in warnings:
            print("  ⚠", w)
    else:
        print("\n检查提示: 未发现明显问题（占位符/字数）")


if __name__ == "__main__":
    main()
