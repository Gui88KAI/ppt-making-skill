#!/usr/bin/env python3
"""deck_builder.py — 用 YAML/JSON 规格从零生成一套路演 PPT（元素化拼装）。

用法:
    python3 deck_builder.py spec.yaml -o out.pptx
    python3 deck_builder.py examples/demo.yaml -o demo.pptx
    python3 deck_builder.py spec.yaml -o out.pptx --check     # 生成后自动跑体检

规格格式见 examples/demo.yaml。设计规则见 references/元素化制作法.md。

依赖: pip install python-pptx pyyaml
"""
import argparse
import json
import os
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Cm, Pt

# ---------- 画布与网格（16:9） ----------
W, H = 33.867, 19.05
MARGIN = 1.5
CONTENT_W = W - MARGIN * 2
TITLE_Y = 1.15
BODY_Y = 3.9
FOOTER_Y = 17.55

DEFAULT_THEME = {
    "main": "#0B5FFF",
    "secondary": "#0A2540",
    "accent": "#00D4FF",
    "bg": "#FFFFFF",
    "text": "#2B2B2B",
    "muted": "#8A8F98",
    "font_title": "思源黑体 Heavy",
    "font_body": "思源黑体",
    "font_num": "DIN",
}

ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}


def hexc(s):
    s = s.lstrip("#")
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def theme_of(spec):
    t = dict(DEFAULT_THEME)
    t.update(spec.get("theme") or {})
    return t


def blank(prs, theme):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill
    bg.solid()
    bg.fore_color.rgb = hexc(theme["bg"])
    return s


# ---------- 元素 ----------
def text(slide, s, x, y, w, h, size=20, bold=False, color="#2B2B2B",
         align="left", font="思源黑体", spacing=1.3, anchor="t", italic=False):
    """文本元素。x/y/w/h 单位 cm。"""
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE,
                          "b": MSO_ANCHOR.BOTTOM}[anchor]
    lines = str(s).split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ALIGN[align]
        p.line_spacing = spacing
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.name = font
        r.font.color.rgb = hexc(color)
    return tb


def rect(slide, x, y, w, h, fill, line=None, radius=0.25, shadow=False):
    """色块元素。"""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
                                   Cm(x), Cm(y), Cm(w), Cm(h))
    if radius:
        try:
            shape.adjustments[0] = min(radius, 0.5)
        except Exception:
            pass
    shape.fill.solid()
    shape.fill.fore_color.rgb = hexc(fill)
    if line:
        shape.line.color.rgb = hexc(line)
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    if not shadow:
        shape.shadow.inherit = False
    return shape


def line(slide, x1, y1, x2, y2, color="#D0D4DA", width=1.0):
    from pptx.enum.shapes import MSO_CONNECTOR
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x1), Cm(y1), Cm(x2), Cm(y2))
    conn.line.color.rgb = hexc(color)
    conn.line.width = Pt(width)
    return conn


def picture(slide, path, x, y, w, h, radius=0.2):
    """图片元素：等比裁切填充（cover）。缺图时画灰块占位。"""
    if not path or not os.path.exists(path):
        rect(slide, x, y, w, h, "#EDEFF2", radius=radius)
        text(slide, "图片占位\n" + os.path.basename(str(path or "未提供")),
             x, y + h / 2 - 0.6, w, 1.2, size=12, color="#9AA0A6",
             align="center", anchor="m")
        return None
    from PIL import Image
    iw, ih = Image.open(path).size
    box_ratio = w / h
    img_ratio = iw / ih
    pic = slide.shapes.add_picture(path, Cm(x), Cm(y), Cm(w), Cm(h))
    if img_ratio > box_ratio:      # 图更宽 → 左右裁
        over = 1 - box_ratio / img_ratio
        pic.crop_left = pic.crop_right = over / 2
    else:                          # 图更高 → 上下裁
        over = 1 - img_ratio / box_ratio
        pic.crop_top = pic.crop_bottom = over / 2
    return pic


def page_title(slide, s, theme, sub=None):
    text(slide, s, MARGIN, TITLE_Y, CONTENT_W, 1.6, size=34, bold=True,
         color=theme["secondary"], font=theme["font_title"], spacing=1.1)
    rect(slide, MARGIN, TITLE_Y + 1.7, 1.6, 0.12, theme["main"], radius=0)
    if sub:
        text(slide, sub, MARGIN, TITLE_Y + 2.0, CONTENT_W, 1.0, size=17,
             color=theme["muted"], font=theme["font_body"])


def footer(slide, theme, num, total, source=None):
    line(slide, MARGIN, FOOTER_Y, W - MARGIN, FOOTER_Y, "#E3E6EA", 0.75)
    if source:
        text(slide, source, MARGIN, FOOTER_Y + 0.15, CONTENT_W - 4, 0.8,
             size=10, color=theme["muted"], font=theme["font_body"])
    text(slide, f"{num}/{total}", W - MARGIN - 3, FOOTER_Y + 0.15, 3, 0.8,
         size=10, color=theme["muted"], align="right", font=theme["font_body"])


# ---------- 页面类型 ----------
def p_cover(prs, theme, spec, ctx):
    s = blank(prs, theme)
    has_img = spec.get("image")
    if has_img:
        picture(s, spec["image"], 0, 0, W, H, radius=0)
        rect(s, 0, 0, W, H, theme["secondary"], radius=0)
        s.shapes[-1].fill.transparency = 0.35
    text(s, spec.get("kicker", ""), MARGIN, 6.2, CONTENT_W, 1.0, size=18,
         color=theme["accent"], font=theme["font_body"])
    text(s, spec.get("title", ""), MARGIN, 7.2, CONTENT_W - 4, 4.2, size=54, bold=True,
         color="#FFFFFF" if has_img else theme["secondary"], font=theme["font_title"], spacing=1.15)
    rect(s, MARGIN, 11.9, 3.2, 0.16, theme["main"], radius=0)
    text(s, spec.get("subtitle", ""), MARGIN, 12.6, CONTENT_W - 6, 2.0, size=22,
         color="#E8EAED" if has_img else theme["text"], font=theme["font_body"])
    text(s, spec.get("org", ""), MARGIN, 15.6, CONTENT_W, 0.9, size=14,
         color="#E8EAED" if has_img else theme["muted"], font=theme["font_body"])


def p_section(prs, theme, spec, ctx):
    s = blank(prs, theme)
    rect(s, 0, 0, W, H, theme["secondary"], radius=0)
    text(s, spec.get("number", ""), MARGIN, 5.0, 12, 5.0, size=140, bold=True,
         color=theme["main"], font=theme["font_num"], spacing=1.0)
    text(s, spec.get("title", ""), MARGIN, 11.0, CONTENT_W, 3.0, size=44, bold=True,
         color="#FFFFFF", font=theme["font_title"])
    if spec.get("note"):
        text(s, spec["note"], MARGIN, 13.6, CONTENT_W - 8, 1.6, size=20,
             color="#B9BFC7", font=theme["font_body"])


def p_cards(prs, theme, spec, ctx):
    s = blank(prs, theme)
    page_title(s, spec.get("title", ""), theme, spec.get("sub"))
    items = spec.get("items", [])[:6]
    cols = int(spec.get("cols", 2))
    rows = (len(items) + cols - 1) // cols
    gap = 0.8
    cw = (CONTENT_W - gap * (cols - 1)) / cols
    ch = (14.2 - gap * (rows - 1)) / rows
    for i, it in enumerate(items):
        r, c = divmod(i, cols)
        x = MARGIN + c * (cw + gap)
        y = BODY_Y + r * (ch + gap)
        rect(s, x, y, cw, ch, "#F7F8FA", radius=0.2)
        rect(s, x, y, 0.18, ch, theme["main"], radius=0)
        text(s, it.get("title", ""), x + 0.9, y + 0.7, cw - 1.5, 1.2, size=24, bold=True,
             color=theme["secondary"], font=theme["font_title"])
        text(s, it.get("note", ""), x + 0.9, y + 2.1, cw - 1.5, ch - 2.4, size=17,
             color=theme["text"], font=theme["font_body"], spacing=1.35)
        if it.get("mark"):
            text(s, it["mark"], x + cw - 2.2, y + 0.55, 1.6, 1.0, size=22, bold=True,
                 color=theme["accent"], align="right", font=theme["font_num"])
    footer(s, theme, ctx["n"], ctx["total"], spec.get("source"))


def p_bullets(prs, theme, spec, ctx):
    s = blank(prs, theme)
    page_title(s, spec.get("title", ""), theme, spec.get("sub"))
    items = spec.get("items", [])
    y = BODY_Y + 0.4
    for it in items[:5]:
        txt = it if isinstance(it, str) else it.get("text", "")
        rect(s, MARGIN + 0.1, y + 0.32, 0.34, 0.34, theme["main"], radius=0.5)
        text(s, txt, MARGIN + 1.1, y, CONTENT_W - 1.2, 3.2, size=22,
             color=theme["text"], font=theme["font_body"], spacing=1.45)
        y += 2.7
    footer(s, theme, ctx["n"], ctx["total"], spec.get("source"))


def p_kpi(prs, theme, spec, ctx):
    s = blank(prs, theme)
    page_title(s, spec.get("title", ""), theme, spec.get("sub"))
    items = spec.get("items", [])[:3]
    n = len(items) or 1
    gap = 1.0
    cw = (CONTENT_W - gap * (n - 1)) / n
    for i, it in enumerate(items):
        x = MARGIN + i * (cw + gap)
        rect(s, x, BODY_Y, cw, 10.6, "#F7F8FA", radius=0.2)
        text(s, it.get("value", ""), x + 0.5, BODY_Y + 1.6, cw - 1.0, 4.4, size=76, bold=True,
             color=theme["main"], align="center", font=theme["font_num"])
        text(s, it.get("unit", ""), x + 0.5, BODY_Y + 6.1, cw - 1.0, 1.0, size=18,
             color=theme["muted"], align="center", font=theme["font_body"])
        text(s, it.get("label", ""), x + 0.7, BODY_Y + 7.4, cw - 1.4, 2.4, size=18,
             color=theme["text"], align="center", font=theme["font_body"])
    footer(s, theme, ctx["n"], ctx["total"], spec.get("source"))


def p_timeline(prs, theme, spec, ctx):
    s = blank(prs, theme)
    page_title(s, spec.get("title", ""), theme, spec.get("sub"))
    items = spec.get("items", [])[:5]
    n = len(items) or 1
    axis_y = BODY_Y + 5.4
    line(s, MARGIN + 1.5, axis_y, W - MARGIN - 1.5, axis_y, theme["main"], 2.0)
    seg = (CONTENT_W - 3) / n
    for i, it in enumerate(items):
        cx = MARGIN + 1.5 + seg * (i + 0.5)
        circle = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(cx - 0.42), Cm(axis_y - 0.42), Cm(0.84), Cm(0.84))
        circle.fill.solid()
        circle.fill.fore_color.rgb = hexc(theme["main"])
        circle.line.color.rgb = hexc(theme["bg"])
        circle.line.width = Pt(3)
        up = i % 2 == 0
        ty = axis_y - 5.2 if up else axis_y + 1.2
        text(s, it.get("label", ""), cx - seg / 2 + 0.2, ty, seg - 0.4, 1.0, size=15, bold=True,
             color=theme["accent"], align="center", font=theme["font_num"])
        text(s, it.get("title", ""), cx - seg / 2 + 0.2, ty + 1.0, seg - 0.4, 1.3, size=20, bold=True,
             color=theme["secondary"], align="center", font=theme["font_title"])
        text(s, it.get("note", ""), cx - seg / 2 + 0.2, ty + 2.4, seg - 0.4, 2.0, size=13,
             color=theme["muted"], align="center", font=theme["font_body"], spacing=1.25)
    footer(s, theme, ctx["n"], ctx["total"], spec.get("source"))


def p_compare(prs, theme, spec, ctx):
    s = blank(prs, theme)
    page_title(s, spec.get("title", ""), theme, spec.get("sub"))
    half = (CONTENT_W - 2.4) / 2
    for side, x, color, bgc in (("left", MARGIN, theme["main"], "#F2F7FF"),
                                ("right", MARGIN + half + 2.4, theme["muted"], "#F7F8FA")):
        block = spec.get(side) or {}
        rect(s, x, BODY_Y, half, 11.0, bgc, radius=0.2)
        text(s, block.get("title", ""), x + 1.0, BODY_Y + 0.9, half - 2.0, 1.4, size=24, bold=True,
             color=color, font=theme["font_title"])
        y = BODY_Y + 2.9
        for it in (block.get("items") or [])[:4]:
            text(s, "· " + str(it), x + 1.0, y, half - 2.0, 2.2, size=18,
                 color=theme["text"], font=theme["font_body"], spacing=1.35)
            y += 2.2
    vs = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(W / 2 - 1.0), Cm(BODY_Y + 4.6), Cm(2.0), Cm(2.0))
    vs.fill.solid()
    vs.fill.fore_color.rgb = hexc(theme["secondary"])
    vs.line.fill.background()
    tf = vs.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    r = tf.paragraphs[0].add_run()
    r.text = "VS"
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.name = theme["font_num"]
    r.font.color.rgb = hexc("#FFFFFF")
    footer(s, theme, ctx["n"], ctx["total"], spec.get("source"))


def p_image_text(prs, theme, spec, ctx):
    s = blank(prs, theme)
    img_w = int(spec.get("image_width", 45)) / 100 * CONTENT_W
    y0 = 0 if spec.get("full_bleed") else TITLE_Y
    picture(s, spec.get("image"), 0, y0, img_w, H if not spec.get("full_bleed") else H, radius=0)
    tx = img_w + 2.2
    tw = W - tx - MARGIN
    text(s, spec.get("title", ""), tx, 3.4, tw, 3.0, size=32, bold=True,
         color=theme["secondary"], font=theme["font_title"], spacing=1.2)
    y = 7.0
    for it in (spec.get("items") or [])[:4]:
        rect(s, tx, y + 0.3, 0.3, 0.3, theme["main"], radius=0.5)
        text(s, str(it), tx + 0.85, y, tw - 0.9, 2.6, size=19,
             color=theme["text"], font=theme["font_body"], spacing=1.4)
        y += 2.5
    footer(s, theme, ctx["n"], ctx["total"], spec.get("source"))


def p_closing(prs, theme, spec, ctx):
    s = blank(prs, theme)
    rect(s, 0, 0, W, H, theme["secondary"], radius=0)
    text(s, spec.get("title", "谢谢"), MARGIN, 7.0, CONTENT_W, 3.4, size=64, bold=True,
         color="#FFFFFF", align="center", font=theme["font_title"])
    if spec.get("note"):
        text(s, spec["note"], MARGIN, 11.0, CONTENT_W, 2.0, size=22,
             color="#B9BFC7", align="center", font=theme["font_body"])
    rect(s, W / 2 - 1.6, 10.2, 3.2, 0.16, theme["main"], radius=0)


BUILDERS = {
    "cover": p_cover, "section": p_section, "cards": p_cards, "bullets": p_bullets,
    "kpi": p_kpi, "timeline": p_timeline, "compare": p_compare,
    "image_text": p_image_text, "three_cols": lambda p, t, s, c: p_cards(p, t, dict(s, cols=3), c),
    "closing": p_closing,
}


def build(spec, out):
    theme = theme_of(spec)
    prs = Presentation()
    prs.slide_width = Cm(W)
    prs.slide_height = Cm(H)
    pages = spec.get("pages") or []
    for i, page in enumerate(pages, 1):
        kind = page.get("type", "cards")
        fn = BUILDERS.get(kind)
        if not fn:
            raise SystemExit(f"未知页面类型: {kind}（可用: {', '.join(BUILDERS)}）")
        ctx = {"n": i, "total": len(pages)}
        fn(prs, theme, page, ctx)
    os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
    prs.save(out)
    return out, len(pages)


def main():
    ap = argparse.ArgumentParser(description="用 YAML/JSON 规格生成 PPT")
    ap.add_argument("spec")
    ap.add_argument("-o", "--out", default="out.pptx")
    ap.add_argument("--check", action="store_true", help="生成后自动跑 pptx_inspect")
    args = ap.parse_args()

    raw = open(args.spec, encoding="utf-8").read()
    spec = json.loads(raw) if args.spec.endswith(".json") else __import__("yaml").safe_load(raw)
    out, n = build(spec, args.out)
    size = os.path.getsize(out) / 1024
    print(f"已生成 {out}（{n} 页，{size:.0f} KB）")
    if args.check:
        import subprocess
        here = os.path.dirname(os.path.abspath(__file__))
        subprocess.run([sys.executable, os.path.join(here, "pptx_inspect.py"), out])


if __name__ == "__main__":
    main()
