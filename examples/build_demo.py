#!/usr/bin/env python3
"""build_demo.py — 生成带 AI 配图的示范 PPT（v2 版式）。

在 build_template.py 的基础上增加：
  - 背景图支持（cover / section / 结尾页）
  - 左图右卡片的「痛点页」版式
  - 大图 + 右侧要点的「方案页」版式
"""

import os
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Cm, Pt

W, H = 33.867, 19.05
THEMES = {
    "green": dict(bg1="#0B2318", main="#3FBF6A", accent="#F2C94C", card="#123A28",
                  line="#2E6B4C", text="#EAF6EE", muted="#9BBCA8"),
    "blue": dict(bg1="#081A33", main="#2F80FF", accent="#00D8FF", card="#0F2E52",
                 line="#25517F", text="#E8F1FF", muted="#8FA6C2"),
}


def hx(s):
    s = s.lstrip("#")
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def set_alpha(shape, alpha):
    """给已有填充设置透明度 0–1。"""
    from lxml import etree
    from pptx.oxml.ns import qn
    solid = shape._element.find(".//" + qn("a:solidFill"))
    if solid is None:
        return
    clr = solid.find(qn("a:srgbClr"))
    if clr is None:
        return
    a = etree.SubElement(clr, qn("a:alpha"))
    a.set("val", str(int(alpha * 100000)))


def rect(slide, x, y, w, h, color=None, line=None, radius=0.2, lw=1.0, alpha=1.0):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
        Cm(x), Cm(y), Cm(w), Cm(h))
    if radius:
        try:
            shp.adjustments[0] = min(radius / max(min(w, h), .01), .5)
        except Exception:
            pass
    if color:
        shp.fill.solid()
        shp.fill.fore_color.rgb = hx(color)
        if alpha < 1:
            set_alpha(shp, alpha)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = hx(line)
        shp.line.width = Pt(lw)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def text(slide, s, x, y, w, h, size=18, bold=False, color="#FFFFFF", align="l",
         anchor="t", font="微软雅黑", spacing=1.3):
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    for i, ln in enumerate(str(s).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        p.line_spacing = spacing
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.name = font
        r.font.color.rgb = hx(color)
    return tb


def hline(slide, x, y, w, color, lw=1.0):
    from pptx.enum.shapes import MSO_CONNECTOR
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x), Cm(y), Cm(x + w), Cm(y))
    c.line.color.rgb = hx(color)
    c.line.width = Pt(lw)


def pic_cover(slide, path, x, y, w, h):
    """按 cover 方式放图（等比裁切填满框）。"""
    from PIL import Image
    iw, ih = Image.open(path).size
    p = slide.shapes.add_picture(path, Cm(x), Cm(y), Cm(w), Cm(h))
    box, img = w / h, iw / ih
    if img > box:
        o = 1 - box / img
        p.crop_left = p.crop_right = o / 2
    else:
        o = 1 - img / box
        p.crop_top = p.crop_bottom = o / 2
    return p


def bg_plain(slide, th):
    b = rect(slide, 0, 0, W, H, th["bg1"], radius=0)


def bg_image(slide, th, path, dark=0.45, side=None):
    if not (path and os.path.exists(path)):
        bg_plain(slide, th)
        return
    pic_cover(slide, path, 0, 0, W, H)
    overlay = rect(slide, 0, 0, W, H, th["bg1"], radius=0, alpha=dark)
    if side == "left":   # 左侧再加深一层，给标题让位
        rect(slide, 0, 0, W * 0.55, H, th["bg1"], radius=0, alpha=0.45)


def title_band(slide, th, title, sub, tag):
    rect(slide, 1.1, 0.95, 1.05, 1.05, th["main"], radius=0.22)
    rect(slide, 1.36, 1.21, 0.54, 0.54, th["bg1"], radius=0.12)
    text(slide, title, 2.55, 0.82, 22.0, 1.5, size=38, bold=True, color=th["text"])
    hline(slide, 2.6, 2.34, 3.4, th["main"], 3.0)
    hline(slide, 6.15, 2.34, 1.0, th["accent"], 3.0)
    if tag:
        text(slide, tag, W - 8.0, 1.05, 6.6, 0.9, size=17, color=th["muted"], align="r")
    if sub:
        text(slide, sub, 10.0, 2.64, W - 11.5, 0.9, size=17, color=th["muted"])


def footer(slide, th, page, total, source):
    hline(slide, 1.1, 17.3, W - 2.2, th["line"], 0.75)
    if source:
        text(slide, source, 1.1, 17.5, 22.0, 0.8, size=14, color=th["muted"])
    text(slide, f"{page:02d} / {total:02d}", W - 5.0, 17.45, 3.9, 0.9, size=14,
         color=th["muted"], align="r", font="Arial")
    rect(slide, W - 5.6, 17.6, 0.4, 0.48, th["main"], radius=0)


# ---------------- 页面 ----------------
def page_cover(prs, th, d, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg_image(s, th, d.get("bg_path"), dark=0.42, side="left")
    rect(s, 1.1, 3.0, 0.18, 8.2, th["main"], radius=0)
    text(s, d["kicker"], 1.6, 2.5, 26.0, 0.9, size=17, color=th["accent"])
    text(s, d["title"], 1.6, 3.7, 22.0, 5.2, size=52, bold=True, color=th["text"], spacing=1.16)
    hline(s, 1.65, 9.5, 4.2, th["main"], 3.5)
    text(s, d["subtitle"], 1.6, 10.2, 24.0, 1.6, size=23, color=th["muted"])
    x = 1.6
    for k in d.get("tags", []):
        w = 0.55 * len(k) + 1.3
        rect(s, x, 12.6, w, 1.1, th["main"], radius=0.5, alpha=0.32)
        text(s, k, x, 12.86, w, 0.7, size=16, color=th["accent"], align="c")
        x += w + 0.6
    text(s, d["org"], 1.6, 15.6, 24.0, 0.9, size=14, color=th["muted"])


def page_section(prs, th, d, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg_plain(s, th)
    rect(s, 1.1, 5.0, 0.3, 8.6, th["main"], radius=0)
    text(s, d["number"], 2.1, 4.4, 14.0, 5.4, size=108, bold=True, color=th["main"], font="Arial")
    text(s, d["title"], 2.2, 10.1, 24.0, 2.4, size=42, bold=True, color=th["text"])
    hline(s, 2.25, 12.7, 5.0, th["accent"], 4.0)
    text(s, d.get("note", ""), 2.25, 13.25, 24.0, 1.4, size=20, color=th["muted"])
    text(s, f"{n:02d} / {total:02d}", W - 5.0, 17.45, 3.9, 0.9, size=14,
         color=th["muted"], align="r", font="Arial")


def page_image_cards(prs, th, d, n, total):
    """左图 + 右侧三张卡片（痛点页）"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg_plain(s, th)
    title_band(s, th, d["title"], d.get("sub", ""), d.get("tag", ""))

    # 左侧大图
    ix, iy, iw, ih = 1.1, 4.1, 11.0, 11.6
    pic_cover(s, d["image"], ix, iy, iw, ih)
    rect(s, ix, iy, iw, ih, th["main"], radius=0.05, alpha=0.18)
    rect(s, ix, iy, iw, ih, None, line=th["line"], radius=0.05, lw=1.2)

    # 右侧三张横卡
    cx = ix + iw + 0.9
    cw = W - cx - 1.1
    gap = 0.7
    ch = (11.6 - gap * 2) / 3
    for i, (t, body, foot) in enumerate(d["cards"]):
        y = iy + i * (ch + gap)
        rect(s, cx, y, cw, ch, th["card"], line=th["line"], radius=0.25)
        rect(s, cx, y, 0.14, ch, th["main"], radius=0)
        text(s, f"{i+1:02d}", cx + 0.55, y + 0.42, 1.2, 0.8, size=16, bold=True,
             color=th["accent"], font="Arial")
        text(s, t, cx + 1.85, y + 0.35, cw - 2.6, 1.0, size=23, bold=True, color=th["text"])
        text(s, body, cx + 0.6, y + 1.55, cw - 1.2, ch - 2.6, size=17, color=th["text"], spacing=1.4)
        rect(s, cx + cw - 3.4, y + ch - 1.15, 2.9, 0.75, th["accent"], radius=0.3, alpha=0.2)
        text(s, foot, cx + cw - 3.4, y + ch - 1.0, 2.9, 0.55, size=14,
             color=th["accent"], align="c")
    footer(s, th, n, total, d.get("source", ""))


def page_image_hero(prs, th, d, n, total):
    """大图 + 右侧要点（方案页）"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg_plain(s, th)
    title_band(s, th, d["title"], d.get("sub", ""), d.get("tag", ""))

    # 右侧图（更大）
    iw, ih = 12.5, 11.6
    ix, iy = W - iw - 1.1, 4.1
    pic_cover(s, d["image"], ix, iy, iw, ih)
    rect(s, ix, iy, iw, ih, None, line=th["main"], radius=0.05, lw=1.5)

    # 左侧要点
    tx, tw = 1.1, W - iw - 1.1 - 1.6
    y = 4.5
    for i, p in enumerate(d["points"]):
        rect(s, tx, y + 0.1, 1.0, 1.0, th["main"], radius=0.2)
        text(s, str(i + 1), tx, y + 0.32, 1.0, 0.7, size=22, bold=True,
             color=th["bg1"], align="c", font="Arial")
        text(s, p, tx + 1.3, y, tw - 1.4, 2.8, size=20, color=th["text"], spacing=1.45)
        y += 3.5
    footer(s, th, n, total, d.get("source", ""))


def page_kpi(prs, th, d, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg_image(s, th, d.get("bg_path"), dark=0.55)
    title_band(s, th, d["title"], d.get("sub", ""), d.get("tag", ""))
    items = d["items"]
    gap = 1.2
    cw = (W - 2.2 - gap * (len(items) - 1)) / len(items)
    for i, it in enumerate(items):
        x = 1.1 + i * (cw + gap)
        y = 4.4
        rect(s, x, y, cw, 11.4, th["card"], line=th["main"], radius=0.3, alpha=0.75)
        rect(s, x + 0.3, y + 0.3, cw - 0.6, 0.12, th["main"], radius=0.05)
        text(s, it["value"], x, y + 2.0, cw, 3.4, size=72, bold=True, color=th["accent"],
             align="c", font="Arial")
        text(s, it["unit"], x, y + 6.0, cw, 1.0, size=19, color=th["text"], align="c")
        hline(s, x + cw * 0.3, y + 7.5, cw * 0.4, th["line"], 1.0)
        text(s, it["label"], x + 0.6, y + 8.0, cw - 1.2, 2.0, size=18, color=th["muted"], align="c")
    footer(s, th, n, total, d.get("source", ""))


def page_compare(prs, th, d, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg_plain(s, th)
    title_band(s, th, d["title"], d.get("sub", ""), d.get("tag", ""))
    half = (W - 2.2 - 2.4) / 2
    y, h = 4.2, 11.6
    for side, x, col in (("left", 1.1, th["muted"]), ("right", 1.1 + half + 2.4, th["main"])):
        blk = d[side]
        rect(s, x, y, half, h, th["card"], line=th["line"], radius=0.3)
        rect(s, x + 0.3, y + 0.3, half - 0.6, 0.14, col, radius=0.06)
        text(s, blk["title"], x + 0.9, y + 0.85, half - 1.8, 1.2, size=26, bold=True, color=th["text"])
        hline(s, x + 0.9, y + 2.4, half - 1.8, th["line"], 1.0)
        yy = y + 3.0
        for it in blk["items"]:
            rect(s, x + 0.9, yy + 0.32, 0.3, 0.3, col, radius=0.5)
            text(s, it, x + 1.5, yy, half - 2.4, 1.9, size=19, color=th["text"], spacing=1.3)
            yy += 1.95
    badge = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(W / 2 - 1.35), Cm(y + h / 2 - 1.35), Cm(2.7), Cm(2.7))
    badge.fill.solid(); badge.fill.fore_color.rgb = hx(th["main"])
    badge.line.color.rgb = hx(th["bg1"]); badge.line.width = Pt(3)
    badge.shadow.inherit = False
    text(s, "VS", W / 2 - 1.35, y + h / 2 - 0.5, 2.7, 1.0, size=20, bold=True,
         color=th["bg1"], align="c", font="Arial")
    footer(s, th, n, total, d.get("source", ""))


def page_end(prs, th, d, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg_image(s, th, d.get("bg_path"), dark=0.6)
    text(s, d.get("title", "谢谢"), 0, 6.6, W, 3.6, size=58, bold=True, color=th["text"], align="c")
    rect(s, W / 2 - 1.6, 10.6, 3.2, 0.16, th["main"], radius=0)
    text(s, d.get("note", ""), 2.0, 11.5, W - 4.0, 1.6, size=21, color=th["muted"], align="c")


PAGES = {"cover": page_cover, "section": page_section, "image_cards": page_image_cards,
         "image_hero": page_image_hero, "kpi": page_kpi, "compare": page_compare,
         "end": page_end}


def build(deck, out, assets_dir="assets"):
    th = THEMES["green"]
    prs = Presentation()
    prs.slide_width = Cm(W)
    prs.slide_height = Cm(H)
    pages = deck["pages"]
    for i, p in enumerate(pages, 1):
        p = dict(p)
        if p.get("bg"):
            p["bg_path"] = os.path.join(assets_dir, p["bg"] + ".jpg")
        PAGES[p["type"]](prs, th, p, i, len(pages))
    prs.save(out)
    return out, len(pages)


if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from content import DECK
    out, n = build(DECK, "demo_v2.pptx")
    print(f"已生成 {out}（{n} 页，{os.path.getsize(out)/1024:.0f} KB）")
