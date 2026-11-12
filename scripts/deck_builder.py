#!/usr/bin/env python3
"""deck_builder.py (v2) — 用 YAML/JSON 规格生成一套路演 PPT。

v2 相对 v1 的改动：
  1. 卡片高度随内容自适应，不再出现大片空白
  2. 标题装饰条与副标题分离，不再压字
  3. 内容区垂直居中，排版不再"上挤下空"
  4. 新增视觉元素：编号圆标、细分割线、浅色底纹、装饰圆环
  5. 封面/章节页有构图（左侧文字 + 右侧抽象图形）
  6. 所有文字块按字数估算行数，尽量不溢出

用法:
    python3 deck_builder.py spec.yaml -o out.pptx
    python3 deck_builder.py spec.yaml -o out.pptx --check

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

# ---------- 画布 ----------
W, H = 33.867, 19.05
M = 1.5                      # 页边距
CW = W - M * 2
TITLE_Y = 0.95
SUB_Y = 2.55
BODY_Y = 4.35
BODY_H = 12.4
FOOTER_Y = 17.6

DEFAULT_THEME = {
    "main": "#0B5FFF",
    "secondary": "#0A2540",
    "accent": "#00D4FF",
    "bg": "#FFFFFF",
    "text": "#2E3440",
    "muted": "#7B8494",
    "line": "#E6E9EF",
    "card": "#F7F9FC",
    "font_title": "微软雅黑",
    "font_body": "微软雅黑",
    "font_num": "Arial",
}

ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}


# ---------- 颜色工具 ----------
def hexc(s):
    s = str(s).lstrip("#")
    if len(s) == 3:
        s = "".join(c * 2 for c in s)
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def _rgb(s):
    s = str(s).lstrip("#")
    return tuple(int(s[i:i + 2], 16) for i in (0, 2, 4))


def _hex(t):
    return "#%02X%02X%02X" % t


def mix(c1, c2, r):
    """把 c1 往 c2 混 r（0=全 c1，1=全 c2）"""
    a, b = _rgb(c1), _rgb(c2)
    return _hex(tuple(int(a[i] + (b[i] - a[i]) * r) for i in range(3)))


def theme_of(spec):
    t = dict(DEFAULT_THEME)
    t.update(spec.get("theme") or {})
    t["tint"] = mix(t["main"], "#FFFFFF", 0.90)      # 主色 10%
    t["tint2"] = mix(t["main"], "#FFFFFF", 0.82)
    t["dark2"] = mix(t["secondary"], "#FFFFFF", 0.10)
    return t


# ---------- 基础元素 ----------
def blank(prs, theme, bg=None):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    f = s.background.fill
    f.solid()
    f.fore_color.rgb = hexc(bg or theme["bg"])
    return s


def text(slide, s, x, y, w, h, size=20, bold=False, color="#2E3440",
         align="left", font=None, spacing=1.3, anchor="t", theme=None):
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE,
                          "b": MSO_ANCHOR.BOTTOM}[anchor]
    for i, line in enumerate(str(s).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ALIGN[align]
        p.line_spacing = spacing
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.name = font or (theme["font_body"] if theme else "思源黑体")
        r.font.color.rgb = hexc(color)
    return tb


def rect(slide, x, y, w, h, fill, line=None, radius=0.2, lw=1.0):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
        Cm(x), Cm(y), Cm(w), Cm(h))
    if radius:
        try:
            shp.adjustments[0] = min(radius / max(min(w, h), 0.01), 0.5)
        except Exception:
            pass
    shp.fill.solid()
    shp.fill.fore_color.rgb = hexc(fill)
    if line:
        shp.line.color.rgb = hexc(line)
        shp.line.width = Pt(lw)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def circle(slide, x, y, d, fill, line=None, lw=2.0):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Cm(x), Cm(y), Cm(d), Cm(d))
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = hexc(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = hexc(line)
        shp.line.width = Pt(lw)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def hline(slide, x, y, w, color="#E6E9EF", lw=1.0):
    from pptx.enum.shapes import MSO_CONNECTOR
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x), Cm(y), Cm(x + w), Cm(y))
    c.line.color.rgb = hexc(color)
    c.line.width = Pt(lw)
    return c


def picture(slide, path, x, y, w, h, radius=0.25):
    if not path or not os.path.exists(path):
        rect(slide, x, y, w, h, "#EEF1F5", line="#E1E5EB", radius=radius)
        text(slide, "图片待补", x, y + h / 2 - 0.55, w, 1.1, size=13, color="#9AA3B2",
             align="center", anchor="m")
        return None
    from PIL import Image
    iw, ih = Image.open(path).size
    br, ir = w / h, iw / ih
    pic = slide.shapes.add_picture(path, Cm(x), Cm(y), Cm(w), Cm(h))
    if ir > br:
        over = 1 - br / ir
        pic.crop_left = pic.crop_right = over / 2
    else:
        over = 1 - ir / br
        pic.crop_top = pic.crop_bottom = over / 2
    # 细边框
    rect(slide, x, y, w, h, "#FFFFFF", line="#E6E9EF", radius=radius, lw=0.75).fill.background()
    return pic


# ---------- 文本测量 ----------
def est_lines(s, size_pt, width_cm):
    """估算中文文本在给定宽度下占几行"""
    em = size_pt / 72 * 2.54
    per = max(int(width_cm / em), 1)
    total = 0
    for line in str(s).split("\n"):
        n = sum(1 if ord(c) > 127 else 0.55 for c in line)
        total += max(1, -(-int(n) // per))
    return total


def est_height(s, size_pt, width_cm, spacing=1.35):
    em = size_pt / 72 * 2.54
    return est_lines(s, size_pt, width_cm) * em * spacing


# ---------- 页头 / 页脚 ----------
def page_head(slide, spec, theme):
    rect(slide, M, TITLE_Y + 0.12, 1.15, 0.42, theme["main"], radius=0)  # 顶部色块
    text(slide, spec.get("title", ""), M + 1.55, TITLE_Y, CW - 1.55, 1.6, size=31, bold=True,
         color=theme["secondary"], font=theme["font_title"], spacing=1.12)
    if spec.get("sub"):
        text(slide, spec["sub"], M + 1.58, SUB_Y, CW - 1.58, 0.9, size=18,
             color=theme["muted"], font=theme["font_body"])


def page_foot(slide, theme, n, total, spec=None):
    hline(slide, M, FOOTER_Y, CW, theme["line"], 0.75)
    if spec and spec.get("source"):
        text(slide, spec["source"], M, FOOTER_Y + 0.22, CW - 4, 0.8, size=10,
             color=theme["muted"], font=theme["font_body"])
    text(slide, f"{n:02d} / {total:02d}", W - M - 3.2, FOOTER_Y + 0.22, 3.2, 0.8, size=10,
         color=theme["muted"], align="right", font=theme["font_body"])


# ---------- 页面类型 ----------
def p_cover(prs, theme, spec, ctx):
    s = blank(prs, theme)
    # 右侧抽象构图
    circle(s, W * 0.63, -3.2, 16.5, None, line=theme["tint2"], lw=2.0)
    circle(s, W * 0.70, 1.6, 10.5, theme["tint"], line=None)
    circle(s, W * 0.80, 8.8, 5.2, theme["main"], line=None)
    circle(s, W * 0.635, 12.4, 2.2, theme["accent"], line=None)
    rect(s, 0, H - 0.55, W, 0.55, theme["secondary"], radius=0)
    text(s, spec.get("kicker", ""), M, 5.0, 18, 1.0, size=17, color=theme["main"],
         font=theme["font_body"])
    text(s, spec.get("title", ""), M, 6.2, 19.5, 4.6, size=46, bold=True,
         color=theme["secondary"], font=theme["font_title"], spacing=1.2)
    rect(s, M, 11.4, 2.6, 0.14, theme["main"], radius=0)
    text(s, spec.get("subtitle", ""), M, 12.1, 19.5, 2.2, size=19,
         color=theme["text"], font=theme["font_body"], spacing=1.35)
    text(s, spec.get("org", ""), M, 15.4, 18, 0.9, size=16, color=theme["muted"],
         font=theme["font_body"])


def p_section(prs, theme, spec, ctx):
    s = blank(prs, theme, theme["secondary"])
    circle(s, W - 9.5, -3.8, 13.5, None, line=theme["dark2"], lw=2.0)
    circle(s, W - 5.2, 10.5, 6.5, None, line=theme["main"], lw=1.5)
    rect(s, M, 5.6, 0.3, 7.6, theme["main"], radius=0)
    text(s, spec.get("number", ""), M + 1.3, 4.35, 12, 5.6, size=100, bold=True,
         color=theme["main"], font=theme["font_num"], spacing=1.0)
    text(s, spec.get("title", ""), M + 1.3, 9.9, CW - 6, 2.6, size=40, bold=True,
         color="#FFFFFF", font=theme["font_title"])
    if spec.get("note"):
        text(s, spec["note"], M + 1.35, 12.9, CW - 10, 1.6, size=20,
             color="#AEB7C4", font=theme["font_body"])
    text(s, f"{ctx['n']:02d} / {ctx['total']:02d}", W - M - 3.2, FOOTER_Y, 3.2, 0.8,
         size=10, color="#6E7A8C", align="right", font=theme["font_body"])


def p_cards(prs, theme, spec, ctx):
    s = blank(prs, theme)
    page_head(s, spec, theme)
    items = (spec.get("items") or [])[:6]
    cols = int(spec.get("cols", 2))
    n = len(items)
    rows = max(1, (n + cols - 1) // cols)
    gap = 0.9
    cw = (CW - gap * (cols - 1)) / cols
    # 自适应卡片高度：内容少的卡片撑高一点，内容多的给足空间
    ideal = 0
    for it in items:
        h_in = 1.7 + est_height(it.get("note", ""), 18, cw - 2.0) + 0.9
        ideal = max(ideal, h_in)
    min_h = 6.6 if rows == 1 else 5.0
    ch = min((BODY_H - gap * (rows - 1)) / rows, max(ideal, min_h))
    block_h = ch * rows + gap * (rows - 1)
    top = BODY_Y + max((BODY_H - block_h) / 2, 0) * 0.85
    for i, it in enumerate(items):
        r, c = divmod(i, cols)
        x = M + c * (cw + gap)
        y = top + r * (ch + gap)
        rect(s, x, y, cw, ch, theme["tint"], line=None, radius=0.28)
        num = str(it.get("mark") or f"{i+1:02d}")
        note_h = est_height(it.get("note", ""), 18, cw - 1.9)
        head_h = 2.2
        block = head_h + note_h + 0.4
        top_in = y + max((ch - block) / 2, 1.0)
        # 编号 + 标题同一行
        text(s, num, x + 0.95, top_in - 0.05, 1.6, 1.2, size=20, bold=True,
             color=theme["main"], font=theme["font_num"])
        text(s, it.get("title", ""), x + 2.35, top_in - 0.08, cw - 3.2, 1.2, size=22, bold=True,
             color=theme["secondary"], font=theme["font_title"])
        hline(s, x + 0.95, top_in + 1.35, 1.4, theme["main"], 2.0)
        text(s, it.get("note", ""), x + 0.95, top_in + head_h, cw - 1.9, note_h + 0.4, size=18,
             color=theme["text"], font=theme["font_body"], spacing=1.4)
    page_foot(s, theme, ctx["n"], ctx["total"], spec)


def p_bullets(prs, theme, spec, ctx):
    s = blank(prs, theme)
    page_head(s, spec, theme)
    items = (spec.get("items") or [])[:6]
    n = len(items) or 1
    gap = min(1.5, (BODY_H - 1.0) / n * 0.28)
    item_h = (BODY_H - gap * (n - 1)) / n
    item_h = min(item_h, 3.3)
    block = item_h * n + gap * (n - 1)
    top = BODY_Y + max((BODY_H - block) / 2, 0)
    for i, it in enumerate(items):
        txt = it if isinstance(it, str) else it.get("text", "")
        note = "" if isinstance(it, str) else it.get("note", "")
        y = top + i * (item_h + gap)
        rect(s, M, y, CW, item_h, theme["tint"], line=None, radius=0.22)
        circle(s, M + 0.85, y + item_h / 2 - 0.5, 1.0, theme["tint2"], line=None)
        text(s, f"{i+1:02d}", M + 0.85, y + item_h / 2 - 0.36, 1.0, 0.72, size=12, bold=True,
             color=theme["main"], align="center", font=theme["font_num"])
        text(s, txt, M + 2.3, y + (0.75 if note else item_h / 2 - 0.55), CW - 3.2, 1.3,
             size=20, bold=False, color=theme["secondary"], font=theme["font_body"],
             anchor="t" if note else "m")
        if note:
            text(s, note, M + 2.3, y + item_h - 1.15, CW - 3.2, 0.9, size=16,
                 color=theme["muted"], font=theme["font_body"])
    page_foot(s, theme, ctx["n"], ctx["total"], spec)


def p_kpi(prs, theme, spec, ctx):
    s = blank(prs, theme)
    page_head(s, spec, theme)
    items = (spec.get("items") or [])[:3]
    n = len(items) or 1
    gap = 1.0
    cw = (CW - gap * (n - 1)) / n
    ch = 9.6
    top = BODY_Y + (BODY_H - ch) / 2
    for i, it in enumerate(items):
        x = M + i * (cw + gap)
        rect(s, x, top, cw, ch, theme["card"], line=theme["line"], radius=0.2)
        rect(s, x, top, cw, 0.18, theme["main"], radius=0)
        text(s, it.get("value", ""), x, top + 1.9, cw, 3.4, size=64, bold=True,
             color=theme["main"], align="center", font=theme["font_num"])
        text(s, it.get("unit", ""), x, top + 5.5, cw, 0.9, size=18,
             color=theme["muted"], align="center", font=theme["font_body"])
        hline(s, x + cw * 0.3, top + 6.6, cw * 0.4, theme["line"], 1.0)
        text(s, it.get("label", ""), x + 0.7, top + 7.0, cw - 1.4, 2.2, size=18,
             color=theme["text"], align="center", font=theme["font_body"])
    page_foot(s, theme, ctx["n"], ctx["total"], spec)


def p_timeline(prs, theme, spec, ctx):
    s = blank(prs, theme)
    page_head(s, spec, theme)
    items = (spec.get("items") or [])[:5]
    n = len(items) or 1
    seg = (CW - 2.0) / n
    axis = BODY_Y + BODY_H / 2 - 0.6
    hline(s, M + 1.0, axis, CW - 2.0, theme["tint2"], 3.0)
    for i, it in enumerate(items):
        cx = M + 1.0 + seg * (i + 0.5)
        circle(s, cx - 0.5, axis - 0.5, 1.0, theme["main"], line=theme["bg"], lw=3)
        up = i % 2 == 0
        cardh = 4.6
        y = axis - cardh - 1.1 if up else axis + 1.1
        rect(s, cx - seg / 2 + 0.2, y, seg - 0.4, cardh, theme["card"], line=theme["line"], radius=0.16)
        text(s, it.get("label", ""), cx - seg / 2 + 0.55, y + 0.5, seg - 1.1, 0.8, size=13, bold=True,
             color=theme["main"], font=theme["font_num"])
        text(s, it.get("title", ""), cx - seg / 2 + 0.55, y + 1.45, seg - 1.1, 1.1, size=18, bold=True,
             color=theme["secondary"], font=theme["font_title"])
        text(s, it.get("note", ""), cx - seg / 2 + 0.55, y + 2.7, seg - 1.1, 1.6, size=13,
             color=theme["muted"], font=theme["font_body"], spacing=1.3)
    page_foot(s, theme, ctx["n"], ctx["total"], spec)


def p_compare(prs, theme, spec, ctx):
    s = blank(prs, theme)
    page_head(s, spec, theme)
    half = (CW - 2.2) / 2
    ch = 10.6
    top = BODY_Y + (BODY_H - ch) / 2
    for side, x, col, bg in (("left", M, theme["muted"], theme["card"]),
                             ("right", M + half + 2.2, theme["main"], theme["tint"])):
        blk = spec.get(side) or {}
        rect(s, x, top, half, ch, bg, line=theme["line"], radius=0.2)
        rect(s, x, top, half, 0.16, col, radius=0)
        text(s, blk.get("title", ""), x + 1.0, top + 0.9, half - 2.0, 1.2, size=21, bold=True,
             color=col if side == "left" else theme["main"], font=theme["font_title"])
        hline(s, x + 1.0, top + 2.5, half - 2.0, theme["line"], 0.75)
        y = top + 3.1
        for it in (blk.get("items") or [])[:4]:
            circle(s, x + 1.0, y + 0.28, 0.24, col, line=None)
            text(s, str(it), x + 1.65, y, half - 2.6, 1.9, size=18,
                 color=theme["text"], font=theme["font_body"], spacing=1.3)
            y += 1.9
    circle(s, W / 2 - 1.0, top + ch / 2 - 1.0, 2.0, theme["secondary"], line=None)
    text(s, "VS", W / 2 - 1.0, top + ch / 2 - 0.42, 2.0, 0.9, size=14, bold=True,
         color="#FFFFFF", align="center", font=theme["font_num"])
    page_foot(s, theme, ctx["n"], ctx["total"], spec)


def p_image_text(prs, theme, spec, ctx):
    s = blank(prs, theme)
    iw = int(spec.get("image_width", 44)) / 100 * W
    picture(s, spec.get("image"), 0, 0, iw, H, radius=0)
    tx, tw = iw + 2.4, W - iw - 2.4 - M
    text(s, spec.get("title", ""), tx, 4.2, tw, 3.2, size=30, bold=True,
         color=theme["secondary"], font=theme["font_title"], spacing=1.2)
    rect(s, tx, 7.9, 2.2, 0.12, theme["main"], radius=0)
    y = 8.9
    for it in (spec.get("items") or [])[:4]:
        circle(s, tx, y + 0.28, 0.26, theme["main"], line=None)
        text(s, str(it), tx + 0.7, y, tw - 0.8, 1.9, size=16,
             color=theme["text"], font=theme["font_body"], spacing=1.35)
        y += 1.95
    page_foot(s, theme, ctx["n"], ctx["total"], spec)


def p_image_center(prs, theme, spec, ctx):
    """整图页：标题在上，图片在内容区居中按比例缩放。"""
    s = blank(prs, theme)
    page_head(s, spec, theme)
    from PIL import Image
    path = spec.get("image")
    ih_area = BODY_H - 0.4
    if path and os.path.exists(path):
        iw, ih = Image.open(path).size
        r = iw / ih
        if CW / ih_area > r:            # 图更高 → 以高度为准
            hh = ih_area
            ww = hh * r
        else:                            # 图更宽 → 以宽度为准
            ww = CW
            hh = ww / r
        x = M + (CW - ww) / 2
        y = BODY_Y + (ih_area - hh) / 2
        picture(s, path, x, y, ww, hh, radius=0.15)
    else:
        picture(s, path, M, BODY_Y, CW, ih_area, radius=0.15)
    page_foot(s, theme, ctx["n"], ctx["total"], spec)


def p_closing(prs, theme, spec, ctx):
    s = blank(prs, theme, theme["secondary"])
    circle(s, W / 2 - 8.0, -6.0, 16.0, None, line=theme["dark2"], lw=2.0)
    text(s, spec.get("title", "谢谢"), M, 6.6, CW, 3.4, size=58, bold=True,
         color="#FFFFFF", align="center", font=theme["font_title"])
    rect(s, W / 2 - 1.3, 10.4, 2.6, 0.14, theme["main"], radius=0)
    if spec.get("note"):
        text(s, spec["note"], M, 11.4, CW, 1.8, size=17, color="#AEB7C4",
             align="center", font=theme["font_body"])



# ---------- 版式变体：避免所有内容页退化成 cards ----------
def p_asym_split(prs, theme, spec, ctx):
    s = blank(prs, theme)
    image = spec.get("image") or spec.get("hero")
    iw = float(spec.get("image_width", 56)) / 100 * W
    side = spec.get("anchor_position", "right")
    if side == "left":
        picture(s, image, 0, 0, iw, H, radius=0)
        tx, tw = iw + 1.5, W - iw - M - 1.5
    else:
        picture(s, image, W - iw, 0, iw, H, radius=0)
        tx, tw = M, W - iw - M - 1.5
    text(s, spec.get("kicker", ""), tx, 2.0, tw, 0.8, size=15, color=theme["main"], font=theme["font_body"])
    text(s, spec.get("title", ""), tx, 3.1, tw, 3.1, size=34, bold=True, color=theme["secondary"], font=theme["font_title"])
    rect(s, tx, 6.8, 2.8, 0.14, theme["main"], radius=0)
    y = 7.7
    for i, item in enumerate((spec.get("items") or [])[:4]):
        value = item if isinstance(item, str) else item.get("text", item.get("title", ""))
        note = "" if isinstance(item, str) else item.get("note", "")
        circle(s, tx, y + 0.15, 0.42, theme["main"], line=None)
        text(s, value, tx + 0.72, y, tw - 0.72, 0.9, size=18, bold=True, color=theme["secondary"], font=theme["font_title"])
        if note:
            text(s, note, tx + 0.72, y + 0.85, tw - 0.72, 1.0, size=14, color=theme["muted"], font=theme["font_body"])
        y += 2.1 if note else 1.55
    page_foot(s, theme, ctx["n"], ctx["total"], spec)


def p_diagonal(prs, theme, spec, ctx):
    s = blank(prs, theme, theme["secondary"])
    band = s.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Cm(-3), Cm(-2), Cm(W * 0.58), Cm(H + 4))
    band.fill.solid(); band.fill.fore_color.rgb = hexc(theme["main"]); band.line.fill.background()
    text(s, spec.get("kicker", "转折"), 1.6, 2.2, 12, 0.8, size=15, color=theme["accent"], font=theme["font_body"])
    text(s, spec.get("title", ""), 1.6, 3.4, 14.5, 3.6, size=36, bold=True, color="#FFFFFF", font=theme["font_title"])
    text(s, spec.get("sub", ""), 1.6, 7.8, 12.5, 1.8, size=18, color="#D7E7DF", font=theme["font_body"])
    items = (spec.get("items") or [])[:3]
    for i, item in enumerate(items):
        x, yy = 16.2 + i * 5.1, 4.4 + i * 1.15
        rect(s, x, yy, 4.4, 5.6, theme["card"], line=theme["line"], radius=0.2)
        title = item if isinstance(item, str) else item.get("title", "")
        note = "" if isinstance(item, str) else item.get("note", "")
        text(s, f"0{i+1}", x + 0.55, yy + 0.55, 1.2, 0.8, size=20, bold=True, color=theme["main"], font=theme["font_num"])
        text(s, title, x + 0.55, yy + 1.75, 3.3, 1.1, size=19, bold=True, color=theme["secondary"], font=theme["font_title"])
        if note:
            text(s, note, x + 0.55, yy + 3.15, 3.3, 1.65, size=14, color=theme["muted"], font=theme["font_body"])
    page_foot(s, theme, ctx["n"], ctx["total"], spec)


def p_orbit(prs, theme, spec, ctx):
    s = blank(prs, theme)
    page_head(s, spec, theme)
    cx, cy = M + CW * 0.53, BODY_Y + BODY_H * 0.48
    circle(s, cx - 3.1, cy - 3.1, 6.2, theme["tint"], line=theme["main"], lw=2.0)
    text(s, spec.get("center", spec.get("title", "核心")), cx - 2.1, cy - 0.8, 4.2, 1.6, size=22, bold=True, color=theme["secondary"], align="center", anchor="m", font=theme["font_title"])
    items = (spec.get("items") or [])[:6]
    math = __import__("math")
    for i, item in enumerate(items):
        a = -math.pi / 2 + i * (2 * math.pi / max(len(items), 1))
        px, py = cx + math.cos(a) * 9.0, cy + math.sin(a) * 5.0
        circle(s, px - 0.46, py - 0.46, 0.92, theme["accent"], line=None)
        title = item if isinstance(item, str) else item.get("title", "")
        note = "" if isinstance(item, str) else item.get("note", "")
        text(s, title, px - 2.5, py + 0.7, 5.0, 0.8, size=16, bold=True, color=theme["secondary"], align="center", font=theme["font_title"])
        if note:
            text(s, note, px - 2.7, py + 1.55, 5.4, 1.0, size=12, color=theme["muted"], align="center", font=theme["font_body"])
    page_foot(s, theme, ctx["n"], ctx["total"], spec)


def p_staircase(prs, theme, spec, ctx):
    s = blank(prs, theme)
    page_head(s, spec, theme)
    items = (spec.get("items") or [])[:4]
    gap = 0.55
    cw = (CW - gap * (len(items) - 1)) / max(len(items), 1)
    for i, item in enumerate(items):
        x, h = M + i * (cw + gap), 4.6 + i * 1.35
        y = BODY_Y + BODY_H - h
        rect(s, x, y, cw, h, theme["tint"] if i % 2 == 0 else theme["card"], line=theme["line"], radius=0.2)
        title = item if isinstance(item, str) else item.get("title", "")
        note = "" if isinstance(item, str) else item.get("note", "")
        text(s, f"0{i+1}", x + 0.55, y + 0.5, 1.4, 0.8, size=20, bold=True, color=theme["main"], font=theme["font_num"])
        text(s, title, x + 0.55, y + 1.55, cw - 1.1, 1.1, size=18, bold=True, color=theme["secondary"], font=theme["font_title"])
        if note:
            text(s, note, x + 0.55, y + 2.85, cw - 1.1, h - 3.35, size=14, color=theme["muted"], font=theme["font_body"])
    page_foot(s, theme, ctx["n"], ctx["total"], spec)


def p_collage(prs, theme, spec, ctx):
    s = blank(prs, theme)
    page_head(s, spec, theme)
    images = [spec.get("image")] + list(spec.get("images") or [])
    images = [x for x in images if x]
    if images:
        picture(s, images[0], M + 1.0, BODY_Y + 0.25, CW * 0.58, BODY_H * 0.86, radius=0.15)
    if len(images) > 1:
        picture(s, images[1], M + CW * 0.48, BODY_Y + 1.15, CW * 0.37, BODY_H * 0.43, radius=0.15)
    if len(images) > 2:
        picture(s, images[2], M + CW * 0.58, BODY_Y + BODY_H * 0.53, CW * 0.29, BODY_H * 0.34, radius=0.15)
    tx = M + CW * 0.72
    text(s, spec.get("title", ""), tx, BODY_Y + 0.8, CW * 0.25, 2.8, size=28, bold=True, color=theme["secondary"], font=theme["font_title"])
    y = BODY_Y + 4.1
    for item in (spec.get("items") or [])[:3]:
        title = item if isinstance(item, str) else item.get("title", item.get("text", ""))
        note = "" if isinstance(item, str) else item.get("note", "")
        text(s, title, tx, y, CW * 0.25, 0.8, size=16, bold=True, color=theme["main"], font=theme["font_title"])
        if note:
            text(s, note, tx, y + 0.85, CW * 0.25, 1.2, size=13, color=theme["muted"], font=theme["font_body"])
        y += 2.2
    page_foot(s, theme, ctx["n"], ctx["total"], spec)


def p_quote(prs, theme, spec, ctx):
    s = blank(prs, theme, theme["secondary"])
    circle(s, W - 12.5, -5.0, 18.5, None, line=theme["dark2"], lw=1.5)
    text(s, spec.get("kicker", "关键判断"), M, 3.1, CW, 0.8, size=15, color=theme["main"], font=theme["font_body"])
    text(s, spec.get("quote", spec.get("title", "")), M, 5.0, CW * 0.78, 5.6, size=42, bold=True, color="#FFFFFF", font=theme["font_title"])
    if spec.get("note"):
        text(s, spec["note"], M, 12.1, CW * 0.72, 1.6, size=18, color="#B8C9BF", font=theme["font_body"])
    if spec.get("source"):
        text(s, spec["source"], M, 16.8, CW, 0.7, size=10, color="#7B9586", font=theme["font_body"])

BUILDERS = {
    "cover": p_cover, "section": p_section, "cards": p_cards, "bullets": p_bullets,
    "kpi": p_kpi, "timeline": p_timeline, "compare": p_compare,
    "image_text": p_image_text, "image_center": p_image_center, "closing": p_closing,
    "three_cols": lambda p, t, sp, c: p_cards(p, t, dict(sp, cols=3), c),
    "two_cols": lambda p, t, sp, c: p_cards(p, t, dict(sp, cols=2), c),
    "asym_split": p_asym_split, "diagonal": p_diagonal, "orbit": p_orbit,
    "staircase": p_staircase, "collage": p_collage, "quote": p_quote,
}


def build(spec, out):
    theme = theme_of(spec)
    prs = Presentation()
    prs.slide_width = Cm(W)
    prs.slide_height = Cm(H)
    pages = spec.get("pages") or []
    try:
        from motion import apply_slide_motion
    except Exception:
        apply_slide_motion = None
    for i, page in enumerate(pages, 1):
        kind = page.get("layout_family") or page.get("layout") or page.get("type", "cards")
        fn = BUILDERS.get(kind)
        if not fn:
            raise SystemExit(f"未知页面类型: {kind}（可用: {', '.join(BUILDERS)}）")
        fn(prs, theme, page, {"n": i, "total": len(pages)})
        if apply_slide_motion:
            apply_slide_motion(prs.slides[-1], page)
    d = os.path.dirname(os.path.abspath(out))
    if d:
        os.makedirs(d, exist_ok=True)
    prs.save(out)
    return out, len(pages)


def main():
    ap = argparse.ArgumentParser(description="用 YAML/JSON 规格生成 PPT")
    ap.add_argument("spec")
    ap.add_argument("-o", "--out", default="out.pptx")
    ap.add_argument("--check", action="store_true")
    args = ap.parse_args()
    raw = open(args.spec, encoding="utf-8").read()
    spec = json.loads(raw) if args.spec.endswith(".json") else __import__("yaml").safe_load(raw)
    out, n = build(spec, args.out)
    print(f"已生成 {out}（{n} 页，{os.path.getsize(out)/1024:.0f} KB）")
    if args.check:
        import subprocess
        here = os.path.dirname(os.path.abspath(__file__))
        subprocess.run([sys.executable, os.path.join(here, "pptx_inspect.py"), out])


if __name__ == "__main__":
    main()
