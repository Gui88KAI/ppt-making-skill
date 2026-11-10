#!/usr/bin/env python3
"""build_template.py — 按素材库逆向出的「中密度配方」生成一套 PPT。

配方来源：references/模板逆向规范.md（20 套模板、471 页统计），中密度档 20–36 元素/页。

一页内容页的构成（可复用）：
  背景层   全幅渐变底 + 光弧 + 网格 + 四角角标
  标题带   图标块 + 标题 40pt + 标题下装饰线 + 副标题 24pt + 顶部装饰条
  内容带   3 张卡片（底框 + 顶部色条 + 编号 + 标题 + 分隔线 + 正文 + 底部标签）
  页脚带   分隔线 + 页码 + 栏目名

用法:
  python3 build_template.py --out 输出.pptx --theme green
  python3 build_template.py --out x.pptx --pages 5    # 只生成前 5 页
"""
import argparse
import math
import os
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Cm, Pt

W, H = 33.867, 19.05

# ---- 逆向出的三段式分区（cm）----
TITLE_BAND = (0.0, 2.7)      # 标题带
CONTENT_BAND = (2.7, 16.5)   # 内容带
FOOTER_BAND = (16.5, 19.05)  # 页脚带

THEMES = {
    "green": dict(bg1="#0B2318", bg2="#123A28", main="#39C46B", accent="#F2C94C",
                  card="#16402C", line="#2E6B4C", text="#EAF6EE", muted="#9BBCA8"),
    "blue": dict(bg1="#081A33", bg2="#0C2C52", main="#2F80FF", accent="#00D8FF",
                 card="#0F2E52", line="#25517F", text="#E8F1FF", muted="#8FA6C2"),
    "red": dict(bg1="#2A0B0B", bg2="#3D1212", main="#D9A441", accent="#FF6B5B",
                card="#4A1A19", line="#7A3B33", text="#F8EEE4", muted="#C2A08F"),
    "light": dict(bg1="#F7F9FC", bg2="#EAF0F8", main="#1E5AA8", accent="#00A3FF",
                  card="#FFFFFF", line="#D3DEEB", text="#1F2833", muted="#6B7A8C"),
}


def hx(s):
    s = s.lstrip("#")
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def rgb(s):
    s = s.lstrip("#")
    return tuple(int(s[i:i + 2], 16) for i in (0, 2, 4))


def alpha_fill(shape, color, alpha):
    """给形状填充并设置透明度（0-1，1 为不透明）。"""
    if color is None:
        shape.fill.background()
        return
    shape.fill.solid()
    shape.fill.fore_color.rgb = hx(color)
    try:
        from pptx.oxml.ns import qn
        from lxml import etree
        sp = shape.fill._xPr.find(qn("a:solidFill"))
        clr = sp.find(qn("a:srgbClr"))
        a = etree.SubElement(clr, qn("a:alpha"))
        a.set("val", str(int(alpha * 100000)))
    except Exception:
        pass


def rect(slide, x, y, w, h, color=None, line=None, radius=0.2, lw=1.0, alpha=1.0):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
        Cm(x), Cm(y), Cm(w), Cm(h))
    if radius:
        try:
            shp.adjustments[0] = min(radius / max(min(w, h), 0.01), 0.5)
        except Exception:
            pass
    if color:
        alpha_fill(shp, color, alpha)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = hx(line)
        shp.line.width = Pt(lw)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def ellipse(slide, x, y, w, h, color=None, line=None, lw=1.0, alpha=1.0):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Cm(x), Cm(y), Cm(w), Cm(h))
    if color:
        alpha_fill(shp, color, alpha)
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
         font="微软雅黑", anchor="t", spacing=1.25):
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    for i, line in enumerate(str(s).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        p.line_spacing = spacing
        r = p.add_run()
        r.text = line
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
    return c


# ---------------- 背景层 ----------------
def background(slide, th):
    # 全幅底（渐变用两张大圆模拟，避免依赖渐变 API 的兼容问题）
    base = rect(slide, 0, 0, W, H, th["bg1"], radius=0)
    base.fill.gradient()
    try:
        stops = base.fill.gradient_stops
        stops[0].color.rgb = hx(th["bg1"])
        stops[0].position = 0.0
        stops[1].color.rgb = hx(th["bg2"])
        stops[1].position = 1.0
        base.fill.gradient_angle = 45
    except Exception:
        pass
    # 光弧
    ellipse(slide, W * 0.58, -14.0, 28.0, 28.0, th["main"], alpha=0.10)
    ellipse(slide, W * 0.72, -9.0, 18.0, 18.0, th["accent"], alpha=0.07)
    # 细网格
    n = 18
    for i in range(1, n):
        x = W * i / n
        hline(slide, x, 0, 0.01, th["line"], 0.5)
        c = slide.shapes.add_connector(1, Cm(x), Cm(0), Cm(x), Cm(H))
        c.line.color.rgb = hx(th["line"])
        c.line.width = Pt(0.5)
        c.line.fill.background() if False else None
        try:
            from pptx.oxml.ns import qn
            from lxml import etree
            ln = c.line._get_or_add_ln()
            solid = ln.find(qn('a:solidFill'))
            if solid is not None:
                clr = solid.find(qn('a:srgbClr'))
                a = etree.SubElement(clr, qn('a:alpha'))
                a.set('val', '12000')
        except Exception:
            pass
    # 四角角标
    for (cx, cy, s, ang) in [(-1.2, -1.2, 5.2, 0), (W - 4.0, -1.2, 5.2, 1),
                             (-1.2, H - 4.0, 5.2, 2), (W - 4.0, H - 4.0, 5.2, 3)]:
        rect(slide, cx, cy, s, s, None, line=th["main"], radius=0.1, lw=1.2)
        rect(slide, cx + 0.5, cy + 0.5, s - 1.0, s - 1.0, None, line=th["accent"], radius=0.1, lw=0.6)


# ---------------- 标题带 ----------------
def title_band(slide, th, title, sub, tag, page, total):
    # 图标块
    rect(slide, 1.1, 0.95, 1.05, 1.05, th["main"], radius=0.22)
    rect(slide, 1.36, 1.21, 0.54, 0.54, th["bg1"], radius=0.12)
    # 标题
    text(slide, title, 2.55, 0.82, 21.0, 1.5, size=40, bold=True, color=th["text"])
    # 标题下装饰线
    hline(slide, 2.6, 2.32, 3.4, th["main"], 3.0)
    hline(slide, 6.15, 2.32, 1.0, th["accent"], 3.0)
    # 右对齐栏目
    text(slide, tag, W - 8.0, 1.05, 6.6, 0.9, size=18, color=th["muted"], align="r")
    # 副标题（24pt，右上）
    if sub:
        text(slide, sub, 10.0, 2.62, W - 11.5, 0.9, size=18, color=th["muted"])
    # 顶部装饰条
    rect(slide, 0, 0, W, 0.22, th["main"], radius=0, alpha=0.75)
    rect(slide, 0, 0.22, 12.0, 0.06, th["accent"], radius=0)


# ---------------- 内容带：三卡片 ----------------
def card_column(slide, th, x, y, w, h, num, title, body, foot):
    # 底框
    rect(slide, x, y, w, h, th["card"], line=th["line"], radius=0.28, lw=1.0)
    # 顶部色条
    rect(slide, x + 0.28, y + 0.28, w - 0.56, 0.12, th["main"], radius=0.05)
    # 编号圆
    ellipse(slide, x + 0.5, y + 0.72, 1.15, 1.15, th["main"], alpha=1.0)
    text(slide, f"{num:02d}", x + 0.5, y + 1.02, 1.15, 0.7, size=15, bold=True,
         color=th["bg1"], align="c", font="Arial")
    # 卡片标题
    text(slide, title, x + 1.95, y + 0.78, w - 2.5, 1.1, size=24, bold=True, color=th["text"])
    # 分隔线
    hline(slide, x + 0.5, y + 2.25, w - 1.0, th["line"], 1.0)
    # 正文
    text(slide, body, x + 0.5, y + 2.55, w - 1.0, h - 4.4, size=19, color=th["text"], spacing=1.5)
    # 正文下方装饰：三条渐隐横线，填空白
    for k in range(3):
        hline(slide, x + 0.5, y + h - 3.3 + k * 0.42, w - 1.0 - k * 1.2, th["line"], 1.0)
    # 底部标签
    rect(slide, x + 0.5, y + h - 1.45, w - 1.0, 0.9, th["accent"], radius=0.18, alpha=0.16)
    text(slide, foot, x + 0.5, y + h - 1.22, w - 1.0, 0.55, size=16,
         color=th["accent"], align="c")


def content_band(slide, th, cards):
    gap = 1.0
    x0 = 1.1
    total_w = W - 2.2
    cw = (total_w - gap * (len(cards) - 1)) / len(cards)
    y = 4.3
    ch = 10.4
    for i, c in enumerate(cards):
        card_column(slide, th, x0 + i * (cw + gap), y, cw, ch, i + 1,
                    c[0], c[1], c[2])
    # 卡片间连接箭头
    for i in range(len(cards) - 1):
        ax = x0 + (i + 1) * cw + i * gap + 0.15
        ellipse(slide, ax, y + ch / 2 - 0.3, 0.6, 0.6, th["accent"], alpha=0.9)


# ---------------- 页脚带 ----------------
def footer_band(slide, th, page, total, left_text):
    hline(slide, 1.1, 17.35, W - 2.2, th["line"], 0.75)
    text(slide, left_text, 1.1, 17.55, 18.0, 0.8, size=14, color=th["muted"])
    text(slide, f"{page:02d} / {total:02d}", W - 5.0, 17.5, 3.9, 0.9, size=14,
         color=th["muted"], align="r", font="Arial")
    rect(slide, W - 5.6, 17.62, 0.4, 0.5, th["main"], radius=0)


def blank(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    return s


# ---------------- 页面类型 ----------------
def page_cover(prs, th, d, n, total):
    s = blank(prs)
    background(s, th)
    rect(s, 0, 0, W, H, th["bg1"], radius=0, alpha=0.55)
    # 左侧竖色条
    rect(s, 1.1, 3.4, 0.18, 7.6, th["main"], radius=0)
    text(s, d["kicker"], 1.6, 2.85, 26.0, 0.9, size=18, color=th["accent"])
    text(s, d["title"], 1.6, 4.1, 22.5, 4.8, size=52, bold=True, color=th["text"], spacing=1.18)
    hline(s, 1.65, 9.35, 4.2, th["main"], 3.5)
    text(s, d["subtitle"], 1.6, 10.0, 26.0, 1.6, size=24, color=th["muted"])
    # 三个关键词胶囊
    x = 1.6
    for k in d.get("tags", []):
        w = 0.55 * len(k) + 1.2
        rect(s, x, 12.4, w, 1.15, th["main"], radius=0.5, alpha=0.30)
        text(s, k, x, 12.68, w, 0.7, size=17, color=th["accent"], align="c")
        x += w + 0.6
    text(s, d["org"], 1.6, 15.6, 25.0, 0.9, size=15, color=th["muted"])
    # 右侧装饰圆环
    ellipse(s, 25.6, 9.2, 9.2, 9.2, None, line=th["main"], lw=1.5)
    ellipse(s, 27.3, 10.9, 5.8, 5.8, th["main"], alpha=0.16)
    ellipse(s, 28.7, 12.3, 3.0, 3.0, th["accent"], alpha=0.45)


def page_section(prs, th, d, n, total):
    s = blank(prs)
    background(s, th)
    rect(s, 0, 0, W, H, th["bg1"], radius=0, alpha=0.72)
    rect(s, 1.1, 5.2, 0.3, 8.4, th["main"], radius=0)
    text(s, d["number"], 2.1, 4.6, 14.0, 5.2, size=110, bold=True, color=th["main"], font="Arial")
    text(s, d["title"], 2.2, 10.15, 24.0, 2.4, size=44, bold=True, color=th["text"])
    hline(s, 2.25, 12.75, 5.0, th["accent"], 4.0)
    text(s, d.get("note", ""), 2.25, 13.3, 24.0, 1.4, size=22, color=th["muted"])
    ellipse(s, 24.0, 2.2, 12.0, 12.0, None, line=th["main"], lw=1.5)
    ellipse(s, 26.6, 4.8, 6.8, 6.8, None, line=th["accent"], lw=1.0)
    text(s, f"{n:02d} / {total:02d}", W - 5.0, 17.5, 3.9, 0.9, size=14,
         color=th["muted"], align="r", font="Arial")


def page_cards(prs, th, d, n, total):
    s = blank(prs)
    background(s, th)
    title_band(s, th, d["title"], d.get("sub", ""), d.get("tag", ""), n, total)
    content_band(s, th, d["cards"])
    footer_band(s, th, n, total, d.get("source", ""))


def page_kpi(prs, th, d, n, total):
    s = blank(prs)
    background(s, th)
    title_band(s, th, d["title"], d.get("sub", ""), d.get("tag", ""), n, total)
    items = d["items"]
    gap = 1.2
    cw = (W - 2.2 - gap * (len(items) - 1)) / len(items)
    for i, it in enumerate(items):
        x = 1.1 + i * (cw + gap)
        y = 4.6
        rect(s, x, y, cw, 10.4, th["card"], line=th["line"], radius=0.3)
        rect(s, x + 0.3, y + 0.3, cw - 0.6, 0.12, th["main"], radius=0.05)
        ellipse(s, x + cw / 2 - 2.06, y + 1.05, 4.12, 4.12, None, line=th["main"], lw=1.2)
        text(s, it["value"], x, y + 1.55, cw, 3.2, size=60, bold=True, color=th["accent"],
             align="c", font="Arial")
        text(s, it["unit"], x, y + 5.9, cw, 0.9, size=18, color=th["muted"], align="c")
        hline(s, x + cw * 0.32, y + 7.05, cw * 0.36, th["line"], 1.0)
        text(s, it["label"], x + 0.6, y + 7.4, cw - 1.2, 2.2, size=18, color=th["text"], align="c")
    footer_band(s, th, n, total, d.get("source", ""))


def page_compare(prs, th, d, n, total):
    s = blank(prs)
    background(s, th)
    title_band(s, th, d["title"], d.get("sub", ""), d.get("tag", ""), n, total)
    half = (W - 2.2 - 2.4) / 2
    y, h = 4.3, 11.4
    for side, x, col in (("left", 1.1, th["muted"]), ("right", 1.1 + half + 2.4, th["main"])):
        blk = d[side]
        rect(s, x, y, half, h, th["card"], line=th["line"], radius=0.3)
        rect(s, x + 0.3, y + 0.3, half - 0.6, 0.12, col, radius=0.05)
        text(s, blk["title"], x + 0.9, y + 0.8, half - 1.8, 1.2, size=26, bold=True, color=th["text"])
        hline(s, x + 0.9, y + 2.35, half - 1.8, th["line"], 1.0)
        yy = y + 2.9
        for it in blk["items"]:
            ellipse(s, x + 0.9, yy + 0.3, 0.32, 0.32, col)
            text(s, it, x + 1.5, yy, half - 2.4, 1.8, size=19, color=th["text"])
            yy += 1.95
    ellipse(s, W / 2 - 1.35, y + h / 2 - 1.35, 2.7, 2.7, th["main"], alpha=0.95)
    text(s, "VS", W / 2 - 1.35, y + h / 2 - 0.5, 2.7, 1.0, size=20, bold=True,
         color=th["bg1"], align="c", font="Arial")
    footer_band(s, th, n, total, d.get("source", ""))


def page_end(prs, th, d, n, total):
    s = blank(prs)
    background(s, th)
    rect(s, 0, 0, W, H, th["bg1"], radius=0, alpha=0.72)
    text(s, d.get("title", "谢谢"), 0, 6.6, W, 3.6, size=60, bold=True, color=th["text"], align="c")
    rect(s, W / 2 - 1.6, 10.5, 3.2, 0.16, th["main"], radius=0)
    text(s, d.get("note", ""), 2.0, 11.4, W - 4.0, 1.6, size=22, color=th["muted"], align="c")


BUILDERS = {"cover": page_cover, "section": page_section, "cards": page_cards,
            "kpi": page_kpi, "compare": page_compare, "end": page_end}


def build(deck, out, theme="green", max_pages=None):
    th = THEMES[theme]
    prs = Presentation()
    prs.slide_width = Cm(W)
    prs.slide_height = Cm(H)
    pages = deck["pages"][:max_pages] if max_pages else deck["pages"]
    for i, p in enumerate(pages, 1):
        fn = BUILDERS[p["type"]]
        fn(prs, th, p, i, len(pages))
    d = os.path.dirname(os.path.abspath(out))
    if d:
        os.makedirs(d, exist_ok=True)
    prs.save(out)
    return out, len(pages)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--theme", default="green", choices=list(THEMES))
    ap.add_argument("--pages", type=int, default=0)
    a = ap.parse_args()
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from deck_content import DECK
    out, n = build(DECK, a.out, a.theme, a.pages or None)
    print(f"已生成 {out}（{n} 页，{os.path.getsize(out)/1024:.0f} KB）")
