#!/usr/bin/env python3
"""organic_page.py — 破格排版样张：圆形 / 斜切 / 错位 / 叠层，打破方块感。"""
import os
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

W, H = 33.867, 19.05
BG, CARD, GOLD, TEXT, MUTED, LINE = "#08170F", "#123526", "#D4A24C", "#F2EFE6", "#9DB3A5", "#2C5A43"
FT = "/usr/share/fonts/truetype/custom/"
F_L, F_B, F_K = FT + "阿里巴巴普惠体 L.ttf", FT + "阿里巴巴普惠体 B.ttf", FT + "华文楷体.TTF"
F_N = "/system/fonts/NotoSansCJK-Regular.ttc"


def hx(s):
    s = s.lstrip("#")
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def _alpha(shp, a):
    solid = shp._element.find(".//" + qn("a:solidFill"))
    if solid is None:
        return
    clr = solid.find(qn("a:srgbClr"))
    if clr is not None:
        e = etree.SubElement(clr, qn("a:alpha"))
        e.set("val", str(int(a * 100000)))


def shape(s, kind, x, y, w, h, color=None, alpha=1.0, line=None, lw=1.0, rot=0):
    shp = s.shapes.add_shape(kind, Cm(x), Cm(y), Cm(w), Cm(h))
    if color:
        shp.fill.solid()
        shp.fill.fore_color.rgb = hx(color)
        if alpha < 1:
            _alpha(shp, alpha)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = hx(line)
        shp.line.width = Pt(lw)
    else:
        shp.line.fill.background()
    if rot:
        shp.rotation = rot
    shp.shadow.inherit = False
    return shp


def rect(s, x, y, w, h, color=None, alpha=1.0, radius=0, line=None, lw=1.0, rot=0):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shp = shape(s, kind, x, y, w, h, color, alpha, line, lw, rot)
    if radius:
        try:
            shp.adjustments[0] = min(radius / max(min(w, h), .01), .5)
        except Exception:
            pass
    return shp


def circle(s, x, y, d, color=None, alpha=1.0, line=None, lw=1.5):
    return shape(s, MSO_SHAPE.OVAL, x, y, d, d, color, alpha, line, lw)


def txt(s, t, x, y, w, h, size=16, color=TEXT, font=None, align="l", anchor="t",
        bold=False, sp=1.3, rot=0):
    tb = s.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE}[anchor]
    for i, ln in enumerate(str(t).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        p.line_spacing = sp
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.name = font or F_L
        r.font.color.rgb = hx(color)
    if rot:
        tb.rotation = rot
    return tb


def pic_circle(s, p, x, y, w, h):
    """把图片裁成圆形（直接改 pptx 的几何为 ellipse）。"""
    ph = s.shapes.add_picture(p, Cm(x), Cm(y), Cm(w), Cm(h))
    spPr = ph._element.spPr
    for old in spPr.findall(qn("a:prstGeom")):
        spPr.remove(old)
    g = etree.SubElement(spPr, qn("a:prstGeom"))
    g.set("prst", "ellipse")
    etree.SubElement(g, qn("a:avLst"))
    return ph


def pic(s, p, x, y, w, h, rot=0):
    ph = s.shapes.add_picture(p, Cm(x), Cm(y), Cm(w), Cm(h))
    if rot:
        ph.rotation = rot
    return ph


def build(out="organic.pptx", a="assets"):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Cm(W), Cm(H)
    s = prs.slides.add_slide(prs.slide_layouts[6])

    # —— 背景 ——
    rect(s, 0, 0, W, H, BG)
    if os.path.exists(os.path.join(a, "field.jpg")):
        pic(s, os.path.join(a, "field.jpg"), 0, 0, W, H)
        rect(s, 0, 0, W, H, "#06150E", alpha=0.78)

    # —— 背景装饰：大圆 + 斜切带 + 弧线 ——
    circle(s, -6, 9, 22, None, line=GOLD, lw=1.0)          # 左下大圆（露一半）
    circle(s, -4.2, 10.8, 18.4, None, line="#2C5A43", lw=0.8)
    rect(s, -3, 6.5, 40, 3.4, GOLD, alpha=0.05, rot=-8)     # 斜切光带
    rect(s, -3, 7.2, 40, 0.06, GOLD, alpha=0.35, rot=-8)
    # 右上角弧形装饰（用多个小圆拼）
    for i in range(16):
        circle(s, W - 9 + i * 0.55, 0.4 + i * 0.32, 0.22, GOLD, alpha=0.25)

    # —— 导航条 ——
    rect(s, 0, 0, W, 0.95, "#04120C", alpha=0.95)
    rect(s, 0, 0.95, W, 0.03, GOLD, alpha=0.9)
    txt(s, "慧飞植保 · 无人机智能植保", 1.1, 0.24, 12, 0.5, size=12, color=GOLD, bold=True)
    chapters = ["项目背景", "技术介绍", "痛点分析", "应对策略", "团队介绍", "发展规划"]
    x = W - 1.2
    for i in range(5, -1, -1):
        x -= 3.05
        if i == 2:
            rect(s, x + 0.08, 0.52, 2.9, 0.36, GOLD, alpha=0.2, radius=0.06)
        txt(s, chapters[i], x + 0.08, 0.53, 2.9, 0.34, size=11,
            color=GOLD if i == 2 else "#8EA79B", align="c")

    # —— 眉标 + 主标题（微旋转，打破水平） ——
    txt(s, "01 · 痛点分析", 1.5, 1.22, 10, 0.6, size=14, color=GOLD, rot=-2)
    txt(s, "一亩地打药", 1.45, 1.88, 12, 1.6, size=36, bold=True, color=TEXT, rot=-2)
    txt(s, "40 分钟", 1.35, 3.25, 12, 2.6, size=64, bold=True, color=GOLD, font=F_N, rot=-2)
    txt(s, "人越来越难找了", 1.5, 5.75, 12, 1.1, size=28, bold=True, color=TEXT, rot=-2)
    txt(s, "三个数字，就是我们要解决的问题", 1.6, 6.85, 13, 0.8, size=14, color=MUTED, font=F_K, rot=-2)

    # —— 三个圆形数据（错落，不排队） ——
    data = [("62%", "农户超 50 岁", 1.35, 8.35, 5.2, 38),
            ("40%", "药液流失率", 6.9, 10.05, 4.6, 34),
            ("8-12%", "减产幅度", 11.7, 8.6, 5.4, 30)]
    for num, label, cx, cy, d, nsz in data:
        circle(s, cx, cy, d, "#0E2A1C", alpha=0.94, line=GOLD, lw=1.4)
        circle(s, cx + 0.28, cy + 0.28, d - 0.56, None, line="#2C5A43", lw=0.7)
        txt(s, num, cx, cy + d * 0.30, d, 1.8, size=nsz,
            bold=True, color=GOLD, font=F_N, align="c")
        txt(s, label, cx, cy + d * 0.62, d, 0.9, size=12, color=MUTED, align="c")

    # —— 右侧大圆形照片（主体） ——
    if os.path.exists(os.path.join(a, "p_drone_c.jpg")):
        circle(s, 21.0, 3.6, 11.6, CARD, alpha=0.9)
        pic_circle(s, os.path.join(a, "p_drone_c.jpg"), 21.35, 3.95, 10.9, 10.9)
        circle(s, 21.0, 3.6, 11.6, None, line=GOLD, lw=2.0)
        circle(s, 20.2, 2.8, 12.8, None, line=GOLD, alpha=0.35, lw=0.8)

    # —— 叠层小圆照片（半压在主图上，制造前后关系） ——
    if os.path.exists(os.path.join(a, "p_coop.jpg")):
        circle(s, 18.4, 11.4, 5.0, CARD, alpha=0.96)
        pic_circle(s, os.path.join(a, "p_coop.jpg"), 18.6, 11.6, 4.6, 4.6)
        circle(s, 18.4, 11.4, 5.0, None, line=LINE, lw=1.0)
    if os.path.exists(os.path.join(a, "p_farmer.jpg")):
        circle(s, 25.6, 12.6, 3.6, CARD, alpha=0.96)
        pic_circle(s, os.path.join(a, "p_farmer.jpg"), 25.75, 12.75, 3.3, 3.3)
        circle(s, 25.6, 12.6, 3.6, None, line=GOLD, lw=0.9)

    # —— 引语（斜置小卡，不是方块） ——
    q = rect(s, 12.0, 15.0, 12.6, 2.3, "#0B2218", alpha=0.94, radius=0.5, line=GOLD, lw=0.8, rot=-3)
    txt(s, "「往年请人打药要排队，错过窗口就减产」", 12.4, 15.5, 11.8, 0.8, size=14, color=TEXT, rot=-3)
    txt(s, "—— 营口 · 李大姐家庭农场", 12.4, 16.5, 11.8, 0.6, size=11, color=MUTED, align="r", rot=-3)

    # —— 细节装饰 ——
    for i in range(5):
        circle(s, 1.2 + i * 0.34, 16.5, 0.14, GOLD, alpha=0.4 - i * 0.05)
    rect(s, W - 9.5, 17.2, 8.2, 0.02, LINE)

    # —— 页脚 ——
    rect(s, 0, 17.6, W, 0.02, LINE)
    txt(s, "慧飞植保 · 智慧农业无人机植保", 1.35, 17.75, 14, 0.6, size=11, color=MUTED)
    txt(s, "痛点分析 · 08", W - 6.2, 17.75, 4.9, 0.6, size=11, color=MUTED, align="r")

    prs.save(out)
    print("已生成", out)


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="organic.pptx")
    ap.add_argument("--assets", default="assets")
    a = ap.parse_args()
    build(a.out, a.assets)
