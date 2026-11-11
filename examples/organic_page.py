#!/usr/bin/env python3
"""organic_page.py — v4：字体分层 + 文字真正居中 + 照片环绕轨道。"""
import math
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
# —— 字体分层：每个角色不同字体 ——
F_H1 = FT + "思源宋体 CN Heavy.otf"      # 大标题：宋体（端正、有笔锋）
F_NUM = FT + "汉仪雅酷黑 75W.ttf"        # 大数字：粗黑（力量感）
F_SUB = FT + "华文楷体.TTF"              # 副标/引语：楷体（人文感）
F_BODY = FT + "阿里巴巴普惠体 L.ttf"     # 正文：普惠体
F_EN = FT + "Montserrat--淘宝店：烦烦优创馆.ttf"  # 英文点缀


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


def _dash(shp, kind="dash"):
    ln = shp._element.spPr.find(qn("a:ln"))
    if ln is None:
        return
    for old in ln.findall(qn("a:prstDash")):
        ln.remove(old)
    d = etree.SubElement(ln, qn("a:prstDash"))
    d.set("val", kind)


def shape(s, kind, x, y, w, h, color=None, alpha=1.0, line=None, lw=1.0, rot=0, dash=None):
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
        if dash:
            _dash(shp, dash)
    else:
        shp.line.fill.background()
    if rot:
        shp.rotation = rot
    shp.shadow.inherit = False
    return shp


def rect(s, x, y, w, h, color=None, alpha=1.0, radius=0, line=None, lw=1.0, rot=0, dash=None):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shp = shape(s, kind, x, y, w, h, color, alpha, line, lw, rot, dash)
    if radius:
        try:
            shp.adjustments[0] = min(radius / max(min(w, h), .01), .5)
        except Exception:
            pass
    return shp


def circle(s, x, y, d, color=None, alpha=1.0, line=None, lw=1.0, rot=0, dash=None):
    return shape(s, MSO_SHAPE.OVAL, x, y, d, d, color, alpha, line, lw, rot, dash)


def txt(s, t, x, y, w, h, size=16, color=TEXT, font=None, align="l", anchor="t",
        bold=False, sp=1.25, rot=0):
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
        r.font.name = font or F_BODY
        r.font.color.rgb = hx(color)
    if rot:
        tb.rotation = rot
    return tb


def label_in(shp, lines, font, color, size, bold=True, sp=1.1):
    """把文字写进形状：margin 归零 + 垂直水平双居中。"""
    tf = shp.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.auto_size = None
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        p.line_spacing = sp
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.name = font
        r.font.color.rgb = hx(color)


def pic_circle(s, p, x, y, w, h):
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


def A(a, n):
    return os.path.join(a, n)


def build(out="organic.pptx", a="assets"):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Cm(W), Cm(H)
    s = prs.slides.add_slide(prs.slide_layouts[6])

    # ---------- 背景 ----------
    rect(s, 0, 0, W, H, BG)
    if os.path.exists(A(a, "field.jpg")):
        pic(s, A(a, "field.jpg"), 0, 0, W, H)
        rect(s, 0, 0, W, H, "#06150E", alpha=0.82)

    # ---------- 装饰：斜光带 + 点阵 + 光斑 ----------
    rect(s, -4, 5.6, 42, 3.0, GOLD, alpha=0.04, rot=-9)
    rect(s, -4, 6.3, 42, 0.05, GOLD, alpha=0.28, rot=-9)
    for i in range(14):
        circle(s, W - 8.6 + i * 0.5, 0.7 + i * 0.28, 0.14, GOLD, alpha=0.2)
    for i in range(5):
        circle(s, 1.2 + i * 0.9, 16.9, 0.16, GOLD, alpha=0.35 - i * 0.05)

    # ---------- 导航 ----------
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

    # ---------- 标题区（端正，不旋转） ----------
    rect(s, 1.4, 1.25, 0.07, 0.62, GOLD)
    txt(s, "01 · 痛点分析", 1.68, 1.24, 10, 0.6, size=13, color=GOLD, font=F_BODY)
    txt(s, "一亩地打药", 1.4, 1.95, 13, 1.5, size=30, bold=True, color=TEXT, font=F_H1)
    txt(s, "40", 1.36, 3.1, 4.4, 3.1, size=88, bold=True, color=GOLD, font=F_NUM)
    txt(s, "分钟", 5.15, 4.55, 3.0, 1.0, size=24, bold=True, color=GOLD, font=F_H1)
    txt(s, "人越来越难找了", 1.46, 6.45, 13, 1.0, size=23, bold=True, color=TEXT, font=F_SUB)
    txt(s, "PAIN POINT · 真实田间数据", 1.5, 7.45, 13, 0.6, size=10, color=GOLD, font=F_EN)
    txt(s, "三个数字，就是我们真正要解决的问题", 1.5, 7.85, 13.5, 0.7, size=12.5,
        color=MUTED, font=F_SUB)

    # ---------- 三个数据圆（文字真正居中） ----------
    data = [("62%", "农户超 50 岁", 1.35, 9.1, 5.0, 34),
            ("40%", "药液流失率", 6.75, 10.15, 4.5, 30),
            ("8-12%", "减产幅度", 11.5, 9.0, 4.7, 22)]
    for num, label, cx, cy, d, nsz in data:
        c1 = circle(s, cx, cy, d, "#0E2A1C", alpha=0.96, line=GOLD, lw=1.5)
        label_in(c1, [num], F_NUM, GOLD, nsz)
        txt(s, label, cx - 0.5, cy + d + 0.18, d + 1.0, 0.6, size=11.5, color=MUTED,
            font=F_SUB, align="c")
        circle(s, cx + d * 0.5 - 0.16, cy + d + 0.05, 0.32, GOLD, alpha=0.55)

    # ---------- 照片环绕轨道 ----------
    if os.path.exists(A(a, "p_drone_c.jpg")):
        ox, oy, od = 21.4, 5.0, 7.8            # 主图圆心/直径
        cx0, cy0 = ox + od / 2, oy + od / 2     # 圆心
        orbit = 5.3                             # 轨道半径（cm）
        # 轨道虚线环
        circle(s, cx0 - orbit, cy0 - orbit, orbit * 2, None, line=GOLD, lw=0.9, dash="sysDash")
        circle(s, cx0 - od / 2 - 0.5, cy0 - od / 2 - 0.5, od + 1.0, None, line="#3B6B52", lw=0.7)
        # 主图
        circle(s, ox, oy, od, CARD, alpha=0.9)
        pic_circle(s, A(a, "p_drone_c.jpg"), ox + 0.2, oy + 0.2, od - 0.4, od - 0.4)
        circle(s, ox, oy, od, None, line=GOLD, lw=2.0)
        # 轨道上的卫星点
        for i in range(10):
            ang = math.radians(i * 36 + 12)
            dot = 0.22 if i % 3 else 0.34
            circle(s, cx0 + orbit * math.cos(ang) - dot / 2,
                   cy0 + orbit * math.sin(ang) - dot / 2, dot, GOLD, alpha=0.5)
        # 沿轨道放置的两张小图（真实环绕）
        sats = [(A(a, "p_coop.jpg"), 200, 3.2), (A(a, "p_farmer.jpg"), 335, 2.9)]
        for path, deg, dia in sats:
            if not os.path.exists(path):
                continue
            ang = math.radians(deg)
            px = cx0 + orbit * math.cos(ang) - dia / 2
            py = cy0 + orbit * math.sin(ang) - dia / 2
            circle(s, px - 0.18, py - 0.18, dia + 0.36, "#0B2218", alpha=0.95, line=GOLD, lw=1.0)
            pic_circle(s, path, px, py, dia, dia)
            circle(s, px - 0.18, py - 0.18, dia + 0.36, None, line=GOLD, lw=1.0)

    # ---------- 引语（端正的卡，只微斜） ----------
    rect(s, 11.4, 14.5, 12.6, 2.1, "#0B2218", alpha=0.95, radius=0.4, line=GOLD, lw=0.8)
    txt(s, "「往年请人打药要排队，错过窗口就减产」", 11.85, 14.9, 11.8, 0.8, size=13,
        color=TEXT, font=F_SUB)
    txt(s, "—— 营口 · 李大姐家庭农场", 11.85, 15.75, 11.8, 0.6, size=10.5,
        color=MUTED, align="r")

    # ---------- 页脚 ----------
    rect(s, 0, 17.6, W, 0.02, LINE)
    rect(s, W - 9.2, 16.95, 8.2, 0.02, LINE, alpha=0.8)
    rect(s, W - 9.2, 17.1, 4.6, 0.02, GOLD, alpha=0.6)
    txt(s, "慧飞植保 · 智慧农业无人机植保", 1.35, 17.75, 14, 0.6, size=10, color=MUTED)
    txt(s, "痛点分析 · 08", W - 6.2, 17.75, 4.9, 0.6, size=10, color=MUTED, align="r")

    prs.save(out)
    print("已生成", out)


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="organic.pptx")
    ap.add_argument("--assets", default="assets")
    a = ap.parse_args()
    build(a.out, a.assets)
