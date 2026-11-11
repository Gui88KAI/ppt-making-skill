#!/usr/bin/env python3
"""deck_v5.py — 设计系统样张（5 页）：原生可编辑文字 + 真实字体。

字体分层：
  标题  思源宋体 CN Heavy（端正衬线）
  点题  演示流云楷（中文笔锋，只用在 1 行）
  副标  华文楷体
  正文  阿里巴巴普惠体 L/R
  数字  Bahnschrift / Montserrat（西文）
版式：标题一律水平；动感只交给背景弧线、斜光带与环绕轨道。
缺数据/素材一律写「待补 / 待测」，不编造。
"""
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
BG = "#0A241D"
CARD = "#10392B"
GOLD = "#D9B875"
WHITE = "#F4EFDF"
MUTED = "#A5BCAE"
LINE = "#3B6B52"
FT = "/usr/share/fonts/truetype/custom/"
F_TITLE = "思源宋体 CN Heavy"
F_BRUSH = "演示流云楷"
F_KAI = "华文楷体"
F_BODY = "阿里巴巴普惠体 L"
F_BODY_B = "阿里巴巴普惠体 B"
F_EN = "Montserrat"
F_NUM = "Bahnschrift"
ASSETS = os.environ.get("ASSETS_DIR", "assets")

NAV = ["项目背景", "痛点分析", "解决方案", "商业落地", "团队规划"]


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


def _dash(shp, kind="sysDash"):
    ln = shp._element.spPr.find(qn("a:ln"))
    if ln is None:
        return
    for old in ln.findall(qn("a:prstDash")):
        ln.remove(old)
    d = etree.SubElement(ln, qn("a:prstDash"))
    d.set("val", kind)


def shape(s, kind, x, y, w, h, color=None, alpha=1.0, line=None, lw=1.0, dash=None):
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
    shp.shadow.inherit = False
    return shp


def rect(s, x, y, w, h, color=None, alpha=1.0, line=None, lw=1.0):
    return shape(s, MSO_SHAPE.RECTANGLE, x, y, w, h, color, alpha, line, lw)


def circle(s, x, y, d, color=None, alpha=1.0, line=None, lw=1.0, dash=None):
    return shape(s, MSO_SHAPE.OVAL, x, y, d, d, color, alpha, line, lw, dash)


def arc(s, x, y, w, h, color=GOLD, lw=1.0):
    """弧形：用 MSO 的 ARC 形状近似。"""
    a = s.shapes.add_shape(MSO_SHAPE.ARC, Cm(x), Cm(y), Cm(w), Cm(h))
    a.fill.background()
    a.line.color.rgb = hx(color)
    a.line.width = Pt(lw)
    a.shadow.inherit = False
    return a


def text(s, t, x, y, w, h, size=16, color=WHITE, font=F_BODY, align="l",
         anchor="t", bold=False, sp=1.25, fit=True):
    """原生文字（可在 PPT 里直接编辑）。fit=True 时按宽度缩字号。"""
    tb = s.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE,
                          "b": MSO_ANCHOR.BOTTOM}[anchor]
    for i, ln in enumerate(str(t).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        p.line_spacing = sp
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.name = font
        r.font.color.rgb = hx(color)
    return tb


def text_in(shp, lines, font, color, size, bold=True, sp=1.15):
    """文字写进形状内部：四边 margin=0 + 双居中，保证不出框。"""
    tf = shp.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
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


def pic_circle(s, p, x, y, d):
    ph = s.shapes.add_picture(p, Cm(x), Cm(y), Cm(d), Cm(d))
    spPr = ph._element.spPr
    for old in spPr.findall(qn("a:prstGeom")):
        spPr.remove(old)
    g = etree.SubElement(spPr, qn("a:prstGeom"))
    g.set("prst", "ellipse")
    etree.SubElement(g, qn("a:avLst"))
    return ph



# ---------------- 小图标（代码绘制，统一线宽与色） ----------------
def icon(s, kind, x, y, d, color=GOLD, lw=1.4):
    """在 (x,y) 画一个 d×d 的线性图标。"""
    u = d / 10.0
    def L(x1, y1, x2, y2):
        c = s.shapes.add_connector(1, Cm(x + x1 * u), Cm(y + y1 * u), Cm(x + x2 * u), Cm(y + y2 * u))
        c.line.color.rgb = hx(color); c.line.width = Pt(lw)
        return c
    if kind == "droplet":
        t = shape(s, MSO_SHAPE.ISOSCELES_TRIANGLE, x + 2.6 * u, y + 1.2 * u, 4.8 * u, 4.6 * u,
                  color=color, alpha=1.0)
        c = circle(s, x + 2.4 * u, y + 4.4 * u, 5.2 * u, color=color, alpha=1.0)
        return [t, c]
    if kind == "leaf":
        a = s.shapes.add_shape(MSO_SHAPE.MOON, Cm(x + 1.4 * u), Cm(y + 1.4 * u), Cm(7.2 * u), Cm(7.2 * u))
        a.fill.background(); a.line.color.rgb = hx(color); a.line.width = Pt(lw)
        a.rotation = 45
        L(2.4, 7.6, 7.6, 2.4)
        return [a]
    if kind == "clock":
        c = circle(s, x + 1.2 * u, y + 1.2 * u, 7.6 * u, None, line=color, lw=lw)
        L(5, 5, 5, 2.2); L(5, 5, 7.3, 6.0)
        return [c]
    if kind == "team":
        circle(s, x + 2.0 * u, y + 1.4 * u, 2.6 * u, None, line=color, lw=lw)
        a = s.shapes.add_shape(MSO_SHAPE.ARC, Cm(x + 0.8 * u), Cm(y + 4.0 * u), Cm(5.0 * u), Cm(5.0 * u))
        a.fill.background(); a.line.color.rgb = hx(color); a.line.width = Pt(lw)
        circle(s, x + 5.6 * u, y + 2.4 * u, 2.0 * u, None, line=color, lw=lw * 0.9)
        return [a]
    if kind == "target":
        for k, r in enumerate((4.2, 2.6, 1.1)):
            circle(s, x + 5.0 * u - r * u, y + 5.0 * u - r * u, 2 * r * u,
                   None if k < 2 else color, line=color if k < 2 else None, lw=lw)
        return []
    if kind == "gear":
        c = circle(s, x + 2.6 * u, y + 2.6 * u, 4.8 * u, None, line=color, lw=lw)
        circle(s, x + 4.3 * u, y + 4.3 * u, 1.4 * u, None, line=color, lw=lw * 0.9)
        for i in range(6):
            import math as _m
            a0 = _m.radians(i * 60)
            L(5 + 3.1 * _m.cos(a0), 5 + 3.1 * _m.sin(a0),
              5 + 4.2 * _m.cos(a0), 5 + 4.2 * _m.sin(a0))
        return [c]


# ---------------- 共用底 ----------------
def page(prs, variant="plain"):
    """四种背景变体，避免每页一样：plain / dark / split / grid。"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, W, H, BG)

    if variant == "dark":
        # 目录：近黑底 + 右上大号数字水印 + 纵向细分割
        rect(s, 0, 0, W, H, "#06170F")
        text(s, "02", 25.0, 10.6, 7.2, 5.0, 210, "#10352A", F_NUM, bold=True, align="r")

    elif variant == "split":
        # 解决方案：上浅下深横向分区 + 右侧大箭头暗纹
        rect(s, 0, 0, W, 9.6, "#0D2C22")
        rect(s, 0, 9.6, W, H - 9.6, "#081B14")
        rect(s, 0, 9.58, W, 0.04, "#4E7F65", alpha=0.6)
        for i in range(3):
            text(s, "→", 26.0 + i * 1.4, 2.2 + i * 3.6, 6, 4, 130, "#11362A",
                 F_BODY, bold=True, align="c")

    elif variant == "grid":
        # 成效：点阵网格底 + 底部深色压条
        for gx in range(0, 34, 2):
            for gy in range(2, 18, 2):
                circle(s, gx + 0.6, gy + 0.4, 0.09, "#2F5A46", alpha=0.55)
        rect(s, 0, 14.6, W, H - 14.6, "#06170F", alpha=0.75)
        rect(s, 0, 14.6, W, 0.03, GOLD, alpha=0.30)

    else:  # plain（痛点页：绿光 + 等高线弧）
        circle(s, 22.0, -3.0, 26, "#123A2C", alpha=0.55)
        circle(s, 24.5, 1.0, 18, "#164534", alpha=0.45)
        for i in range(3):
            c = circle(s, 20.0 - i * 1.1, 1.0 - i * 0.8, 20 + i * 2.2, None,
                       line="#4E7359", lw=0.6)
            _alpha(c, 0.16)

    rect(s, 1.6, 1.62, W - 3.2, 0.02, "#9DB3A5", alpha=0.35)
    rect(s, 1.6, 17.35, W - 3.2, 0.02, "#9DB3A5", alpha=0.35)
    return s


def nav(s, idx=-1):
    text(s, "慧飞植保", 1.6, 0.62, 4.2, 0.8, 13, GOLD, F_KAI)
    text(s, "FIELD INTELLIGENCE", 4.6, 0.7, 6.0, 0.6, 8, MUTED, F_EN)
    x = 15.4
    for i, n in enumerate(NAV):
        text(s, n, x, 0.68, 3.2, 0.7, 11, GOLD if i == idx else MUTED,
             F_BODY_B if i == idx else F_BODY, "l")
        x += 3.35
    if idx >= 0:
        rect(s, 15.4 + idx * 3.35, 1.3, 1.5, 0.05, GOLD)


def footer(s, n, label="DESIGN SAMPLE · 数据与案例尚未核实"):
    text(s, label, 1.6, 17.55, 20, 0.6, 8.5, MUTED, F_BODY)
    text(s, "HUIFEI", 28.2, 17.55, 2.2, 0.6, 8.5, GOLD, F_EN)
    text(s, f"{n:02d}", 30.6, 17.5, 1.8, 0.7, 11, GOLD, F_NUM)


# ---------------- P1 封面 ----------------
def p_cover(prs):
    """封面沿用「封面-融合版」：书法字 + 人物 + 稻田 + 英文点题行 + 信息条。"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    img = os.path.join(ASSETS, "cover_final.jpg")
    if os.path.exists(img):
        s.shapes.add_picture(img, 0, 0, Cm(W), Cm(H))
    else:
        rect(s, 0, 0, W, H, BG)
        text(s, "封面图缺失", 2, 8, 20, 1, 20, MUTED, F_BODY)


# ---------------- P2 目录 ----------------
def p_toc(prs):
    s = page(prs, "dark")
    nav(s, -1)
    text(s, "目录", 1.6, 2.3, 6, 1.4, 40, GOLD, F_TITLE, bold=True)
    text(s, "CONTENTS", 1.75, 3.85, 7, 0.8, 12, MUTED, F_EN)
    rect(s, 1.75, 4.7, 4.0, 0.06, GOLD)
    items = [("01", "项目背景", "政策切口 · 社会痛点 · 技术痛点"),
             ("02", "痛点分析", "用工 · 用药 · 农时"),
             ("03", "解决方案", "三个创新点对应三个痛点"),
             ("04", "商业落地", "服务模式 · 案例 · 推广"),
             ("05", "团队规划", "团队 · 凭证 · 发展路径")]
    y = 6.2
    for num, name, note in items:
        rect(s, 1.75, y, 28.6, 1.85, CARD, alpha=0.55, line="#3E6E56", lw=0.6)
        text(s, num, 2.2, y + 0.42, 2.2, 1.0, 20, GOLD, F_NUM)
        text(s, name, 4.6, y + 0.35, 8.0, 1.0, 20, WHITE, F_TITLE, bold=True)
        text(s, note, 13.2, y + 0.55, 15.0, 0.8, 13, MUTED, F_KAI)
        y += 2.05
    footer(s, 2)


# ---------------- P3 痛点（数据待测） ----------------
def p_pain(prs):
    s = page(prs)
    nav(s, 1)
    text(s, "02", 1.62, 2.3, 2.0, 0.9, 18, GOLD, F_NUM)
    text(s, "田间痛点  /  FIELD CHALLENGES", 2.7, 2.35, 14, 0.7, 12, MUTED, F_BODY)
    text(s, "植保不该", 1.6, 3.3, 14, 2.0, 40, WHITE, F_TITLE, bold=True)
    text(s, "困在人力里", 1.58, 5.2, 14, 2.2, 44, GOLD, F_BRUSH)
    text(s, "从「找人打药」，走向按需、精准作业", 1.65, 7.7, 15, 0.9, 16, WHITE, F_KAI)
    text(s, "把用工、用药与农时放在同一张作业图里。", 1.65, 8.6, 15, 0.8, 12, MUTED, F_BODY)
    # 三个指标圆（数字在圆心）
    data = [("待测", "用工", "作业耗时", 2.0, 10.9, 4.3),
            ("待测", "用药", "有效利用率", 6.9, 11.5, 3.9),
            ("待测", "农时", "响应时效", 11.4, 10.7, 4.2)]
    for num, tag, lab, cx, cy, d, ic in [(d[0], d[1], d[2], d[3], d[4], d[5], k)
                                         for d, k in zip(data, ["team", "droplet", "clock"])]:
        c = circle(s, cx, cy, d, CARD, alpha=0.95, line="#7FA98B", lw=1.2)
        text_in(c, [num], F_TITLE, GOLD, 20)
        icon(s, ic, cx + d / 2 - 0.8, cy + d / 2 - 3.5, 1.6, "#8FB59B", 1.1)
        text(s, tag, cx, cy - 1.05, d, 0.6, 10.5, MUTED, F_KAI, align="c")
        text(s, lab, cx - 0.6, cy + d + 0.3, d + 1.2, 0.55, 11, WHITE, F_KAI, align="c")
    text(s, "单位与统计口径待补 · 实测记录、农户访谈、原始统计表", 1.65, 16.5, 20, 0.7, 10.5, MUTED, F_BODY)

    # 右侧环绕：主图 + 轨道 + 两个卫星
    cx, cy, rx, ry, tilt = 24.4, 10.0, 5.6, 4.2, math.radians(-18)
    circle(s, cx - rx, cy - ry, rx * 2, None, line=GOLD, lw=0.9, dash="sysDash")
    for i in range(18):
        t = math.radians(i * 20)
        u, v = rx * math.cos(t), ry * math.sin(t)
        px = cx + u * math.cos(tilt) - v * math.sin(tilt) - 0.07
        py = cy + u * math.sin(tilt) + v * math.cos(tilt) - 0.07
        circle(s, px, py, 0.14, GOLD, alpha=0.45)
    circle(s, cx - 3.6, cy - 3.6, 7.2, CARD, alpha=0.95, line=GOLD, lw=1.6)
    circle(s, cx - 3.3, cy - 3.3, 6.6, "#1B4A3A", line="#7FA98B", lw=0.8)
    text_in(s.shapes[-1], ["产品实拍", "待补"], F_KAI, MUTED, 15)
    text(s, "机 · 精准施药", cx - 2.4, cy + 3.9, 8, 0.7, 11, GOLD, F_KAI, align="c")
    for deg, tag in [(198, "人 · 作业组织"), (48, "田 · 作物状态")]:
        t = math.radians(deg)
        u, v = rx * math.cos(t), ry * math.sin(t)
        px = cx + u * math.cos(tilt) - v * math.sin(tilt)
        py = cy + u * math.sin(tilt) + v * math.cos(tilt)
        d = 2.9
        c = circle(s, px - d / 2, py - d / 2, d, CARD, alpha=0.97, line=GOLD, lw=1.0)
        text_in(c, ["照片", "待补"], F_KAI, MUTED, 9)
        lw_ = 6.4
        lx = min(max(px - lw_ / 2, 16.0), W - 1.8 - lw_)
        text(s, tag, lx, py + d / 2 + 0.16, lw_, 0.55, 10, WHITE, F_KAI, align="c")
    text(s, "人、机、田协同，而不是孤立作业", 18.4, 15.95, 14.0, 0.7, 13, WHITE, F_TITLE,
         align="c", bold=True)
    text(s, "示意图 · 待替换为项目实拍", 18.4, 16.62, 14.0, 0.5, 9, MUTED, F_BODY, align="c")
    footer(s, 3)


# ---------------- P4 解决方案 ----------------
def p_solution(prs):
    s = page(prs, "split")
    nav(s, 2)
    text(s, "03", 1.62, 2.3, 2.0, 0.9, 18, GOLD, F_NUM)
    text(s, "解决方案  /  SOLUTION", 2.7, 2.35, 14, 0.7, 12, MUTED, F_BODY)
    text(s, "每个痛点，对应一个动作", 1.6, 3.3, 26, 2.0, 36, WHITE, F_TITLE, bold=True)
    text(s, "不是堆功能，是把三件事分别解决", 1.65, 5.4, 24, 0.9, 15, MUTED, F_KAI)

    pairs = [("用工找不到人", "作业队 + 调度平台", "按亩下单，作业过程可追溯"),
             ("药液浪费大", "变量施药模型", "按处方图控制流量，减少重喷漏喷"),
             ("农时抓不准", "多光谱巡检", "先定位问题区域，再决定作业范围")]
    y = 7.3
    for i, (pain, sol, note) in enumerate(pairs):
        # 左侧痛点（灰）
        text(s, f"0{i+1}", 1.7, y + 0.15, 1.6, 0.8, 15, MUTED, F_NUM)
        icon(s, ["team", "droplet", "target"][i], 3.3, y + 0.05, 1.5, "#8FB59B", 1.1)
        text(s, pain, 5.2, y + 0.1, 7.4, 0.9, 17, MUTED, F_KAI)
        # 中间箭头
        text(s, "→", 12.6, y + 0.05, 1.4, 0.9, 17, GOLD, F_BODY, align="c")
        # 右侧方案（亮）
        rect(s, 14.4, y - 0.15, 17.4, 2.3, CARD, alpha=0.7, line="#4E7F65", lw=0.6)
        text(s, sol, 15.0, y + 0.15, 16.0, 1.0, 19, WHITE, F_TITLE, bold=True)
        text(s, note, 15.0, y + 1.2, 16.0, 0.7, 11, MUTED, F_BODY)
        y += 2.85
    text(s, "创新点细节与实验数据 · 待补", 1.65, 16.0, 20, 0.7, 11, MUTED, F_BODY)
    footer(s, 4)


# ---------------- P5 成效 ----------------
def p_result(prs):
    s = page(prs, "grid")
    nav(s, 3)
    text(s, "04", 1.62, 2.3, 2.0, 0.9, 18, GOLD, F_NUM)
    text(s, "阶段成效  /  RESULTS", 2.7, 2.35, 14, 0.7, 12, MUTED, F_BODY)
    text(s, "先跑通一个县", 1.6, 3.3, 26, 2.0, 36, WHITE, F_TITLE, bold=True)
    text(s, "把样板做扎实，再谈复制", 1.65, 5.4, 24, 0.9, 15, MUTED, F_KAI)

    kpis = [("待补", "服务面积", "亩", "target"), ("待补", "合作主体", "个", "team"),
            ("待补", "单机效率", "亩/天", "gear"), ("待补", "节药比例", "%", "droplet")]
    x = 1.7
    for i, (v, lab, unit, ic) in enumerate(kpis):
        rect(s, x, 7.1, 7.4, 5.2, CARD, alpha=0.55, line="#4E7F65", lw=0.6)
        icon(s, ic, x + 0.7, 7.55, 1.5, "#8FB59B", 1.1)
        text(s, v, x + 0.6, 9.2, 6.2, 1.5, 26, GOLD, F_TITLE, bold=True)
        text(s, f"单位：{unit}", x + 0.6, 10.9, 6.2, 0.5, 9.5, MUTED, F_BODY)
        rect(s, x + 0.6, 11.5, 2.6, 0.04, GOLD, alpha=0.65)
        text(s, lab, x + 0.6, 11.75, 6.2, 0.6, 13, WHITE, F_KAI)
        x += 7.9
    text(s, "本页所有数值均为占位：请提供真实统计口径与原始表格后替换。",
         1.7, 13.2, 28, 0.8, 11.5, MUTED, F_BODY)

    # 证据条：三类凭证位
    text(s, "证据位（待补）", 1.7, 15.0, 8, 0.6, 10, GOLD, F_BODY_B)
    ev = ["合作协议", "检测报告", "现场记录"]
    x = 1.7
    for name in ev:
        circle(s, x, 15.5, 1.4, None, line=GOLD, lw=1.0, dash="sysDash")
        text(s, name, x + 1.75, 16.0, 6.0, 0.6, 10.5, MUTED, F_BODY)
        x += 9.4
    footer(s, 5)


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="design_v5.pptx")
    a = ap.parse_args()
    prs = Presentation()
    prs.slide_width, prs.slide_height = Cm(W), Cm(H)
    p_cover(prs)
    p_toc(prs)
    p_pain(prs)
    p_solution(prs)
    p_result(prs)
    prs.save(a.out)
    print(f"已生成 {a.out}（5 页）")


if __name__ == "__main__":
    main()
