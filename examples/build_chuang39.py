#!/usr/bin/env python3
"""build_chuang39.py — 创赛（互联网+）完整 30 页样本。

按 39 页标准结构（types/01-创赛互联网+.md）压缩成 30 页，四大章节。
风格：深墨绿 + 金，书法标题封面，内页 36–40pt 标题，华文楷体副标题。

用法:
    python3 build_chuang39.py --out 创赛样本.pptx --assets ./assets
"""
import argparse
import os
import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

W, H = 33.867, 19.05
BG = "#08170F"
BG2 = "#0E2A1C"
CARD = "#123526"
GOLD = "#D4A24C"
TEXT = "#F2EFE6"
MUTED = "#9DB3A5"
LINE = "#2C5A43"

FT = "/usr/share/fonts/truetype/custom/"
F_BODY = FT + "阿里巴巴普惠体 L.ttf"
F_BODY_B = FT + "阿里巴巴普惠体 B.ttf"
F_KAI = FT + "华文楷体.TTF"
F_SONG = FT + "思源宋体 CN Light.otf"
F_NUM = "/system/fonts/NotoSansCJK-Regular.ttc"


def hx(s):
    s = s.lstrip("#")
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def _alpha(shape, a):
    solid = shape._element.find(".//" + qn("a:solidFill"))
    if solid is None:
        return
    clr = solid.find(qn("a:srgbClr"))
    if clr is None:
        return
    e = etree.SubElement(clr, qn("a:alpha"))
    e.set("val", str(int(a * 100000)))


def rect(s, x, y, w, h, color=None, alpha=1.0, radius=0, line=None, lw=1.0):
    shp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
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
            _alpha(shp, alpha)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = hx(line)
        shp.line.width = Pt(lw)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def txt(s, t, x, y, w, h, size=18, color=TEXT, font=None, align="l", anchor="t",
        bold=False, spacing=1.3):
    tb = s.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    for i, ln in enumerate(str(t).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        p.line_spacing = spacing
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.name = font or F_BODY
        r.font.color.rgb = hx(color)
    return tb


def pic(s, path, x, y, w, h):
    return s.shapes.add_picture(path, Cm(x), Cm(y), Cm(w), Cm(h))


def hline(s, x, y, w, color=LINE, lw=1.0):
    c = s.shapes.add_connector(1, Cm(x), Cm(y), Cm(x + w), Cm(y))
    c.line.color.rgb = hx(color)
    c.line.width = Pt(lw)
    return c


# ---------------- 通用底 ----------------
DECOR = []   # 由 main 填充：装饰图路径


def add_decor(s, idx=0):
    """装饰层：左侧竖排小图 + 底边小图 + 图标 + 点阵 + 细线（不压内容区）。"""
    if not DECOR:
        return
    try:
        # 1 顶部极细纹理条（不侵占标题区）
        pic(s, DECOR[idx % len(DECOR)], 0, 0, W, 0.32)
        rect(s, 0, 0, W, 0.32, BG, alpha=0.55)
        # 2 左边距竖排 4 张小图
        for k in range(4):
            pic(s, DECOR[(idx + k * 3) % len(DECOR)], 0.12, 4.4 + k * 2.75, 0.95, 2.3)
        # 3 底边 3 张小图（页脚线之上）
        for k in range(3):
            pic(s, DECOR[(idx + 2 + k * 4) % len(DECOR)], 1.3 + k * 3.9, 16.05, 3.5, 0.95)
            rect(s, 1.3 + k * 3.9, 16.05, 3.5, 0.95, BG, alpha=0.25)
        # 4 图标行（页脚线之上，右侧）
        for k in range(4):
            rect(s, W - 5.6 + k * 0.85, 16.15, 0.55, 0.55, GOLD, alpha=0.14, radius=0.1)
        # 5 点阵（右上角）
        for r_ in range(3):
            for c_ in range(5):
                rect(s, W - 5.4 + c_ * 0.26, 1.15 + r_ * 0.26, 0.08, 0.08, GOLD, alpha=0.3)
        # 6 细线
        hline(s, W - 8.6, 15.95, 7.3, LINE, 0.6)
    except Exception:
        pass


def base(prs, dark=True):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, W, H, BG if dark else "#F7F5EF", radius=0)
    # 细网格
    for i in range(1, 12):
        l = s.shapes.add_connector(1, Cm(W * i / 12), Cm(0), Cm(W * i / 12), Cm(H))
        l.line.color.rgb = hx("#123024")
        l.line.width = Pt(0.4)
    # 四角装饰
    for (cx, cy, fx, fy) in [(0.6, 0.6, 1, 1), (W - 2.4, 0.6, -1, 1),
                             (0.6, H - 2.4, 1, -1), (W - 2.4, H - 2.4, -1, -1)]:
        rect(s, cx, cy, 1.8, 0.06, GOLD, alpha=0.55)
        rect(s, cx if fx > 0 else cx + 1.74, cy, 0.06, 1.8, GOLD, alpha=0.55)
    return s


NAV_CHAPTERS = ["项目背景", "技术介绍", "痛点分析", "应对策略", "团队介绍", "发展规划"]


def nav_bar(s, current_idx=-1):
    """顶部导航条：左项目名 + 右 6 个章节（当前高亮）。"""
    rect(s, 0, 0, W, 0.9, "#04120C", alpha=0.92)
    hline(s, 0, 0.9, W, GOLD, 0.6)
    txt(s, "慧飞植保 · 无人机智能植保", 1.1, 0.22, 11, 0.5, size=12, color=GOLD, bold=True)
    x = W - 1.2
    for i in range(len(NAV_CHAPTERS) - 1, -1, -1):
        name = NAV_CHAPTERS[i]
        wch = 3.05
        x -= wch
        active = (i == current_idx)
        if active:
            rect(s, x + 0.08, 0.5, wch - 0.16, 0.34, GOLD, alpha=0.22, radius=0.06)
        txt(s, name, x + 0.08, 0.5, wch - 0.16, 0.35, size=11,
            color=GOLD if active else "#8FA79B", align="c")


def head(s, chapter, title, sub="", idx=0):
    add_decor(s, idx)
    # 章节号映射到导航位置
    cur = -1
    for i, c in enumerate(["背景", "技术", "痛点", "策略", "团队", "规划"]):
        if c in chapter:
            cur = i
            break
    nav_bar(s, cur)
    txt(s, chapter, 1.3, 1.18, 14, 0.7, size=15, color=GOLD)
    txt(s, title, 1.3, 1.92, W - 3.5, 1.5, size=33, bold=True, color=TEXT)
    if sub:
        # 结论条：深色半透明横条 + 白字（国赛常见做法）
        bw = min(W - 2.6, 1.6 + len(sub) * 0.42)
        rect(s, 1.3, 3.5, bw, 1.0, "#0C2418", alpha=0.85, radius=0.08, line=GOLD, lw=0.7)
        txt(s, sub, 1.72, 3.66, bw - 0.9, 0.7, size=16, color=TEXT)


def foot(s, n, total, label="慧飞植保 · 智慧农业无人机植保"):
    hline(s, 1.3, 17.55, W - 2.6, LINE, 0.8)
    txt(s, label, 1.3, 17.75, 20, 0.7, size=11, color=MUTED)
    txt(s, f"{n:02d} / {total:02d}", W - 5.4, 17.72, 4.1, 0.7, size=11, color=MUTED, align="r")


def transition(s, kind="fade"):
    old = s._element.find(qn("p:transition"))
    if old is not None:
        s._element.remove(old)
    t = etree.SubElement(s._element, qn("p:transition"))
    t.set("spd", "med")
    etree.SubElement(t, qn("p:" + ("morph" if kind == "morph" else "fade")))


# ---------------- 页面类型 ----------------
def p_cover(prs, d, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    pic(s, d["cover_img"], 0, 0, W, H)
    transition(s, "morph")


def p_opening(prs, d, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    pic(s, d["img"], 0, 0, W, H)
    transition(s, "fade")


def p_section(prs, d, n, total):
    s = base(prs)
    rect(s, 0, 0, W, H, "#000000", alpha=0.35)
    txt(s, d["num"], 1.6, 4.4, 13, 5.2, size=110, color=GOLD, font=F_NUM, bold=True)
    txt(s, d["title"], 1.9, 9.7, 22, 2.2, size=44, bold=True, color=TEXT)
    hline(s, 1.95, 12.2, 5.0, GOLD, 4.0)
    txt(s, d.get("note", ""), 1.95, 12.7, 22, 1.4, size=22, color=MUTED, font=F_KAI)
    transition(s, "fade")



def p_toc(prs, d, n, total):
    s = base(prs)
    add_decor(s, 3)
    nav_bar(s, -1)
    txt(s, "目录", 1.3, 1.9, 6, 1.8, size=48, bold=True, color=GOLD)
    txt(s, "CONTENTS", 1.35, 3.7, 8, 0.8, size=18, color=MUTED)
    hline(s, 1.35, 4.5, 4.0, GOLD, 2.5)
    items = d["items"]
    cw, ch, gap = 7.4, 3.2, 0.7
    for i, (num, title, note) in enumerate(items[:6]):
        r, c = divmod(i, 3)
        x = 1.3 + c * (cw + gap)
        y = 5.4 + r * (ch + gap)
        rect(s, x, y, cw, ch, CARD, radius=0.16, line=LINE)
        rect(s, x, y, 0.1, ch, GOLD)
        txt(s, num, x + 0.55, y + 0.45, 2.2, 1.0, size=28, bold=True, color=GOLD, font=F_NUM)
        txt(s, title, x + 2.5, y + 0.5, cw - 3.0, 0.9, size=21, bold=True, color=TEXT)
        txt(s, note, x + 0.55, y + 1.6, cw - 1.1, 1.3, size=14, color=MUTED, spacing=1.3)
    foot(s, n, total)

def p_overview(prs, d, n, total):
    s = base(prs)
    head(s, d["chapter"], d["title"], d.get("sub", ""))
    cells = d["cells"]
    cw, ch, gap = 9.2, 4.6, 0.5
    for i, (k, v) in enumerate(cells[:5]):
        r, c = divmod(i, 3)
        x = 1.3 + c * (cw + gap)
        y = 4.9 + r * (ch + gap)
        rect(s, x, y, cw, ch, CARD, radius=0.2, line=LINE)
        rect(s, x, y, 0.12, ch, GOLD)
        txt(s, k, x + 0.5, y + 0.55, cw - 1, 0.8, size=15, color=GOLD)
        txt(s, v, x + 0.5, y + 1.45, cw - 1, ch - 1.9, size=16, color=TEXT, spacing=1.35)
    foot(s, n, total)


def p_media(prs, d, n, total, img_left=True):
    s = base(prs)
    head(s, d["chapter"], d["title"], d.get("sub", ""))
    iw = 13.2
    ih = 11.0
    if img_left:
        pic(s, d["img"], 1.3, 4.9, iw, ih)
        rect(s, 1.3, 4.9, iw, ih, None, line=GOLD, lw=1.2)
        tx = 1.3 + iw + 1.0
    else:
        pic(s, d["img"], W - 1.3 - iw, 4.9, iw, ih)
        rect(s, W - 1.3 - iw, 4.9, iw, ih, None, line=GOLD, lw=1.2)
        tx = 1.3
    tw = W - 2.6 - iw - 1.0
    y = 5.2
    for p in d["points"][:4]:
        rect(s, tx, y + 0.25, 0.42, 0.42, GOLD, radius=0.5)
        txt(s, p, tx + 0.85, y, tw - 0.9, 2.4, size=20, color=TEXT, spacing=1.45)
        y += 2.7
    foot(s, n, total)


def p_funnel(prs, d, n, total):
    s = base(prs)
    head(s, d["chapter"], d["title"], d.get("sub", ""))
    layers = d["layers"]
    widths = [26, 20, 14]
    y = 5.0
    for i, (t, v) in enumerate(layers):
        w = widths[i]
        x = (W - w) / 2
        h = 3.3
        rect(s, x, y, w, h, CARD if i < 2 else "#1A4030", radius=0.18, line=GOLD if i == 2 else LINE)
        txt(s, t, x + 0.6, y + 0.35, w - 1.2, 0.7, size=14, color=GOLD, align="c")
        txt(s, v, x + 0.6, y + 1.1, w - 1.2, h - 1.4, size=18 if i < 2 else 19,
            color=TEXT, align="c", bold=(i == 2), spacing=1.3)
        if i < 2:
            # 箭头
            ax = W / 2
            txt(s, "▼", ax - 0.4, y + h + 0.02, 0.8, 0.5, size=14, color=GOLD, align="c")
        y += h + 0.6
    foot(s, n, total)


def p_cards(prs, d, n, total, cols=3):
    s = base(prs)
    head(s, d["chapter"], d["title"], d.get("sub", ""))
    items = d["cards"]
    gap = 0.7
    cw = (W - 2.6 - gap * (cols - 1)) / cols
    ch = 9.8
    y = 4.9
    for i, (t, body, tag) in enumerate(items[:cols]):
        x = 1.3 + i * (cw + gap)
        rect(s, x, y, cw, ch, CARD, radius=0.2, line=LINE)
        rect(s, x, y, cw, 0.12, GOLD)
        txt(s, f"{i+1:02d}", x + 0.6, y + 0.7, 2, 0.9, size=18, color=GOLD, font=F_NUM, bold=True)
        txt(s, t, x + 0.6, y + 1.8, cw - 1.2, 1.2, size=24, bold=True, color=TEXT)
        txt(s, body, x + 0.6, y + 3.3, cw - 1.2, ch - 4.6, size=16, color=TEXT, spacing=1.4)
        if tag:
            rect(s, x + cw - 2.8, y + ch - 1.3, 2.3, 0.75, GOLD, alpha=0.18, radius=0.2)
            txt(s, tag, x + cw - 2.8, y + ch - 1.18, 2.3, 0.6, size=14, color=GOLD, align="c")
    foot(s, n, total)


def p_innovate(prs, d, n, total):
    s = base(prs)
    head(s, d["chapter"], d["title"], d.get("sub", ""))
    rows = d["rows"]
    y = 4.9
    for i, (pain, sol) in enumerate(rows):
        rect(s, 1.3, y, 12.0, 2.5, CARD, radius=0.18, line=LINE)
        txt(s, "痛点", 1.7, y + 0.35, 2, 0.6, size=13, color=MUTED)
        txt(s, pain, 1.7, y + 0.95, 11.2, 1.2, size=19, color=TEXT)
        txt(s, "→", 13.55, y + 0.75, 1.2, 1.0, size=26, color=GOLD, align="c")
        rect(s, 14.9, y, 17.6, 2.5, "#1A4030", radius=0.18, line=GOLD, lw=1.2)
        txt(s, f"创新点 {i+1}", 15.3, y + 0.35, 3, 0.6, size=13, color=GOLD)
        txt(s, sol, 15.3, y + 0.95, 16.8, 1.2, size=19, color=TEXT)
        y += 2.9
    foot(s, n, total)


def p_kpi(prs, d, n, total):
    s = base(prs)
    head(s, d["chapter"], d["title"], d.get("sub", ""))
    items = d["items"]
    gap = 1.0
    cw = (W - 2.6 - gap * (len(items) - 1)) / len(items)
    for i, (v, u, lab) in enumerate(items):
        x = 1.3 + i * (cw + gap)
        y = 5.2
        rect(s, x, y, cw, 9.6, CARD, radius=0.22, line=LINE)
        txt(s, v, x, y + 1.6, cw, 3.2, size=68, color=GOLD, font=F_NUM, bold=True, align="c")
        txt(s, u, x, y + 5.2, cw, 0.9, size=18, color=MUTED, align="c")
        hline(s, x + cw * 0.3, y + 6.6, cw * 0.4, LINE, 1.0)
        txt(s, lab, x + 0.6, y + 7.1, cw - 1.2, 1.8, size=17, color=TEXT, align="c", spacing=1.35)
    foot(s, n, total)


def p_table(prs, d, n, total):
    s = base(prs)
    head(s, d["chapter"], d["title"], d.get("sub", ""))
    rows = d["table"]
    ncol = len(rows[0])
    colw = (W - 2.6) / ncol
    y = 5.0
    rowh = 1.9
    for ri, row in enumerate(rows):
        for ci, cell in enumerate(row):
            x = 1.3 + ci * colw
            fill = "#17392A" if ri == 0 else (CARD if ri % 2 else "#0F2D1F")
            rect(s, x, y, colw - 0.05, rowh - 0.05, fill, radius=0.06,
                 line=GOLD if (ri == 0 or ci == 0) else LINE)
            c = GOLD if (ri == 0 or ci == 0) else TEXT
            txt(s, cell, x + 0.3, y + 0.35, colw - 0.6, rowh - 0.7,
                size=17 if ri else 18, color=c, anchor="m",
                bold=(ri == 0 or ci == 0))
        y += rowh
    foot(s, n, total)


def p_chain(prs, d, n, total):
    s = base(prs)
    head(s, d["chapter"], d["title"], d.get("sub", ""))
    nodes = d["nodes"]
    n = len(nodes)
    gap = 0.9
    cw = (W - 2.6 - gap * (n - 1)) / n
    y = 6.2
    for i, (t, sub) in enumerate(nodes):
        x = 1.3 + i * (cw + gap)
        rect(s, x, y, cw, 6.4, CARD, radius=0.22, line=GOLD if i == 0 else LINE, lw=1.2)
        txt(s, t, x + 0.5, y + 0.7, cw - 1, 2.0, size=19, bold=True, color=TEXT,
            align="c", spacing=1.3)
        txt(s, sub, x + 0.5, y + 3.0, cw - 1, 3.0, size=16, color=MUTED, align="c", spacing=1.4)
        if i < n - 1:
            txt(s, "▶", x + cw + 0.1, y + 2.8, 0.7, 0.8, size=18, color=GOLD, align="c")
    foot(s, n, total)


def p_cases(prs, d, n, total):
    s = base(prs)
    head(s, d["chapter"], d["title"], d.get("sub", ""))
    items = d["cases"]
    gap = 0.7
    cw = (W - 2.6 - gap * 2) / 3
    y = 5.0
    for i, (place, before, after) in enumerate(items[:3]):
        x = 1.3 + i * (cw + gap)
        rect(s, x, y, cw, 9.9, CARD, radius=0.2, line=LINE)
        rect(s, x, y, cw, 1.6, "#1A4030", radius=0.2)
        txt(s, place, x + 0.5, y + 0.4, cw - 1, 0.9, size=19, bold=True, color=GOLD)
        txt(s, "改造前", x + 0.5, y + 2.1, 3, 0.6, size=13, color=MUTED)
        txt(s, before, x + 0.5, y + 2.75, cw - 1, 2.4, size=17, color=TEXT, spacing=1.35)
        txt(s, "使用后", x + 0.5, y + 5.5, 3, 0.6, size=13, color=GOLD)
        txt(s, after, x + 0.5, y + 6.15, cw - 1, 3.0, size=17, color=TEXT, spacing=1.35)
    foot(s, n, total)


def p_team(prs, d, n, total):
    s = base(prs)
    head(s, d["chapter"], d["title"], d.get("sub", ""))
    members = d["members"]
    gap = 0.6
    cw = (W - 2.6 - gap * 3) / 4
    y = 5.0
    for i, (name, role, exp) in enumerate(members[:8]):
        r, c = divmod(i, 4)
        x = 1.3 + c * (cw + gap)
        yy = y + r * 5.6
        rect(s, x, yy, cw, 5.1, CARD, radius=0.18, line=LINE)
        # 头像占位：金色圆
        cir = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(x + cw / 2 - 0.8), Cm(yy + 0.5), Cm(1.6), Cm(1.6))
        cir.fill.solid()
        cir.fill.fore_color.rgb = hx(GOLD)
        cir.line.fill.background()
        txt(s, name[0], x + cw / 2 - 0.8, yy + 0.85, 1.6, 0.9, size=22, color=BG,
            align="c", bold=True)
        txt(s, name, x + 0.3, yy + 2.3, cw - 0.6, 0.8, size=17, bold=True, color=TEXT, align="c")
        txt(s, role, x + 0.3, yy + 3.0, cw - 0.6, 0.7, size=13, color=GOLD, align="c")
        txt(s, exp, x + 0.3, yy + 3.65, cw - 0.6, 1.3, size=14, color=MUTED, align="c", spacing=1.25)
    foot(s, n, total)


def p_certs(prs, d, n, total):
    s = base(prs)
    head(s, d["chapter"], d["title"], d.get("sub", ""))
    items = d["certs"]
    gap = 0.5
    cw = (W - 2.6 - gap * 4) / 5
    y = 5.2
    for i, label in enumerate(items[:10]):
        r, c = divmod(i, 5)
        x = 1.3 + c * (cw + gap)
        yy = y + r * 5.2
        rect(s, x, yy, cw, 4.7, "#F2EFE6", radius=0.08, line=GOLD, lw=1.2)
        hline(s, x + 0.35, yy + 1.1, cw - 0.7, "#C9BFA6", 0.8)
        hline(s, x + 0.35, yy + 3.3, cw - 0.7, "#C9BFA6", 0.8)
        txt(s, "证书", x + 0.2, yy + 0.35, cw - 0.4, 0.6, size=12, color="#8A7F68", align="c")
        txt(s, label, x + 0.3, yy + 1.5, cw - 0.6, 1.7, size=14, color="#2B3A2E",
            align="c", bold=True, spacing=1.25)
        txt(s, "国家知识产权局", x + 0.2, yy + 3.6, cw - 0.4, 0.6, size=10,
            color="#8A7F68", align="c")
    foot(s, n, total)


def p_timeline(prs, d, n, total):
    s = base(prs)
    head(s, d["chapter"], d["title"], d.get("sub", ""))
    items = d["items"]
    n0 = len(items)
    y_axis = 9.6
    pad_x = 3.2
    hline(s, pad_x, y_axis, W - pad_x * 2, GOLD, 2.5)
    seg = (W - pad_x * 2) / max(n0 - 1, 1)
    for i, (t, title, note) in enumerate(items):
        x = pad_x + seg * i
        c = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(x - 0.3), Cm(y_axis - 0.3), Cm(0.6), Cm(0.6))
        c.fill.solid()
        c.fill.fore_color.rgb = hx(GOLD)
        c.line.fill.background()
        up = i % 2 == 0
        ty = y_axis - 3.6 if up else y_axis + 0.9
        txt(s, t, x - 2.2, ty, 4.4, 0.7, size=16, color=GOLD, align="c", bold=True)
        txt(s, title, x - 2.4, ty + 0.8, 4.8, 1.0, size=20, color=TEXT, align="c", bold=True)
        txt(s, note, x - 2.4, ty + 1.9, 4.8, 1.6, size=14, color=MUTED, align="c", spacing=1.3)
    foot(s, n, total)


def p_end(prs, d, n, total):
    s = base(prs)
    rect(s, 0, 0, W, H, "#000000", alpha=0.4)
    txt(s, d["title"], 0, 6.2, W, 3.4, size=56, bold=True, color=TEXT, align="c")
    rect(s, W / 2 - 1.6, 10.2, 3.2, 0.14, GOLD)
    txt(s, d.get("note", ""), 2, 11.0, W - 4, 1.6, size=22, color=MUTED, align="c", font=F_KAI)


# ---------------- 内容 ----------------
def deck_spec(a):
    A = lambda name: os.path.join(a.assets, name)
    return [
        ("p_opening", dict(img=A("open1.jpg"))),
        ("p_opening", dict(img=A("open2.jpg"))),
        ("p_opening", dict(img=A("open3.jpg"))),
        ("p_cover", dict(cover_img=A("cover.jpg"))),

        ("p_toc", dict(items=[("01", "项目背景", "政策切口 · 社会痛点 · 技术痛点 · 主人公"),
                              ("02", "解决方案", "三个创新点对应三个痛点"),
                              ("03", "商业逻辑", "三方评价 · 知识产权 · 产品 · 竞品 · 模式"),
                              ("04", "落地方案", "中试验证 · 帮扶案例 · 市场推广"),
                              ("05", "团队", "核心团队 · 专家顾问 · 凭证 · 财务"),
                              ("06", "社会成效", "社会效益 · 媒体报道 · 教育维度 · 规划")])),
        ("p_section", dict(num="01", title="项目背景故事", note="从政策到痛点，从痛点到技术")),
        ("p_overview", dict(chapter="01 · 项目概述", title="一页看懂：我们做什么",
                            sub="200 字说不清的，就不是好项目",
                            cells=[("做什么", "给合作社和种粮大户\n提供无人机植保服务\n按亩收费"),
                                   ("给谁", "50 岁以上农户占 62%\n缺劳力、缺技术"),
                                   ("怎么赚", "服务费 60 元/亩\n比人工省 130 元/亩"),
                                   ("壁垒", "本地作业队 + 变量施药模型\n+ 1200 组田间数据"),
                                   ("进展", "2025 年服务 120 个合作社\n营收 380 万")])),
        ("p_media", dict(chapter="01 · 项目背景①", title="大政策里，我们切的是这一块",
                         sub="不是所有农业都适合无人机，我们只做水稻植保",
                         img=A("field.jpg"),
                         points=["《全国智慧农业行动计划》明确 2028 年植保无人飞机保有量 25 万架",
                                 "东北水稻区连片 1000 亩以上地块占比 38%，最适合飞防",
                                 "我们锁定辽宁、吉林两省的水稻主产区，先做透一个县"])),
        ("p_media", dict(chapter="01 · 项目背景②", title="农村没人打药了",
                         sub="不是农民懒，是真的找不到人",
                         img="assets/p_farmer.jpg",
                         points=["辽宁水稻区 60 岁以上劳动力占 23%，青壮年外出务工",
                                 "一亩地打药 40 分钟，一个劳力一天最多干 5 亩",
                                 "农忙窗口只有 7–10 天，错过就要减药效"])),
        ("p_funnel", dict(chapter="01 · 项目背景③", title="从社会痛点，推到技术痛点",
                          sub="痛点一层层往下掉，最后落在技术上",
                          layers=[("社会痛点", "农村劳动力 60 岁以上占 23%，打药找不到人"),
                                  ("产业痛点", "人工喷洒农药利用率仅 35%，每亩浪费 42 元"),
                                  ("技术痛点", "现有植保机雾滴粗、无变量控制，重喷漏喷严重")])),
        ("p_media", dict(chapter="01 · 痛点", title="一亩地，人工打药要 40 分钟",
                         sub="三个数字，就是我们要解决的问题",
                         img=A("pain.jpg"),
                         points=["打药靠人扛：背负式喷雾器日均作业 3–5 亩，50 岁以上农户占 62%",
                                 "雾滴粗、浪费大：人工喷雾粒径 200μm 以上，农药流失率超 40%",
                                 "不会看虫情：凭经验施药，错过最佳窗口，平均减产 8–12%"],
                         img_left=True)),
        ("p_media", dict(chapter="01 · 主人公①", title="这事为什么是我们干",
                         sub="不是想出来的，是干出来的",
                         img=A("hero_scene.jpg"),
                         points=["2023 年入学，从画第一块 STM32 最小系统板开始",
                                 "2024 年做物联网智能家居，跑通传感器与远程控制"],
                         img_left=False)),
        ("p_cards", dict(chapter="01 · 主人公②", title="三年，四件事，都是真项目",
                         sub="每一个里程碑都能拿出实物或证书",
                         cards=[("自制主控板", "10×10cm，原理图到打样全自己来", "2024"),
                                ("NAO 全国一等奖", "仿人单人舞蹈，负责动作编程", "2025"),
                                ("YOLOv11 识别", "文冠果虫害，手标 1200 张图", "2026")])),
        ("p_innovate", dict(chapter="01 · 解决方案", title="三个创新点，对应三个痛点",
                            sub="痛点 1 → 创新点 1，一一对应，不多不少",
                            rows=[("人工打药效率低、找不到人", "多旋翼植保机 + 作业队模式，8 分钟/亩"),
                                  ("雾滴粗、农药浪费 40%", "离心雾化喷头，雾滴 80–120μm，利用率 70%"),
                                  ("看不见虫情，凭经验打药", "多光谱巡检先定位虫害区，只打有虫地块")])),
        ("p_cards", dict(chapter="01 · 创新点①", title="理念：把「打药」变成「精准投递」",
                         sub="不是把喷雾器搬到天上，是重新定义施药",
                         cards=[("按需施药", "先巡检出虫害分布图，再按图施药，不是整块地均匀喷", "理念"),
                                ("雾滴可控", "离心雾化，粒径可在 80–150μm 之间调，按作物调整", "原理"),
                                ("作业可追溯", "每架次飞行轨迹、用药量、作业面积自动回传", "设计")])),
        ("p_cards", dict(chapter="01 · 创新点②", title="技术：三个自研模块，成本降 60%",
                         sub="核心部件国产化，不依赖进口飞控",
                         cards=[("自研飞控", "基于 STM32 的植保专用飞控，抗风等级 5 级", "硬件"),
                                ("视觉识别", "YOLOv11 轻量化模型，边缘端识别虫害区域", "算法"),
                                ("变量喷洒", "按处方图实时调节流量，重喷率从 18% 降到 4%", "控制")])),
        ("p_cards", dict(chapter="01 · 创新点③", title="工艺：把实验室的东西做成能下地的",
                         sub="不是能飞就行，要能连续作业 8 小时",
                         cards=[("模块化换药", "药箱 30 秒快拆，一天换 12 次不堵", "工艺"),
                                ("抗风作业", "5 级风下航线偏移 < 0.3m", "方案"),
                                ("维护成本低", "整机 8.8 万，年维护 3000 元", "成熟度")])),
        ("p_kpi", dict(chapter="01 · 中试验证", title="在 3 个合作社，跑了 1200 亩",
                       sub="不是实验室数据，是田间实测",
                       items=[("8", "分钟/亩", "对比人工 40 分钟"),
                              ("70%", "农药利用率", "行业平均 35%"),
                              ("130", "元/亩 省", "含农药与人工")])),

        ("p_section", dict(num="02", title="商业与运营逻辑", note="怎么赚钱，怎么落地，怎么复制")),
        ("p_media", dict(chapter="02 · 三方评价", title="三方都认可，才算真落地",
                         sub="技术方、行业方、受益者，各一句",
                         img=A("field2.jpg"),
                         points=["技术方：沈阳农业大学植保学院 — 「变量施药精度达到行业先进水平」",
                                 "行业方：辽宁某农服公司 — 「单机日作业量比同类高 20%」",
                                 "受益者：盘锦合作社王社长 — 「去年省了 6 万块农药钱」"],
                         img_left=False)),
        ("p_certs", dict(chapter="02 · 知识产权", title="2 项专利、3 项软著，都在受理或授权",
                         sub="证书原件可随时提供",
                         certs=["发明专利申请", "实用新型专利", "软件著作权①", "软件著作权②",
                                "软件著作权③", "查新报告", "检测报告", "田间试验报告",
                                "合作协议①", "合作协议②"])),
        ("p_table", dict(chapter="02 · 产品与服务", title="一台机器，三种服务",
                         sub="卖设备、卖服务、卖数据，三条线并行",
                         table=[["服务类型", "内容", "定价", "毛利"],
                                ["飞防作业", "按亩收费，含药剂", "60 元/亩", "35%"],
                                ["设备销售", "植保机整机 + 培训", "8.8 万/台", "28%"],
                                ["数据服务", "虫情监测 + 处方图", "1.2 万/年", "62%"]])),
        ("p_table", dict(chapter="02 · 竞品分析", title="同类产品参数对比",
                         sub="不比价格比综合成本，不比参数比作业效果",
                         table=[["维度", "人工", "传统植保机", "慧飞（我们）"],
                                ["作业效率", "3–5 亩/天", "200 亩/天", "300 亩/天"],
                                ["雾滴粒径", "200μm+", "150–200μm", "80–120μm"],
                                ["农药利用率", "35%", "50%", "70%"],
                                ["综合成本", "190 元/亩", "95 元/亩", "60 元/亩"]])),
        ("p_chain", dict(chapter="02 · 商业模式", title="上游、中游、下游，我们卡在哪",
                         sub="不做重资产，做作业网络",
                         nodes=[("上游", "整机厂 / 药剂厂\n集采降本"),
                                ("中游\n（我们）", "作业队 + 调度平台\n按亩收费"),
                                ("下游", "合作社 / 种粮大户\n按效果付费")])),
        ("p_cases", dict(chapter="02 · 帮扶案例①", title="盘锦：2000 亩，20 人的活 3 天干完",
                         sub="规模合作社，最看重效率和成本",
                         cases=[("改造前", "人工打药 20 人干 10 天，药费 8 万", "人均日作业 5 亩"),
                                ("改造后", "2 台机器 3 天干完，药费 5.2 万", "单机日作业 330 亩"),
                                ("账本", "省药费 2.8 万 + 省人工 3.2 万 = 年省 6 万", "投入 17.6 万，3 年回本")])),
        ("p_cases", dict(chapter="02 · 帮扶案例②", title="营口：请不到人，就买服务",
                         sub="家庭农场，最看重省心",
                         cases=[("改造前", "缺劳力，往年请人打药要排队等", "错过窗口减产 8%"),
                                ("改造后", "按亩买服务，一年 4 次，共 2.4 万", "手机下单，48 小时上门"),
                                ("账本", "比自己买机器省 15 万投入", "按次付费，用多少付多少")])),
        ("p_cases", dict(chapter="02 · 帮扶案例③", title="铁岭：村集体有了自己的作业队",
                         sub="村集体，最看重带动就业",
                         cases=[("改造前", "想发展集体经济，但没技术没设备", "村集体收入 3 万/年"),
                                ("改造后", "培训 6 名飞手，成立作业队", "服务周边 8 个村"),
                                ("账本", "年作业 1.2 万亩次，净收入 18 万", "带动 6 人就业")])),
        ("p_kpi", dict(chapter="02 · 市场推广", title="已经跑通的三个渠道",
                       sub="不铺广告，靠示范田和口碑",
                       items=[("120", "个合作社", "2025 年累计服务"),
                              ("3", "个示范县", "盘锦 / 营口 / 铁岭"),
                              ("380", "万元", "2025 年营收")])),

        ("p_section", dict(num="03", title="组织与团队", note="谁在做，凭什么做成")),
        ("p_team", dict(chapter="03 · 核心团队", title="6 个人，专业互补",
                        sub="1 名负责人 + 5 名专业成员",
                        members=[("李明", "项目负责人", "嵌入式开发 3 年"),
                                 ("王强", "硬件负责人", "PCB 设计 / 飞控"),
                                 ("张丽", "算法负责人", "YOLO / 视觉识别"),
                                 ("赵磊", "作业队长", "持证飞手 / 1200 亩"),
                                 ("陈静", "运营负责人", "合作社对接"),
                                 ("刘洋", "财务负责人", "成本与融资")])),
        ("p_cards", dict(chapter="03 · 专家顾问", title="三位顾问，各管一段",
                         sub="不是挂名，是真的参与",
                         cards=[("张教授", "沈阳农大植保学院，变量施药算法指导", "技术"),
                                ("王工", "某农机企业总工，结构与可靠性把关", "工程"),
                                ("李总", "农服公司创始人，合作社渠道对接", "商业")])),
        ("p_table", dict(chapter="03 · 协议与凭证", title="已经签下的，和已经收到的",
                         sub="合同、发票、流水，都可以现场核验",
                         table=[["凭证类型", "数量", "金额", "状态"],
                                ["服务合同", "12 份", "286 万", "已执行"],
                                ["意向书", "8 份", "180 万", "待签约"],
                                ["银行流水", "—", "312 万", "可查"],
                                ["发票", "46 张", "218 万", "已开具"]])),
        ("p_kpi", dict(chapter="03 · 财务", title="5 年，做到 1.2 亿营收",
                       sub="2026 年 800 万，2028 年 5000 万，2030 年 1.2 亿",
                       items=[("800", "万元（2026）", "现有订单 + 新签"),
                              ("5000", "万元（2028）", "复制到 10 个县"),
                              ("30%", "净利率", "服务规模化后")])),

        ("p_section", dict(num="04", title="社会成效与规划", note="不只是生意，也是答案")),
        ("p_media", dict(chapter="04 · 领导关怀", title="市领导两次到示范田",
                         sub="不是走过场，是现场看作业",
                         img="assets/p_coop.jpg",
                         points=["2025 年 6 月，市农业农村局到盘锦示范田调研",
                                 "2025 年 9 月，省农机推广站现场观摩作业",
                                 "被列为市级「智慧农业示范项目」"])),
        ("p_cards", dict(chapter="04 · 媒体报道", title="三家媒体，三次报道",
                         sub="报道可查，随时可核",
                         cards=[("辽宁日报", "《无人机植保，让打药不再靠人扛》", "2025.07"),
                                ("农民日报", "《一支青年团队的水稻飞防实验》", "2025.09"),
                                ("央视三农", "《智慧农业新装备》专题片", "2026.03")])),
        ("p_cards", dict(chapter="04 · 地方反馈", title="两封感谢信，一面锦旗",
                         sub="比奖状更有说服力",
                         cards=[("盘锦合作社", "感谢解决农忙用工荒，附感谢信扫描件", "2025.08"),
                                ("铁岭村集体", "感谢培训飞手，附锦旗照片", "2025.11"),
                                ("营口家庭农场", "感谢及时上门作业，附签字回执", "2026.01")])),
        ("p_kpi", dict(chapter="04 · 社会效益", title="四个数字，说明白了",
                       sub="数据化呈现，不喊口号",
                       items=[("1200", "亩 减药", "农药使用量下降 32%"),
                              ("86", "人 就业", "飞手 / 地勤 / 调度"),
                              ("2.4", "万 亩次", "累计作业面积"),
                              ("380", "万元", "带动合作社增收")])),
        ("p_cards", dict(chapter="04 · 教育维度①", title="教育：项目变成了课",
                         sub="从实验室走进课堂",
                         cards=[("案例进课堂", "《嵌入式系统设计》2 届学生使用", "课程"),
                                ("教改项目", "省级教改立项 1 项", "教改"),
                                ("教学成果奖", "校级教学成果二等奖", "成果")])),
        ("p_cards", dict(chapter="04 · 教育维度②", title="科技：成果往一线走",
                         sub="不是发完论文就结束",
                         cards=[("专利转化", "2 项专利作价入股，与校资产公司合作", "转化"),
                                ("技术推广", "向 3 家农服公司输出作业标准", "推广"),
                                ("产学研", "与沈阳农大共建试验基地", "合作")])),
        ("p_cards", dict(chapter="04 · 教育维度③", title="人才：6 个人怎么长出来的",
                         sub="项目是育人载体，不是干活工具",
                         cards=[("3 人保研", "继续做农机方向研究", "升学"),
                                ("2 人创业", "成立农服公司，已接单", "创业"),
                                ("1 人入职", "进入农机企业做飞控", "就业")])),
        ("p_timeline", dict(chapter="04 · 发展规划", title="三步走，先把手头做透",
                            sub="2026 打样板，2028 复制省，2030 出海",
                            items=[("2026", "做透一个县", "3 个示范点，营收 800 万"),
                                   ("2027", "省内复制", "10 个县，营收 2500 万"),
                                   ("2028", "跨省扩张", "东北三省，营收 5000 万"),
                                   ("2029", "服务标准化", "输出作业标准与培训体系"),
                                   ("2030", "出海", "东南亚水稻区，营收 1.2 亿")])),
        ("p_end", dict(title="谢谢", note="慧飞植保 · 让每一亩地都用上无人机")),
    ]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="创赛样本.pptx")
    ap.add_argument("--assets", default="assets")
    a = ap.parse_args()
    global DECOR
    deco = [os.path.join(a.assets, n) for n in ("p_texture.jpg", "p_chipline.jpg", "p_land.jpg")]
    DECOR = [d for d in deco if os.path.exists(d)]
    spec = deck_spec(a)
    prs = Presentation()
    prs.slide_width = Cm(W)
    prs.slide_height = Cm(H)
    total = len(spec)
    for i, (fn, d) in enumerate(spec, 1):
        if fn == "p_media":
            d = dict(d)
            fn = lambda prs_, d_, n_, t_, dd=d: p_media(prs_, dd, n_, t_, img_left=dd.get("img_left", True))
        else:
            fn = globals()[fn]
        fn(prs, d, i, total)
    prs.save(a.out)
    print(f"已生成 {a.out}（{total} 页，{os.path.getsize(a.out)/1024/1024:.1f} MB）")


if __name__ == "__main__":
    main()
