#!/usr/bin/env python3
"""deck43.py — 用设计系统铺满 43 页（缺素材处留占位，不编数据）。"""
import os, sys, math
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deck_v5 import (W, H, BG, CARD, GOLD, WHITE, MUTED, LINE, F_TITLE, F_BRUSH, F_KAI,
                     F_BODY, F_BODY_B, F_EN, F_NUM, ARTS, ASSETS,
                     hx, rect, circle, shape, text, text_in, icon, page, nav, footer)
from pptx import Presentation
from pptx.util import Cm
from pptx.enum.shapes import MSO_SHAPE

N = [0]
def head(s, num, eyebrow, title, sub="", brush=None):
    text(s, num, 1.8, 2.2, 2.0, 0.9, 18, GOLD, F_NUM)
    text(s, eyebrow, 2.9, 2.25, 16, 0.7, 12, MUTED, F_BODY)
    text(s, title, 1.8, 3.15, 26, 2.0, 34, WHITE, F_TITLE, bold=True)
    if brush:
        text(s, brush, 1.78, 4.75, 26, 1.6, 30, GOLD, F_BRUSH)
        text(s, sub, 1.85, 6.35, 24, 0.8, 14, MUTED, F_KAI)
    elif sub:
        text(s, sub, 1.85, 4.95, 24, 0.8, 15, MUTED, F_KAI)

def art(s, name, x, y, d):
    p = os.path.join(ARTS, name)
    if os.path.exists(p):
        return s_pic(s, p, x, y, d)

def s_pic(s, p, x, y, d):
    return s.shapes.add_picture(p, Cm(x), Cm(y), Cm(d), Cm(d))

def ph(s, x, y, w, h, label="待补素材", note=""):
    c = shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, CARD, 0.5, GOLD, 0.9)
    try:
        from pptx.oxml.ns import qn
        from lxml import etree
        ln = c._element.spPr.find(qn('a:ln'))
        d = etree.SubElement(ln, qn('a:prstDash')); d.set('val', 'sysDash')
        for o in ln.findall(qn('a:prstDash'))[:-1]: ln.remove(o)
    except Exception: pass
    text_in(c, [label] + ([note] if note else []), F_KAI, MUTED, 11)

def p_section(prs, num, title, note, variant="plain", reason="章节页：重背景制造停顿，与内容页拉开层次"):
    s = page(prs, variant); nav(s, -1)
    text(s, num, 2.0, 5.0, 12, 4.2, 96, GOLD, F_NUM, bold=True)
    text(s, title, 2.2, 9.6, 22, 1.8, 34, WHITE, F_TITLE, bold=True)
    rect(s, 2.3, 11.6, 5.0, 0.05, GOLD)
    text(s, note, 2.3, 12.1, 22, 0.9, 15, MUTED, F_KAI)
    footer(s, N[0])

def p_media(prs, num, eyebrow, title, sub, art_name, points, variant="plain", left=True, reason="图文页：大图为主体，背景退让用纯色+"  "淡纹理"):
    s = page(prs, variant); nav(s, -1)
    head(s, num, eyebrow, title, sub)
    iw = 12.0
    ix, tx = (1.8, 15.0) if left else (20.0, 1.8)
    ph(s, ix, 7.4, iw, 7.4, "现场照片", "待补")
    if art_name: art(s, art_name, ix + 1.4, 7.8, 9.2)
    y = 7.5
    for p in points[:4]:
        rect(s, tx, y + 0.2, 0.3, 0.3, GOLD, 0.85)
        text(s, p, tx + 0.6, y, 14.0, 1.6, 14, WHITE, F_BODY)
        y += 2.0
    footer(s, N[0])

def p_cards3(prs, num, eyebrow, title, sub, cards, arts, variant="plain", reason="卡片页：内容密度高，背景压到最轻"):
    s = page(prs, variant); nav(s, -1)
    head(s, num, eyebrow, title, sub)
    x = 1.8
    for i, (t, body) in enumerate(cards):
        rect(s, x, 7.2, 9.6, 8.2, CARD, 0.55, "#4E7F65", 0.6)
        if arts and i < len(arts) and arts[i]:
            art(s, arts[i], x + 2.6, 7.6, 4.0)
        text(s, t, x + 0.6, 12.1, 8.4, 0.9, 15, WHITE, F_TITLE, bold=True)
        text(s, body, x + 0.6, 13.1, 8.4, 2.0, 11.5, MUTED, F_BODY)
        x += 10.2
    footer(s, N[0])

def p_rows3(prs, num, eyebrow, title, sub, rows, variant="split", reason="痛点→动作页：用上下分区表达从问题到解决"):
    s = page(prs, variant); nav(s, -1)
    head(s, num, eyebrow, title, sub)
    for cx, h in ((1.8, "现状"), (8.6, "动作"), (15.4, "变化")):
        text(s, h, cx, 7.1, 6.0, 0.6, 10.5, GOLD, F_BODY_B)
        rect(s, cx, 7.75, 6.4, 0.03, "#4E7F65")
    y = 8.2
    for i, (a, b, c, ic) in enumerate(rows):
        icon(s, ic, 1.8, y + 0.55, 1.2, "#8FB59B", 1.0)
        text(s, a, 3.3, y + 0.6, 4.6, 0.8, 14, MUTED, F_KAI)
        rect(s, 8.6, y, 6.4, 2.3, CARD, 0.6, "#4E7F65", 0.6)
        text(s, b, 9.0, y + 0.55, 5.6, 1.1, 15.5, WHITE, F_TITLE, bold=True)
        text(s, c, 15.4, y + 0.6, 6.4, 1.4, 11, MUTED, F_BODY)
        y += 2.55
    footer(s, N[0])

def p_kpi4(prs, num, eyebrow, title, sub, items, variant="grid", reason="数据页：网格纹理暗示数据，底部压条稳住视线"):
    s = page(prs, variant); nav(s, -1)
    head(s, num, eyebrow, title, sub)
    x = 1.8
    for v, lab, unit, ic in items:
        rect(s, x, 7.4, 5.6, 5.0, CARD, 0.5, "#4E7F65", 0.6)
        icon(s, ic, x + 0.55, 7.75, 1.2, "#8FB59B", 0.9)
        text(s, v, x + 0.5, 9.2, 4.6, 1.2, 22, GOLD, F_TITLE, bold=True)
        text(s, f"单位：{unit}", x + 0.5, 10.7, 4.6, 0.5, 8.5, MUTED, F_BODY)
        text(s, lab, x + 0.5, 11.4, 4.8, 0.6, 11.5, WHITE, F_KAI)
        x += 6.3
    rect(s, 1.8, 13.0, 24.8, 1.0, "#0B2018", 0.9, "#3E6E56", 0.5)
    text(s, "统计口径待补：按作业季累计，以书面确认为准", 2.2, 13.3, 24.0, 0.6, 10, MUTED, F_BODY)
    footer(s, N[0])

def p_table(prs, num, eyebrow, title, sub, rows, variant="plain", reason="表格页：信息为主，背景纯色减少干扰"):
    s = page(prs, variant); nav(s, -1)
    head(s, num, eyebrow, title, sub)
    ncol = len(rows[0]); cw = 30.0 / ncol; y = 7.2
    for ri, row in enumerate(rows):
        for ci, cell in enumerate(row):
            x = 1.8 + ci * cw
            fill = CARD if ri == 0 else ("#0F2E22" if ri % 2 else "#12362A")
            rect(s, x, y, cw - 0.08, 1.5, fill, 0.95, GOLD if ri == 0 or ci == 0 else "#3E6E56", 0.6)
            text(s, cell, x + 0.25, y + 0.45, cw - 0.6, 0.8, 11.5 if ri else 12,
                 GOLD if ri == 0 or ci == 0 else WHITE, F_BODY_B if ri == 0 or ci == 0 else F_BODY)
        y += 1.6
    footer(s, N[0])

def p_timeline(prs, num, eyebrow, title, sub, items, variant="contour", reason="规划页：等高线暗示路径与延伸"):
    s = page(prs, variant); nav(s, -1)
    head(s, num, eyebrow, title, sub)
    y0 = 11.0
    rect(s, 2.2, y0, 29.2, 0.04, GOLD, 0.85)
    seg = 28.6 / max(len(items) - 1, 1)
    for i, (t, label, note) in enumerate(items):
        x = 2.2 + seg * i
        circle(s, x - 0.22, y0 - 0.2, 0.44, GOLD, 1.0)
        up = i % 2 == 0
        ty = y0 - 3.2 if up else y0 + 0.9
        text(s, t, x - 2.2, ty, 4.4, 0.7, 13, GOLD, F_NUM, align="c")
        text(s, label, x - 2.4, ty + 0.85, 4.8, 0.9, 15, WHITE, F_TITLE, align="c", bold=True)
        text(s, note, x - 2.4, ty + 1.9, 4.8, 1.2, 10.5, MUTED, F_BODY, align="c")
    footer(s, N[0])

def p_team(prs, num, members, variant="plain", reason="团队页：人物为主，背景最简"):
    s = page(prs, variant); nav(s, -1)
    head(s, num, "团队 · TEAM", "六个人，专业互补", "负责人 + 5 名专业成员")
    x = 1.8
    for i, (name, role, exp) in enumerate(members):
        c = circle(s, x, 7.4, 4.2, CARD, 0.8, GOLD, 1.0)
        text_in(c, [name[0]], F_TITLE, GOLD, 24)
        text(s, name, x, 11.9, 4.2, 0.7, 13, WHITE, F_KAI, align="c")
        text(s, role, x, 12.6, 4.2, 0.6, 9.5, GOLD, F_BODY, align="c")
        text(s, exp, x, 13.2, 4.2, 1.0, 9, MUTED, F_BODY, align="c")
        x += 4.9
    ph(s, 1.8, 14.6, 29.2, 2.0, "团队合影", "待补：白色背景半身照各一张")
    footer(s, N[0])

def p_certs(prs, num, items, variant="dark", reason="证书墙：近黑衬托纸质证书，形成反差"):
    s = page(prs, variant); nav(s, -1)
    head(s, num, "知识产权 · IP", "专利与软著：编号待补", "证书扫描件可现场核验")
    x, y = 1.8, 7.4
    for i, label in enumerate(items[:10]):
        row_, cx_ = divmod(i, 5)
        px = 1.8 + cx_ * 6.1; py = 7.4 + row_ * 4.6
        rect(s, px, py, 5.6, 4.1, "#F1ECDC", 0.92, GOLD, 0.8)
        rect(s, px + 0.3, py + 0.5, 5.0, 0.03, "#C9BFA6")
        text(s, label, px + 0.3, py + 1.3, 5.0, 1.6, 10.5, "#20301F", F_BODY_B, align="c")
        text(s, "证书编号 · 待补", px + 0.3, py + 3.4, 5.0, 0.4, 8, "#8A7F68", F_BODY, align="c")
    footer(s, N[0])

def p_end(prs, title, note, variant="contour", reason="封底：呼应开篇，用最重的纹理收尾"):
    s = page(prs, variant); nav(s, -1)
    text(s, title, 2.0, 7.0, 26, 2.6, 44, WHITE, F_TITLE, bold=True)
    rect(s, 2.2, 10.2, 6.0, 0.06, GOLD)
    text(s, note, 2.2, 10.8, 22, 0.9, 14, MUTED, F_KAI)
    footer(s, N[0])


def build(out="deck43.pptx"):
    prs = Presentation(); prs.slide_width, prs.slide_height = Cm(W), Cm(H)
    # 封面沿用成品图
    s = prs.slides.add_slide(prs.slide_layouts[6])
    img = os.path.join(ASSETS, "cover_final.jpg")
    if os.path.exists(img): s.shapes.add_picture(img, 0, 0, Cm(W), Cm(H))
    N[0] = 1
    # 3 页开屏（用封面做暗度递增）
    for frac, dark in ((0.0, 0.25), (0.38, 0.62), (0.66, 0.85)):
        N[0] += 1
        s = page(prs, "dark"); nav(s, -1)
        ph(s, 6.0, 6.2, 22.0, 7.0, "开屏画面", "蒙版展开")
    pages = []
    # 目录
    N[0] += 1; s = page(prs, "dark"); nav(s, -1)
    text(s, "目录", 1.8, 2.2, 6, 1.4, 34, GOLD, F_TITLE, bold=True)
    text(s, "CONTENTS", 1.95, 3.6, 7, 0.7, 10, MUTED, F_EN)
    rect(s, 1.95, 4.3, 3.4, 0.05, GOLD)
    items = [("01","项目背景","政策 · 社会痛点 · 技术痛点"),("02","痛点分析","用工 · 用药 · 农时"),
             ("03","解决方案","三个创新点"),("04","中试验证","三个合作社 · 1200 亩"),
             ("05","商业落地","模式 · 案例 · 推广"),("06","团队规划","团队 · 凭证 · 财务")]
    y = 5.6
    for num, name, note in items:
        rect(s, 1.8, y, 29.2, 1.75, CARD, 0.45, "#3E6E56", 0.5)
        text(s, num, 2.2, y + 0.4, 2.0, 0.9, 17, GOLD, F_NUM)
        text(s, name, 4.4, y + 0.35, 8.0, 0.9, 17, WHITE, F_TITLE, bold=True)
        text(s, note, 13.2, y + 0.55, 15.0, 0.7, 11, MUTED, F_KAI)
        y += 1.9
    footer(s, N[0])

    spec = [
        ("sec", ("01", "项目背景故事", "从政策到痛点，从痛点到技术")),
        ("media", ("01", "项目背景①", "大政策里，我们切的是这一块", "只做水稻植保", "paddies",
                   ["《全国智慧农业行动计划》提出 2028 年植保无人机保有量目标",
                    "东北水稻区连片 1000 亩以上地块占比高，适合飞防",
                    "先做透一个县：辽宁盘锦"]), "auto", True),
        ("media", ("01", "项目背景②", "农村没人打药了", "不是懒，是找不到人", "people",
                   ["60 岁以上劳动力占比高，青壮年外出务工",
                    "一亩地人工打药 40 分钟，一天最多 5 亩",
                    "农忙窗口只有 7–10 天，错过就减药效"]), "auto", False),
        ("rows", ("01", "项目背景③", "从社会痛点推到技术痛点", "痛点一层层往下掉", [
            ("劳动力短缺", "作业队 + 调度平台", "按亩下单，过程可追溯", "team"),
            ("药液浪费", "变量施药模型", "按处方图控流量", "droplet"),
            ("农时抓不准", "多光谱巡检", "先定位再作业", "target")]), "auto"),
        ("cards", ("01", "痛点分析", "三个数字，说明问题", "待测指标，口径待补",
                   [("用工","单亩人工耗时 40 分钟，一天 5 亩上限"),
                    ("用药","雾滴粗，药液流失明显"),
                    ("农时","错过窗口平均减产 8–12%")], ["terminal","droplet","clock"]), "plain"),
        ("media", ("01", "主人公", "这事为什么是我们干", "不是想出来的，是干出来的", "drone",
                   ["2023 年入学，从 STM32 最小系统板起步",
                    "2024 年物联网智能家居，跑通传感器与远程控制",
                    "2026 年 YOLOv11 目标识别，手标 1200 张"]), "auto", True),
        ("cards", ("01", "解决方案", "三个创新点，对应三个痛点", "一一对应，不多不少",
                   [("精准投递","按需施药，不整块均匀喷"),
                    ("雾滴可控","离心雾化，粒径可调"),
                    ("作业可追溯","飞行轨迹与用药量回传")], ["drone","terminal","people"]), "auto"),
        ("kpi4", ("01", "中试验证", "在三个合作社跑通", "田间实测，待补原始记录",
                  [("待补","服务面积","亩","target"),("待补","合作主体","个","team"),
                   ("待补","单机效率","亩/天","gear"),("待补","节药比例","%","droplet")]), "auto"),
        ("sec", ("02", "商业与运营", "怎么赚钱，怎么落地",)),
        ("table", ("02", "产品与服务", "一台设备，三种服务", "定价待核",
                   [["服务类型","内容","定价","毛利"],["飞防作业","按亩收费含药剂","待补","待补"],
                    ["设备销售","整机 + 培训","待补","待补"],["数据服务","虫情 + 处方图","待补","待补"]]), "plain"),
        ("table", ("02", "竞品分析", "不比价格，比综合成本", "参数表待补",
                   [["维度","人工","传统植保机","我们"],["作业效率","3–5 亩/天","待补","待补"],
                    ["雾滴粒径","200μm+","待补","待补"],["综合成本","190 元/亩","待补","待补"]]), "plain"),
        ("cards", ("02", "帮扶案例", "三个村，三种用法", "具体数字待补",
                   [("盘锦合作社","规模经营，看重效率与成本"),
                    ("营口家庭农场","缺劳力，按亩买服务"),
                    ("铁岭村集体","培训飞手，成立作业队")], ["paddies","terminal","people"]), "plain"),
        ("media", ("02", "市场推广", "先做示范田，再谈复制", "渠道清单待补", "paddies",
                   ["示范田 + 合作社口碑",
                    "农服公司渠道合作",
                    "县域作业队加盟"]), "auto", False),
        ("sec", ("03", "组织与团队", "谁在做，凭什么做成",)),
        ("team", None, "plain"),
        ("certs", None, "plain"),
        ("media", ("01", "政策依据", "两份文件，定了方向", "文件号待补", "paddies",
                   ["国家层面：智慧农业行动计划",
                    "省级：农机购置补贴目录",
                    "县级：示范县配套"]), "auto", False),
        ("cards", ("01", "创新点①", "理念：把打药变成精准投递", "重新定义施药",
                   [("按需施药","先巡检出虫害分布，再按图作业"),
                    ("雾滴可控","粒径随作物调整"),
                    ("过程留痕","轨迹与用药量自动回传")], ["drone","droplet","terminal"]), "plain"),
        ("cards", ("01", "创新点②", "技术：三个自研模块", "核心部件国产化",
                   [("自研飞控","抗风 5 级"),
                    ("视觉识别","边缘端识别虫害区"),
                    ("变量喷洒","重喷率下降")], ["gear","target","terminal"]), "auto"),
        ("media", ("01", "创新点③", "工艺：能连续作业 8 小时", "不是能飞就行", "drone",
                   ["药箱快拆，换药不堵",
                    "5 级风下航线偏移可控",
                    "整机维护成本待核"]), "auto", True),
        ("table", ("02", "协议与凭证", "已签与已收", "金额待补",
                   [["凭证","数量","金额","状态"],["服务合同","待补","待补","待补"],
                    ["意向书","待补","待补","待补"],["发票","待补","待补","待补"]]), "plain"),
        ("kpi4", ("03", "财务", "五年营收路径", "数字待补",
                  [("待补","2026","万元","gear"),("待补","2028","万元","target"),
                   ("待补","2030","万元","team"),("待补","净利率","%","droplet")]), "auto"),
        ("cards", ("03", "专家顾问", "三位顾问，各管一段", "不是挂名",
                   [("技术顾问","高校植保方向"),
                    ("工程顾问","农机企业总工"),
                    ("商业顾问","农服公司创始人")], ["people","gear","target"]), "plain"),
        ("cards", ("04", "教育维度①", "教育：项目变成了课", "从实验室走进课堂",
                   [("案例进课堂","嵌入式课程使用"),
                    ("教改立项","省级 1 项"),
                    ("教学成果","校级奖项")], ["people","terminal","target"]), "plain"),
        ("cards", ("04", "教育维度②", "科技：成果往一线走", "不是发完论文就结束",
                   [("专利转化","作价入股"),
                    ("技术推广","向农服公司输出标准"),
                    ("产学研","共建试验基地")], ["gear","paddies","team"]), "auto"),
        ("media", ("04", "教育维度③", "人才：六个人怎么长出来的", "项目是育人载体", "people",
                   ["3 人保研，继续农机方向",
                    "2 人创业，成立农服公司",
                    "1 人入职农机企业"]), "auto", False),
        ("kpi4", ("04", "媒体报道", "报道与反馈", "链接待补",
                  [("待补","媒体报道","篇","target"),("待补","感谢信","封","team"),
                   ("待补","现场观摩","次","gear"),("待补","示范点","个","droplet")]), "plain"),
        ("timeline", ("04", "近期计划", "未来 12 个月", "节点可调",
                      [("Q1","示范田","3 个点位"),("Q2","作业季","完成首次服务"),
                       ("Q3","数据沉淀","形成处方图库"),("Q4","复制","进入第 2 个县")]), "auto"),
        ("cards", ("01", "项目概述", "一页看懂：我们做什么", "200 字说不清，就不是好项目",
                   [("做什么","给合作社与种粮大户提供飞防服务"),
                    ("怎么赚","按亩收费 + 数据服务"),
                    ("进展","三个示范点，营收待补")], ["drone","terminal","paddies"]), "plain"),
        ("media", ("02", "服务模式", "上游采购，中游作业，下游按亩付费", "不做重资产", "terminal",
                   ["上游：整机与药剂集采",
                    "中游：作业队 + 调度平台",
                    "下游：合作社按效果付费"]), "auto", True),
        ("media", ("03", "风险应对", "三个风险，三条对策", "风险清单待补", "gear",
                   ["天气窗口：多区域调度",
                    "设备故障：备机与本地维修",
                    "政策变动：跟进补贴目录"]), "auto", False),
        ("cards", ("04", "社会反馈", "地方的三种反馈", "原件待补",
                   [("合作社感谢信","解决农忙用工"),
                    ("村集体锦旗","培训飞手"),
                    ("县级示范认定","列入示范项目")], ["people","team","target"]), "auto"),
        ("sec", ("04", "社会成效与规划", "不只是生意，也是答案",)),
        ("kpi4", ("04", "社会效益", "四个指标，数据待补", "口径与服务期一致",
                  [("待补","减药比例","%","droplet"),("待补","带动就业","人","team"),
                   ("待补","累计作业","亩次","gear"),("待补","带动增收","万元","target")]), "auto"),
        ("timeline", ("04", "发展规划", "三步走，先把手头做透", "时间节点可调",
                      [("2026","做透一个县","3 个示范点"),("2027","省内复制","10 个县"),
                       ("2028","跨省扩张","东北三省"),("2029","标准化","输出作业标准"),
                       ("2030","出海","东南亚水稻区")]), "auto"),
    ]
    for kind, args, *rest in spec:
        N[0] += 1
        if kind == "sec": p_section(prs, *args)
        elif kind == "media": p_media(prs, *args)
        elif kind == "cards": p_cards3(prs, *args)
        elif kind == "rows": p_rows3(prs, *args)
        elif kind == "kpi4": p_kpi4(prs, *args)
        elif kind == "table": p_table(prs, *args)
        elif kind == "timeline": p_timeline(prs, *args)
        elif kind == "team": p_team(prs, N[0], [("李明","负责人","嵌入式 3 年"),("王强","硬件","PCB/飞控"),
                                   ("张丽","算法","视觉识别"),("赵磊","作业队长","持证飞手"),
                                   ("陈静","运营","合作社对接"),("刘洋","财务","成本与融资")])
        elif kind == "certs": p_certs(prs, N[0], ["发明专利","实用新型","软著①","软著②","软著③",
                                   "查新报告","检测报告","试验报告","合作协议①","合作协议②"])
    N[0] += 1; p_end(prs, "谢谢", "慧飞植保 · 让每一亩地都用上无人机")
    prs.save(out)
    print(f"已生成 {out}（{len(prs.slides.__iter__.__self__._sldIdLst)} 页）")

if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(); ap.add_argument("--out", default="deck43.pptx")
    build(ap.parse_args().out)
