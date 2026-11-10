#!/usr/bin/env python3
"""gen_assets.py — 生成 PPT 需要的素材，不要截图拼凑。

四种素材，全部按主题色生成、透明背景、2 倍分辨率，可直接放进模板图片框：

    # 1. 图表（bar / hbar / line / donut / area）
    python3 gen_assets.py chart --kind bar --data data.json --out chart.png --theme tech

    # 2. 流程（水平流程条）
    python3 gen_assets.py flow --steps "数据采集,模型训练,落地部署" --out flow.png --theme tech

    # 3. 时间轴
    python3 gen_assets.py timeline --items "2024|立项,2025|试点,2026|推广" --out tl.png --theme tech

    # 4. 背景（封面/章节页底图）
    python3 gen_assets.py bg --style tech --out bg.png --width 2400 --height 1350

data.json 格式:
    {"categories": ["Q1","Q2","Q3","Q4"],
     "series": [{"name": "营收", "values": [12, 18, 25, 31]}]}

主题: tech(深蓝科技) / red(红金) / rural(绿意) / clean(浅色极简)
"""
import argparse
import json
import math
import os
import sys

from PIL import Image, ImageDraw, ImageFont

CJK = "/system/fonts/NotoSansCJK-Regular.ttc"
CJK_B = "/system/fonts/NotoSansCJK-Bold.ttc"

THEMES = {
    "tech":  {"bg": "#0A2540", "main": "#2F80FF", "accent": "#00D4FF", "text": "#E8EEF9",
              "muted": "#7A8DA8", "light": "#16385C"},
    "red":   {"bg": "#2A0F0F", "main": "#C8102E", "accent": "#D4AF37", "text": "#F6EEE6",
              "muted": "#B79A94", "light": "#43201D"},
    "rural": {"bg": "#10240F", "main": "#2E7D32", "accent": "#F2B705", "text": "#EEF6EA",
              "muted": "#93AC8C", "light": "#1E3D1C"},
    "clean": {"bg": "#FFFFFF", "main": "#0B5FFF", "accent": "#00A3FF", "text": "#1F2430",
              "muted": "#8A93A6", "light": "#EEF3FF"},
}


def hex2rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def font(size, bold=False):
    p = CJK_B if bold and os.path.exists(CJK_B) else CJK
    return ImageFont.truetype(p, size)


# ---------- 图表 ----------
def make_chart(kind, data, out, theme_name, title="", width=1600, transparent=True):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib import font_manager
    except ImportError:
        sys.exit("需要 matplotlib: pip install matplotlib")
    th = THEMES[theme_name]
    # 中文字体
    if os.path.exists(CJK):
        font_manager.fontManager.addfont(CJK)
        plt.rcParams["font.family"] = font_manager.FontProperties(fname=CJK).get_name()
    plt.rcParams["axes.unicode_minus"] = False

    cats = data.get("categories", [])
    series = data.get("series", [])
    dpi = 200
    figsize = (width / dpi, width / dpi * 0.5)
    fig, ax = plt.subplots(figsize=figsize, dpi=dpi)
    if transparent:
        fig.patch.set_alpha(0)
        ax.set_facecolor("none")
    else:
        fig.patch.set_facecolor(th["bg"])
        ax.set_facecolor(th["bg"])

    colors = [th["main"], th["accent"], "#7CD4FF", "#FFD166", "#EF476F", "#06D6A0"]
    tc = th["text"] if not transparent or theme_name != "clean" else th["text"]

    if kind in ("bar", "hbar", "line", "area"):
        n = len(series) or 1
        idx = range(len(cats))
        bar_w = 0.8 / n
        for i, s in enumerate(series):
            vals = s.get("values", [])
            color = colors[i % len(colors)]
            if kind == "bar":
                ax.bar([x + i * bar_w for x in idx], vals, width=bar_w * 0.86,
                       color=color, label=s.get("name", ""), zorder=3)
            elif kind == "hbar":
                ax.barh([x + i * bar_w for x in idx], vals, height=bar_w * 0.8,
                        color=color, label=s.get("name", ""), zorder=3)
            elif kind == "line":
                ax.plot(list(idx), vals, color=color, linewidth=3.2,
                        marker="o", markersize=8, label=s.get("name", ""), zorder=3)
            else:
                ax.fill_between(list(idx), vals, color=color, alpha=0.35)
                ax.plot(list(idx), vals, color=color, linewidth=3, label=s.get("name", ""))
        if kind == "hbar":
            ax.set_yticks(list(idx))
            ax.set_yticklabels(cats, color=tc, fontsize=16)
            ax.set_xticks([])
            ax.invert_yaxis()
        else:
            ax.set_xticks(list(idx))
            ax.set_xticklabels(cats, color=tc, fontsize=16)
        for sp in ("top", "right", "left", "bottom"):
            ax.spines[sp].set_visible(False)
        ax.tick_params(colors=tc, labelsize=16, length=0)
        ax.grid(axis="y", color=th["muted"], alpha=0.25, linewidth=0.8, zorder=0)
        ax.set_axisbelow(True)
        # 数值标签（hbar 标在条形末端右侧，其余标在顶部）
        for s in series:
            for x, v in enumerate(s.get("values", [])):
                if kind == "hbar":
                    ax.annotate(f"{v}", (v, x), textcoords="offset points", xytext=(8, 0),
                                ha="left", va="center", color=tc, fontsize=15, fontweight="bold")
                else:
                    ax.annotate(f"{v}", (x, v), textcoords="offset points", xytext=(0, 8),
                                ha="center", color=tc, fontsize=15, fontweight="bold")
    elif kind == "donut":
        vals = (series[0].get("values") if series else []) or []
        labels = cats or [f"{i+1}" for i in range(len(vals))]
        wedges, _ = ax.pie(vals, colors=colors[:len(vals)], startangle=90,
                           wedgeprops={"width": 0.38, "edgecolor": "none"})
        ax.legend(wedges, labels, loc="center left", bbox_to_anchor=(1.02, 0.5),
                  frameon=False, labelcolor=tc, fontsize=13)

    if title:
        ax.set_title(title, color=tc, fontsize=22, fontweight="bold", pad=16)
    if len(series) > 1 and kind != "donut":
        ax.legend(frameon=False, labelcolor=tc, loc="upper left", fontsize=12,
                  bbox_to_anchor=(0, 1.12), ncol=len(series))
    fig.tight_layout()
    fig.savefig(out, transparent=transparent)
    plt.close(fig)
    return out


# ---------- 流程 ----------
def make_flow(steps, out, theme_name, width=1800, transparent=True):
    th = THEMES[theme_name]
    n = len(steps)
    h = 360
    img = Image.new("RGBA", (width, h), (0, 0, 0, 0) if transparent else hex2rgb(th["bg"]) + (255,))
    d = ImageDraw.Draw(img)
    pad = 40
    gap = 46
    box_w = (width - pad * 2 - gap * (n - 1)) / n
    box_h = 150
    y0 = (h - box_h) / 2
    f_num = font(26, True)
    f_txt = font(22)
    for i, s in enumerate(steps):
        x = pad + i * (box_w + gap)
        d.rounded_rectangle([x, y0, x + box_w, y0 + box_h], radius=18,
                            fill=hex2rgb(th["light"]) + (235,),
                            outline=hex2rgb(th["main"]) + (255,), width=3)
        # 编号圆
        r = 26
        cx, cy = x + 40, y0 + 40
        d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=hex2rgb(th["accent"]) + (255,))
        num = str(i + 1)
        tw = d.textlength(num, font=f_num)
        d.text((cx - tw / 2, cy - 16), num, font=f_num, fill=hex2rgb(th["bg"]))
        # 文字
        lines = s.split("|")
        tx = x + 22
        ty = y0 + 74
        d.text((tx, ty), lines[0], font=f_txt, fill=hex2rgb(th["text"]))
        if len(lines) > 1:
            d.text((tx, ty + 34), lines[1], font=font(16), fill=hex2rgb(th["muted"]))
        # 箭头
        if i < n - 1:
            ax0 = x + box_w + 8
            ax1 = x + box_w + gap - 8
            ay = y0 + box_h / 2
            d.line([ax0, ay, ax1 - 8, ay], fill=hex2rgb(th["accent"]), width=4)
            d.polygon([(ax1, ay), (ax1 - 14, ay - 9), (ax1 - 14, ay + 9)],
                      fill=hex2rgb(th["accent"]))
    img.save(out)
    return out


# ---------- 时间轴 ----------
def make_timeline(items, out, theme_name, width=2000, height=420, transparent=True):
    th = THEMES[theme_name]
    img = Image.new("RGBA", (width, height), (0, 0, 0, 0) if transparent else hex2rgb(th["bg"]) + (255,))
    d = ImageDraw.Draw(img)
    n = len(items)
    y_axis = height // 2
    d.line([80, y_axis, width - 80, y_axis], fill=hex2rgb(th["main"]) + (255,), width=4)
    seg = (width - 160) / max(n - 1, 1)
    f_l = font(24, True)
    f_t = font(24, True)
    f_n = font(17)
    for i, it in enumerate(items):
        parts = it.split("|")
        label = parts[0]
        title = parts[1] if len(parts) > 1 else ""
        note = parts[2] if len(parts) > 2 else ""
        cx = 80 + seg * i
        up = i % 2 == 0
        # 节点
        d.ellipse([cx - 14, y_axis - 14, cx + 14, y_axis + 14],
                  fill=hex2rgb(th["accent"]) + (255,), outline=hex2rgb(th["bg"]) + (255,), width=4)
        # 竖线
        yv0 = y_axis - 60 if up else y_axis + 60
        d.line([cx, y_axis, cx, yv0], fill=hex2rgb(th["main"]) + (200,), width=3)
        ty = yv0 - 110 if up else yv0 + 8
        tw = d.textlength(label, font=f_l)
        d.text((cx - tw / 2, ty), label, font=f_l, fill=hex2rgb(th["accent"]))
        tw = d.textlength(title, font=f_t)
        d.text((cx - tw / 2, ty + 38), title, font=f_t, fill=hex2rgb(th["text"]))
        if note:
            tw = d.textlength(note, font=f_n)
            d.text((cx - tw / 2, ty + 76), note, font=f_n, fill=hex2rgb(th["muted"]))
    img.save(out)
    return out


# ---------- 背景 ----------
def make_bg(style, out, width=2400, height=1350):
    th = THEMES.get(style, THEMES["tech"])
    img = Image.new("RGB", (width, height), hex2rgb(th["bg"]))
    d = ImageDraw.Draw(img)
    # 垂直渐变
    c1, c2 = hex2rgb(th["bg"]), hex2rgb(th["light"])
    for y in range(height):
        t = y / height
        c = tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3))
        d.line([(0, y), (width, y)], fill=c)
    # 细网格
    grid = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    gd = ImageDraw.Draw(grid)
    step = 80
    for x in range(0, width, step):
        gd.line([(x, 0), (x, height)], fill=hex2rgb(th["main"]) + (26,), width=1)
    for y in range(0, height, step):
        gd.line([(0, y), (width, y)], fill=hex2rgb(th["main"]) + (26,), width=1)
    img = Image.alpha_composite(img.convert("RGBA"), grid)
    # 光晕
    glow = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    gdraw = ImageDraw.Draw(glow)
    for (gx, gy, gr) in [(width * 0.78, height * 0.22, 420),
                         (width * 0.18, height * 0.82, 320),
                         (width * 0.55, height * 0.10, 240)]:
        for k in range(24, 0, -1):
            alpha = int(9 * (1 - k / 24))
            rr = gr * k / 24
            gdraw.ellipse([gx - rr, gy - rr, gx + rr, gy + rr],
                          fill=hex2rgb(th["accent"]) + (alpha,))
    img = Image.alpha_composite(img, glow)
    # 细斜线装饰
    lines = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    ld = ImageDraw.Draw(lines)
    for i in range(-height, width, 140):
        ld.line([(i, height), (i + height, 0)], fill=hex2rgb(th["accent"]) + (14,), width=2)
    img = Image.alpha_composite(img, lines)
    img.convert("RGB").save(out, quality=92)
    return out


def main():
    ap = argparse.ArgumentParser(description="生成 PPT 素材（图表/流程/时间轴/背景）")
    sub = ap.add_subparsers(dest="cmd", required=True)

    c = sub.add_parser("chart")
    c.add_argument("--kind", default="bar", choices=["bar", "hbar", "line", "donut", "area"])
    c.add_argument("--data", required=True)
    c.add_argument("--out", required=True)
    c.add_argument("--theme", default="tech", choices=list(THEMES))
    c.add_argument("--title", default="")
    c.add_argument("--width", type=int, default=1600)
    c.add_argument("--no-transparent", action="store_true")

    f = sub.add_parser("flow")
    f.add_argument("--steps", required=True, help="用逗号分隔；可用 | 写副标题：数据采集|传感器")
    f.add_argument("--out", required=True)
    f.add_argument("--theme", default="tech", choices=list(THEMES))
    f.add_argument("--width", type=int, default=1800)

    t = sub.add_parser("timeline")
    t.add_argument("--items", required=True, help="逗号分隔：2024|立项|第一代样机,2025|试点")
    t.add_argument("--out", required=True)
    t.add_argument("--theme", default="tech", choices=list(THEMES))

    b = sub.add_parser("bg")
    b.add_argument("--style", default="tech", choices=list(THEMES))
    b.add_argument("--out", required=True)
    b.add_argument("--width", type=int, default=2400)
    b.add_argument("--height", type=int, default=1350)

    a = ap.parse_args()
    if a.cmd == "chart":
        data = json.loads(a.data if a.data.strip().startswith("{") else open(a.data, encoding="utf-8").read())
        p = make_chart(a.kind, data, a.out, a.theme, a.title, a.width, not a.no_transparent)
    elif a.cmd == "flow":
        p = make_flow(a.steps.split(","), a.out, a.theme, a.width)
    elif a.cmd == "timeline":
        p = make_timeline(a.items.split(","), a.out, a.theme)
    else:
        p = make_bg(a.style, a.out, a.width, a.height)
    print(f"已生成 {p}（{os.path.getsize(p)/1024:.0f} KB）")


if __name__ == "__main__":
    main()
