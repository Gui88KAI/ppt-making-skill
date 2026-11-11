#!/usr/bin/env python3
"""cover_example.py — 一键做创赛封面（4 页：3 页开屏 + 1 页正式封面）。

用法:
    python3 examples/cover_example.py \
      --title 慧飞植保 \
      --subtitle "让每一亩地都用上无人机" \
      --tagline "飞防服务 + 数据驱动的精准施药" \
      --competition "第十届中国国际“互联网+”大学生创新创业大赛" \
      --info "赛道 乡村振兴 · 组别 创意组 · 负责人 XXX · 大连民族大学" \
      --bg assets/bg.jpg --hero assets/hero.jpg \
      --palette green --out 封面.pptx

依赖: pip install python-pptx pillow
素材: --bg 背景场景图（gen_image.py --kind background）
      --hero 人物/产品图（gen_image.py --kind hero --style illustration）
"""
import argparse
import os
import sys

from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont
from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Cm

W_CM, H_CM = 33.867, 19.05
PX_W, PX_H = 1920, 1080

# 类型 → 封面配置（左上角文字 / 配色 / 字体风格）
TYPE_PRESETS = {
    "01": dict(eyebrow="第十届中国国际“互联网+”大学生创新创业大赛", palette="green", style="culture"),
    "02": dict(eyebrow="商业计划书 · BUSINESS PLAN 2026", palette="blue", style="business"),
    "03": dict(eyebrow="天使轮路演 · ANGEL ROUND 2026", palette="blue", style="tech"),
    "04": dict(eyebrow="本科毕业论文答辩 · 2026 届", palette="blue", style="business"),
    "05": dict(eyebrow="《课程名》课程期末汇报", palette="green", style="culture"),
}

PALETTES = {
    "green": dict(dark=(7, 19, 14), gold=(212, 162, 76), gold2=(216, 185, 108),
                  text=(242, 239, 230), muted=(157, 179, 165), stroke=(20, 40, 28)),
    "blue": dict(dark=(6, 21, 48), gold=(232, 244, 255), gold2=(160, 220, 255),
                 text=(232, 241, 255), muted=(143, 166, 194), stroke=(8, 30, 60)),
    "red": dict(dark=(42, 11, 11), gold=(217, 164, 65), gold2=(240, 200, 120),
                text=(248, 238, 228), muted=(194, 160, 143), stroke=(60, 20, 18)),
}

CJK = "/system/fonts/NotoSansCJK-Regular.ttc"
FT = "/usr/share/fonts/truetype/custom/"

# 四种字体方案（统计自素材库 40 套模板，见 references/字体方案.md）
FONT_STACKS = {
    "culture": dict(sub=FT + "华文楷体.TTF",
                    body=FT + "思源宋体 CN Light.otf",
                    info=FT + "阿里巴巴普惠体 L.ttf"),
    "tech": dict(sub=FT + "汉仪雅酷黑 75W.ttf",
                 body=FT + "阿里巴巴普惠体 L.ttf",
                 info=FT + "阿里巴巴普惠体 R.ttf"),
    "business": dict(sub=FT + "阿里巴巴普惠体 B.ttf",
                     body=FT + "阿里巴巴普惠体 L.ttf",
                     info=FT + "阿里巴巴普惠体 R.ttf"),
    "royal": dict(sub=FT + "思源宋体 CN Heavy.otf",
                  body=FT + "思源宋体 CN Light.otf",
                  info=FT + "阿里巴巴普惠体 L.ttf"),
}
FALLBACK_FONTS = [
    "/usr/share/fonts/truetype/custom/演示流云楷--淘宝店：烦烦优创馆.ttf",
    "/usr/share/fonts/truetype/custom/演示流云楷.ttf",
    "/usr/share/fonts/truetype/custom/汉仪许静行楷W.ttf",
]


def find_calligraphy():
    for f in FALLBACK_FONTS:
        if os.path.exists(f):
            return f
    return None


def render_char(ch, out_path, pal, size=900):
    """把一个字渲染成透明 PNG（金色渐变 + 深色描边 + 阴影）。"""
    font_path = find_calligraphy()
    if font_path is None:
        print("⚠ 没找到书法字体，请把 演示流云楷.ttf 放到 templates/fonts/ 并安装")
        sys.exit(1)
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    f = ImageFont.truetype(font_path, size)
    bbox = d.textbbox((0, 0), ch, font=f)
    x = (size - (bbox[2] - bbox[0])) // 2 - bbox[0]
    y = (size - (bbox[3] - bbox[1])) // 2 - bbox[1]
    # 描边
    for dx in range(-6, 7, 3):
        for dy in range(-6, 7, 3):
            d.text((x + dx, y + dy), ch, font=f, fill=pal["stroke"] + (255,))
    # 金色渐变
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).text((x, y), ch, font=f, fill=255)
    grad = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    gd = ImageDraw.Draw(grad)
    for yy in range(size):
        t = yy / size
        c = tuple(int(pal["gold"][i] + (pal["gold2"][i] - pal["gold"][i]) * t) for i in range(3))
        gd.line([(0, yy), (size, yy)], fill=c + (255,))
    grad.putalpha(mask)
    # 阴影
    shadow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    ImageDraw.Draw(shadow).text((x + 10, y + 10), ch, font=f, fill=(0, 0, 0, 120))
    shadow = shadow.filter(ImageFilter.GaussianBlur(8))
    out = Image.alpha_composite(shadow, Image.alpha_composite(img, grad))
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    out.save(out_path)


def compose_cover(title, subtitle, tagline, competition, info, bg_path, hero_path, pal, out_path, style="culture"):
    fs = FONT_STACKS.get(style, FONT_STACKS["culture"])
    bg = Image.open(bg_path).convert("RGB").resize((PX_W, PX_H), Image.LANCZOS)
    bg = ImageEnhance.Brightness(bg).enhance(0.55)
    left = Image.new("RGBA", (PX_W, PX_H), (0, 0, 0, 0))
    ld = ImageDraw.Draw(left)
    for x in range(PX_W):
        t = max(0, 1 - x / (PX_W * 0.62))
        ld.line([(x, 0), (x, PX_H)], fill=pal["dark"] + (int(150 * t),))
    bg = Image.alpha_composite(bg.convert("RGBA"), left)

    if hero_path and os.path.exists(hero_path):
        hero = Image.open(hero_path).convert("RGBA")
        m = Image.new("L", hero.size, 0)
        ImageDraw.Draw(m).ellipse([hero.width * .05, hero.height * .02,
                                   hero.width * .95, hero.height * 1.02], fill=255)
        m = m.filter(ImageFilter.GaussianBlur(35))
        hero.putalpha(m)
        th = int(PX_H * 0.95)
        r = th / hero.height
        hero = hero.resize((int(hero.width * r), th), Image.LANCZOS)
        bg.alpha_composite(hero, (PX_W - int(hero.width * 0.88), int(PX_H * 0.10)))

    chars = list(title)
    n = max(len(chars), 1)
    x0, step = 230, min(290, int(1180 / n))
    for i, ch in enumerate(chars):
        p = f"chars/{i}_{ch}.png"
        render_char(ch, p, pal)
        im = Image.open(p).convert("RGBA").resize((400 - i * 18, 400 - i * 18), Image.LANCZOS)
        bg.alpha_composite(im, (x0 + i * step, 175 - i * 33))

    def F(path, s):
        try:
            return ImageFont.truetype(path, s)
        except Exception:
            try:
                return ImageFont.truetype(CJK, s)
            except Exception:
                return ImageFont.load_default()

    d = ImageDraw.Draw(bg)
    d.text((100, 62), competition, font=F(fs["info"], 30), fill=pal["gold"])
    d.rectangle([232, 690, 610, 698], fill=pal["gold"])
    d.text((228, 720), subtitle, font=F(fs["sub"], 52), fill=pal["text"])
    d.text((230, 800), tagline, font=F(fs["body"], 30), fill=pal["muted"])
    bar = Image.new("RGBA", (PX_W, PX_H), (0, 0, 0, 0))
    ImageDraw.Draw(bar).rounded_rectangle([228, 872, 1420, 946], radius=12,
                                          fill=pal["dark"] + (170,),
                                          outline=pal["gold"] + (140,), width=2)
    bg = Image.alpha_composite(bg, bar)
    ImageDraw.Draw(bg).text((254, 890), info, font=F(fs["info"], 26), fill=pal["text"])
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    bg.convert("RGB").save(out_path, quality=92)


def build_pptx(cover_path, out_path):
    def transition(slide, kind="fade"):
        old = slide._element.find(qn("p:transition"))
        if old is not None:
            slide._element.remove(old)
        t = etree.SubElement(slide._element, qn("p:transition"))
        t.set("spd", "slow")
        if kind == "morph":
            etree.SubElement(t, qn("p:morph"), {"option": "byObject"})
        else:
            etree.SubElement(t, qn("p:fade"))

    full = Image.open(cover_path)
    w, h = full.size
    dark = full.point(lambda p: int(p * 0.25))
    os.makedirs("opening", exist_ok=True)
    frames = []
    for i, frac in enumerate([1.0, 0.62, 0.34], 1):
        canvas = dark.copy()
        top = int(h * (1 - frac))
        canvas.paste(full.crop((0, top, w, h)), (0, top))
        p = f"opening/step{i}.jpg"
        canvas.save(p, quality=90)
        frames.append(p)
    frames.append(cover_path)

    prs = Presentation()
    prs.slide_width = Cm(W_CM)
    prs.slide_height = Cm(H_CM)
    for i, f in enumerate(frames):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        s.shapes.add_picture(f, 0, 0, Cm(W_CM), Cm(H_CM))
        transition(s, "fade" if i < 3 else "morph")
    prs.save(out_path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--title", required=True)
    ap.add_argument("--subtitle", default="")
    ap.add_argument("--tagline", default="")
    ap.add_argument("--competition", default="第十届中国国际“互联网+”大学生创新创业大赛")
    ap.add_argument("--info", default="赛道 乡村振兴 · 组别 创意组 · 负责人 XXX")
    ap.add_argument("--bg", required=True)
    ap.add_argument("--hero", default="")
    ap.add_argument("--palette", default=None, choices=list(PALETTES))
    ap.add_argument("--style", default="", choices=[""] + list(FONT_STACKS))
    ap.add_argument("--type", default="", choices=[""] + list(TYPE_PRESETS),
                    help="01创赛 02BP 03路演 04答辩 05课程；会自动设置左上角文字/配色/字体")
    ap.add_argument("--out", default="封面.pptx")
    a = ap.parse_args()
    # 类型预设优先（用户显式传的 --style/--palette 仍可覆盖）
    preset = TYPE_PRESETS.get(a.type, {})
    if preset:
        a.competition = preset["eyebrow"]
        if a.palette is None:
            a.palette = preset["palette"]
        if not a.style:
            a.style = preset["style"]
    if a.palette is None:
        a.palette = "green"
    if not a.style:
        a.style = "culture"
    pal = PALETTES[a.palette]
    compose_cover(a.title, a.subtitle, a.tagline, a.competition, a.info,
                  a.bg, a.hero, pal, "assets/cover_composed.jpg", style=a.style)
    build_pptx("assets/cover_composed.jpg", a.out)
    print(f"已生成 {a.out}（4 页：3 页开屏 + 正式封面）")


if __name__ == "__main__":
    main()
