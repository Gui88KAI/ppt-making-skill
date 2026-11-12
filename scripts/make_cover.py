#!/usr/bin/env python3
"""make_cover.py — 一条命令产出同款创赛封面（4 页：3 页蒙版开屏 + 正式封面）。

用法:
  python3 scripts/make_cover.py \
    --title 慧飞植保 --subtitle "让每一亩地都用上无人机" \
    --tagline "飞防服务 + 数据驱动的精准施药" \
    --info "赛道 乡村振兴 · 组别 创意组 · 负责人 XXX · 大连民族大学" \
    --competition "第十届中国国际“互联网+”大学生创新创业大赛" \
    --bg assets/bg.jpg --hero assets/hero.png \
    --palette green --out 封面.pptx

依赖: pip install python-pptx pillow
字体: templates/fonts/演示流云楷.ttf（或系统里同名字体），缺失时会提示
"""
import argparse, json, os, sys
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont
from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Cm

W_CM, H_CM = 33.867, 19.05
PW, PH = 1920, 1080
CJK = "/system/fonts/NotoSansCJK-Regular.ttc"

PALETTES = {
    "green": dict(dark=(7, 19, 14), gold=(212, 162, 76), gold2=(216, 185, 108),
                  text=(242, 239, 230), muted=(157, 179, 165), stroke=(20, 40, 28)),
    "blue":  dict(dark=(6, 21, 48), gold=(232, 244, 255), gold2=(150, 210, 250),
                  text=(232, 241, 255), muted=(143, 166, 194), stroke=(8, 30, 60)),
    "red":   dict(dark=(42, 11, 11), gold=(217, 164, 65), gold2=(240, 200, 120),
                  text=(248, 238, 228), muted=(194, 160, 143), stroke=(60, 20, 18)),
}
# 参数指纹：改这里 = 改所有封面
SPEC = dict(
    char_size=[400, 378, 356, 336],      # 每字像素边长
    char_x0=230, char_gap=250,           # 起点与字距
    char_base_y=215,                     # 基线 y
    char_offsets=[12, -22, 18, -12],     # 上下错落（关键：不单调上爬）
    stroke_px=6, shadow_dx=9, shadow_dy=9, shadow_blur=7, shadow_alpha=130,
    grad_top=(212, 160, 74), grad_bottom=(188, 194, 118),
    bg_darken=0.72, bg_satur=1.15, bg_contrast=1.10,
    hero_bright=0.88, hero_feather=55, hero_match=0.78,
    ribbon_alpha=0.0,
    eyebrow=(250, 86, 30), subtitle=(248, 742, 56), tagline=(252, 846, 30),
    info_bar=(250, 912, 1400, 986), info_size=26,
    open_steps=[0.0, 0.38, 0.66],
)

FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/custom/演示流云楷--淘宝店：烦烦优创馆.ttf",
    "/usr/share/fonts/truetype/custom/演示流云楷.ttf",
    "/usr/share/fonts/truetype/custom/汉仪许静行楷W.ttf",
]
FONT_BODY = ["/usr/share/fonts/truetype/custom/阿里巴巴普惠体 L.ttf", CJK]
FONT_SUB = ["/usr/share/fonts/truetype/custom/思源宋体 CN Heavy.otf", CJK]


def pick(cands, size):
    for p in cands:
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    return ImageFont.truetype(CJK, size)


def hexrgb(s):
    s = s.lstrip("#")
    return tuple(int(s[i:i + 2], 16) for i in (0, 2, 4))


def render_char(ch, path, pal, size):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    f = pick(FONT_CANDIDATES, size)
    b = d.textbbox((0, 0), ch, font=f)
    x = (size - (b[2] - b[0])) // 2 - b[0]
    y = (size - (b[3] - b[1])) // 2 - b[1]
    for dx in range(-SPEC["stroke_px"], SPEC["stroke_px"] + 1, 3):
        for dy in range(-SPEC["stroke_px"], SPEC["stroke_px"] + 1, 3):
            d.text((x + dx, y + dy), ch, font=f, fill=pal["stroke"] + (255,))
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).text((x, y), ch, font=f, fill=255)
    grad = Image.new("RGBA", (size, size))
    gd = ImageDraw.Draw(grad)
    c0, c1 = SPEC["grad_top"], SPEC["grad_bottom"]
    for yy in range(size):
        t = yy / size
        gd.line([(0, yy), (size, yy)],
                fill=tuple(int(c0[i] + (c1[i] - c0[i]) * t) for i in range(3)) + (255,))
    grad.putalpha(mask)
    sh = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    ImageDraw.Draw(sh).text((x + SPEC["shadow_dx"], y + SPEC["shadow_dy"]), ch, font=f,
                            fill=(0, 0, 0, SPEC["shadow_alpha"]))
    sh = sh.filter(ImageFilter.GaussianBlur(SPEC["shadow_blur"]))
    out = Image.alpha_composite(sh, Image.alpha_composite(img, grad))
    out.save(path)
    return out


def color_match(hero, ref, strength):
    import numpy as np
    h = np.asarray(hero.convert("RGB")).astype("float32")
    r = np.asarray(ref.convert("RGB")).astype("float32")
    for c in range(3):
        hs, rs = h[:, :, c].std(), r[:, :, c].std()
        if hs < 1e-3:
            continue
        h[:, :, c] = (h[:, :, c] - h[:, :, c].mean()) / hs * rs + r[:, :, c].mean()
    out = Image.fromarray(np.clip(h, 0, 255).astype("uint8"))
    return Image.blend(hero.convert("RGB"), out, strength)


def radial_alpha(img_size):
    w, h = img_size
    m = Image.new("L", (w, h), 0)
    d = ImageDraw.Draw(m)
    steps = 60
    for i in range(steps, 0, -1):
        t = i / steps
        d.ellipse([w * .5 - w * .52 * t, h * .52 - h * .56 * t,
                   w * .5 + w * .52 * t, h * .52 + h * .56 * t],
                  fill=int(255 * (1 - t) ** 0.55))
    return m.filter(ImageFilter.GaussianBlur(30))


def compose(args, pal):
    bg = Image.open(args.bg).convert("RGB").resize((PW, PH), Image.LANCZOS)
    bg = ImageEnhance.Brightness(bg).enhance(SPEC["bg_darken"])
    bg = ImageEnhance.Color(bg).enhance(SPEC["bg_satur"])
    bg = ImageEnhance.Contrast(bg).enhance(SPEC["bg_contrast"])
    shade = Image.new("RGBA", (PW, PH), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shade)
    for x in range(PW):
        t = max(0, 1 - x / (PW * 0.66))
        sd.line([(x, 0), (x, PH)], fill=pal["dark"] + (int(190 * t),))
    bg = Image.alpha_composite(bg.convert("RGBA"), shade)

    # 人物：径向羽化 + 色彩匹配 + 边缘光 + 前景遮挡
    if args.hero and os.path.exists(args.hero):
        hero = Image.open(args.hero).convert("RGBA")
        hh = int(PH * 0.92)
        r = hh / hero.height
        hero = hero.resize((int(hero.width * r), hh), Image.LANCZOS)
        a0 = hero.split()[3]
        rad = radial_alpha(hero.size)
        import numpy as np
        a = Image.fromarray(np.minimum(np.asarray(a0, "uint16"), np.asarray(rad, "uint16")).astype("uint8"))
        a = a.filter(ImageFilter.GaussianBlur(SPEC["hero_feather"]))
        hero = ImageEnhance.Brightness(color_match(hero, bg, SPEC["hero_match"])).enhance(SPEC["hero_bright"])
        hero = hero.convert("RGBA"); hero.putalpha(a)
        hx_ = PW - int(hero.width * 0.86)
        hy_ = int(PH * 0.10)
        bg.alpha_composite(hero, (hx_, hy_))
        # 前景：从背景底部裁一条盖住人物下半身（融入感关键）
        fg_top = int(PH * 0.70)
        fg = bg.crop((0, fg_top, PW, PH)).filter(ImageFilter.GaussianBlur(2.5))
        fm = Image.new("L", fg.size, 0)
        fd = ImageDraw.Draw(fm)
        for yy in range(fg.height):
            a_y = min(1, yy / (fg.height * 0.55))
            for xx in range(fg.width):
                # 左侧与正下方保留前景，人物面部区域(x>0.62W 且 y<0.5H)不遮挡
                keep_x = 1.0 - max(0.0, (xx / fg.width - 0.60) / 0.40)
                a = int(255 * a_y * min(1.0, keep_x * 1.6))
                if a > 0:
                    fm.putpixel((xx, yy), a)
        fg = fg.convert("RGBA"); fg.putalpha(fm)
        bg.alpha_composite(fg, (0, fg_top))

    # 书法字
    chars_dir = "chars"
    os.makedirs(chars_dir, exist_ok=True)
    for i, ch in enumerate(list(args.title)[:6]):
        p = os.path.join(chars_dir, f"{i}_{ch}.png")
        render_char(ch, p, pal, SPEC["char_size"][min(i, 5)])
        im = Image.open(p).convert("RGBA")
        s = SPEC["char_size"][min(i, 5)]
        im = im.resize((s, s), Image.LANCZOS)
        y = SPEC["char_base_y"] + SPEC["char_offsets"][i % len(SPEC["char_offsets"])]
        bg.alpha_composite(im, (SPEC["char_x0"] + i * SPEC["char_gap"], y))

    d = ImageDraw.Draw(bg)
    ey = SPEC["eyebrow"]
    d.text((ey[0], ey[1]), args.competition, font=pick(FONT_BODY, ey[2]), fill=pal["gold"])
    st = SPEC["subtitle"]
    d.rectangle([st[0], st[1] - 44, st[0] + 380, st[1] - 40], fill=pal["gold"])
    d.text((st[0], st[1]), args.subtitle, font=pick(FONT_SUB, st[2]), fill=pal["text"])
    tg = SPEC["tagline"]
    d.text((tg[0], tg[1]), args.tagline, font=pick(FONT_BODY, tg[2]), fill=pal["muted"])
    ib = SPEC["info_bar"]
    bar = Image.new("RGBA", (PW, PH), (0, 0, 0, 0))
    ImageDraw.Draw(bar).rounded_rectangle([ib[0], ib[1], ib[2], ib[3]], radius=10,
                                          fill=pal["dark"] + (180,),
                                          outline=pal["gold"] + (120,), width=2)
    bg = Image.alpha_composite(bg, bar)
    ImageDraw.Draw(bg).text((ib[0] + 26, ib[1] + 20), args.info,
                            font=pick(FONT_BODY, SPEC["info_size"]), fill=pal["text"])
    # 统一暗角 + 颗粒
    vig = Image.new("L", (PW, PH), 0)
    ImageDraw.Draw(vig).ellipse([-PW * .2, -PH * .3, PW * 1.2, PH * 1.3], fill=255)
    vig = vig.filter(ImageFilter.GaussianBlur(PW // 10))
    dark = Image.new("RGB", (PW, PH), pal["dark"])
    bg = Image.composite(bg.convert("RGB"), dark, vig)
    return bg


def build_pptx(cover_path, out):
    full = Image.open(cover_path)
    w, h = full.size
    dark = full.point(lambda p: int(p * 0.25))
    os.makedirs("opening", exist_ok=True)
    frames = []
    for i, frac in enumerate(SPEC["open_steps"], 1):
        c = dark.copy()
        top = int(h * (1 - frac)) if frac else h
        if frac:
            c.paste(full.crop((0, top, w, h)), (0, top))
        p = f"opening/step{i}.jpg"; c.save(p, quality=90); frames.append(p)
    frames.append(cover_path)
    prs = Presentation(); prs.slide_width = Cm(W_CM); prs.slide_height = Cm(H_CM)
    for i, f in enumerate(frames):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        s.shapes.add_picture(f, 0, 0, Cm(W_CM), Cm(H_CM))
        old = s._element.find(qn("p:transition"))
        if old is not None:
            s._element.remove(old)
        t = etree.SubElement(s._element, qn("p:transition")); t.set("spd", "slow")
        etree.SubElement(t, qn("p:morph" if i == 3 else "p:fade"))
    prs.save(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--title", required=True)
    ap.add_argument("--subtitle", default="")
    ap.add_argument("--tagline", default="")
    ap.add_argument("--info", default="")
    ap.add_argument("--competition", default="第十届中国国际“互联网+”大学生创新创业大赛")
    ap.add_argument("--bg", required=True)
    ap.add_argument("--hero", default="")
    ap.add_argument("--palette", default="green", choices=list(PALETTES))
    ap.add_argument("--out", default="封面.pptx")
    ap.add_argument("--dump-spec", action="store_true", help="输出参数指纹 JSON")
    a = ap.parse_args()
    if a.dump_spec:
        print(json.dumps(SPEC, ensure_ascii=False, indent=2))
        return
    pal = PALETTES[a.palette]
    os.makedirs("assets", exist_ok=True)
    cover = "assets/cover_composed.jpg"
    compose(a, pal).save(cover, quality=93)
    build_pptx(cover, a.out)
    print(f"已生成 {a.out}（4 页：3 页开屏 + 正式封面）；合成图 {cover}")


if __name__ == "__main__":
    main()
