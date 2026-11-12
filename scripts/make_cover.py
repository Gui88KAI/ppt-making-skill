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
    hero_bright=1.02, hero_feather=1.2, hero_match=0.32, hero_sharpen=1.35,
    hero_scale=0.98, hero_rim=0.85, hero_shadow=True, fg_ratio=0.82,
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


def render_char(ch, path, pal, size, color=None, font_path=None):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    f = pick([font_path] if font_path and os.path.exists(font_path) else FONT_CANDIDATES, size)
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
    c0, c1 = (color, tuple(min(255, int(v * 1.12)) for v in color)) if color else (SPEC["grad_top"], SPEC["grad_bottom"])
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


def hero_mask(size):
    """实心椭圆遮罩：中心不透明，只有边缘 12% 羽化（人物要清晰）。"""
    w, h = size
    m = Image.new("L", (w, h), 0)
    d = ImageDraw.Draw(m)
    d.ellipse([w * .02, h * .01, w * .98, h * 1.02], fill=255)
    return m.filter(ImageFilter.GaussianBlur(max(6, int(min(w, h) * .06))))


def rim_light(hero, color, strength):
    """轮廓光：膨胀 alpha 减原 alpha 得边缘，涂暖色。"""
    import numpy as np
    a = hero.split()[3]
    dil = a.filter(ImageFilter.MaxFilter(9))
    edge = Image.fromarray(np.clip(np.asarray(dil, "int16") - np.asarray(a, "int16"), 0, 255).astype("uint8"))
    edge = edge.filter(ImageFilter.GaussianBlur(6))
    rim = Image.new("RGBA", hero.size, color + (0,))
    rim.putalpha(edge.point(lambda v: int(v * strength)))
    return Image.alpha_composite(hero, rim)




def title_layout(chars_png, sizes, jitter, safe_left=190, safe_right=None,
                 gap=6, base_y=180, center_y=None):
    """把逐字图片排进安全区：紧凑、互不遮挡、不压人物。

    返回 [(png_path, x, y, w, h), ...]，坐标为 1920×1080 像素系。
    """
    if safe_right is None:
        safe_right = PW - 830          # 右侧给人物留 830px，标题不压人
    # 1) 先按目标字号取字形实际宽度
    items = []
    for i, p in enumerate(chars_png):
        im = Image.open(p).convert("RGBA")
        target = sizes[min(i, len(sizes) - 1)]
        im = im.resize((target, target), Image.LANCZOS)
        bb = im.split()[3].getbbox()
        if bb:
            im = im.crop(bb)
        items.append((p, im, bb))
    # 2) 计算总宽，若超出安全区则整体等比缩小
    total = sum(im.width for _, im, _ in items) + gap * (len(items) - 1)
    avail = safe_right - safe_left
    scale = 1.0 if total <= avail else avail / total
    total_w = total * scale
    # 3) 水平居中于安全区，逐字累加
    cursor = safe_left + max(0, (avail - total_w) / 2)
    out = []
    max_h = max(int(im.height * scale) for _, im, _ in items)
    for i, (p, im, bb) in enumerate(items):
        w = int(im.width * scale); h = int(im.height * scale)
        im2 = im.resize((w, h), Image.LANCZOS)
        dy = jitter[i % len(jitter)] if jitter else 0
        if center_y is not None:
            y = int(center_y - max_h / 2 + (max_h - h) / 2 + dy)
        else:
            y = int(base_y + dy)
        out.append((p, int(cursor), y, w, h, im2))
        cursor += w + gap * scale
    return out

def compose(args, pal, skip_title=False, skip_text=False):
    # 背景：按 1.35 倍高度放大后取窗口，支持上下平移（把主体推到想要的位置）
    _src = Image.open(args.bg).convert("RGB")
    _z = 1.35
    _sw, _sh = int(PW * _z), int(PH * _z)
    _r = max(_sw / _src.width, _sh / _src.height)
    _src = _src.resize((int(_src.width * _r) + 1, int(_src.height * _r) + 1), Image.LANCZOS)
    _cx = (_src.width - PW) // 2
    _pan = max(0.0, min(1.0, getattr(args, "bg_pan", 0.32)))
    _cy = int((_src.height - PH) * _pan)
    bg = _src.crop((_cx, _cy, _cx + PW, _cy + PH))
    bg = ImageEnhance.Brightness(bg).enhance(SPEC["bg_darken"])
    bg = ImageEnhance.Color(bg).enhance(SPEC["bg_satur"])
    bg = ImageEnhance.Contrast(bg).enhance(SPEC["bg_contrast"])
    shade = Image.new("RGBA", (PW, PH), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shade)
    for x in range(PW):
        t = max(0, 1 - x / (PW * 0.66))
        sd.line([(x, 0), (x, PH)], fill=pal["dark"] + (int(190 * t),))
    bg = Image.alpha_composite(bg.convert("RGBA"), shade)

    # 人物：实心羽化 + 轻度色彩匹配 + 轮廓光 + 投影（务必清晰）
    if args.hero and os.path.exists(args.hero):
        hero = Image.open(args.hero).convert("RGBA")
        hh = int(PH * SPEC["hero_scale"])
        r = hh / hero.height
        hero = hero.resize((int(hero.width * r), hh), Image.LANCZOS)
        a0 = hero.split()[3]
        is_cutout = a0.getextrema()[0] < 250              # 是否为抠好的透明图
        if is_cutout:
            m = a0.filter(ImageFilter.GaussianBlur(1.5))  # 抠图只需极轻羽化
        else:
            m = hero_mask(hero.size)
        # 色彩匹配只作用于 RGB，绝不能让 alpha 丢失
        matched = color_match(hero.convert("RGB"), bg, SPEC["hero_match"])
        matched = ImageEnhance.Brightness(matched).enhance(SPEC["hero_bright"])
        matched = ImageEnhance.Sharpness(matched).enhance(SPEC["hero_sharpen"])
        matched = ImageEnhance.Contrast(matched).enhance(1.06)
        hero = matched.convert("RGBA")
        hero.putalpha(m)
        if is_cutout:
            # 抠图边缘去掉残影：alpha < 24 的全清掉
            a = hero.split()[3].point(lambda v: 0 if v < 24 else v)
            hero.putalpha(a)
        hero = rim_light(hero, pal["gold"], SPEC["hero_rim"])
        hx_ = PW - int(hero.width * 0.94)
        hy_ = int(PH * 0.06)
        if SPEC["hero_shadow"] and not is_cutout:
            sh = Image.new("RGBA", (PW, PH), (0, 0, 0, 0))
            ImageDraw.Draw(sh).ellipse([hx_ + hero.width * .1, hy_ + hero.height * .82,
                                        hx_ + hero.width * .95, hy_ + hero.height * .99],
                                       fill=(0, 0, 0, 110))
            bg = Image.alpha_composite(bg, sh.filter(ImageFilter.GaussianBlur(40)))
        bg.alpha_composite(hero, (hx_, hy_))
        # 前景只盖最下方 18%，不再糊人
        fg_top = int(PH * SPEC["fg_ratio"])
        fg = bg.crop((0, fg_top, PW, PH)).filter(ImageFilter.GaussianBlur(2.0))
        fm = Image.new("L", fg.size, 255)
        fd = ImageDraw.Draw(fm)
        for yy in range(fg.height):
            fd.line([(0, yy), (fg.width, yy)], fill=int(255 * min(1, yy / (fg.height * 0.6))))
        fg = fg.convert("RGBA"); fg.putalpha(fm)
        bg.alpha_composite(fg, (0, fg_top))

    # 书法字（四种版式）
    chars_dir = "chars"
    os.makedirs(chars_dir, exist_ok=True)
    chars = list(args.title)[:6]
    # 逐字变化参数（横排整体不变，只让每个字有差异）
    col_spec = args.char_colors.split(",") if args.char_colors else []
    fnt_spec = args.char_fonts.split(",") if args.char_fonts else []
    jit_spec = [float(v) for v in args.char_jitter.split(",")] if args.char_jitter else SPEC["char_offsets"]
    title_layer = Image.new('RGBA', (PW, PH), (0, 0, 0, 0))
    imgs = []
    for i, ch in enumerate(chars):
        col = None
        if col_spec:
            c = col_spec[i % len(col_spec)].strip()
            col = hexrgb(c) if c.startswith("#") else pal.get(c.replace("gold2", "gold2"), pal["gold"])
        fpath = None
        if fnt_spec:
            name = fnt_spec[i % len(fnt_spec)].strip()
            mapping = {
                "流云楷": "/usr/share/fonts/truetype/custom/演示流云楷--淘宝店：烦烦优创馆.ttf",
                "许静行楷": "/usr/share/fonts/truetype/custom/汉仪许静行楷W.ttf",
                "夏行楷": "/usr/share/fonts/truetype/custom/演示夏行楷.ttf",
                "镇魂行楷": "/usr/share/fonts/truetype/custom/演示镇魂行楷--淘宝店：烦烦优创馆.ttf",
                "悠然小楷": "/usr/share/fonts/truetype/custom/演示悠然小楷.ttf",
            }
            fpath = mapping.get(name)
        p = os.path.join(chars_dir, f"{i}_{ch}.png")
        render_char(ch, p, pal, 460, color=col, font_path=fpath)
        imgs.append(Image.open(p).convert("RGBA"))

    # 横排：紧凑排进安全区（与 PPT 端同一套计算）
    sizes = [360, 340, 322, 306, 292, 280]
    if not skip_title:
        hero_left = PW - 830
        cy = args.title_y if args.title_y else 430        # 默认放到左侧空白区中部
        for p, x, y, w, h, im2 in title_layout([os.path.join(chars_dir, f"{i}_{ch}.png")
                                                for i, ch in enumerate(chars)],
                                               sizes, jit_spec, 190, hero_left - 20, 4, 185, cy):
            title_layer.alpha_composite(im2, (x, y))

    if not skip_title:
        bg.alpha_composite(title_layer)

    d = ImageDraw.Draw(bg)
    dy = 0                                 # 单行横排：文字块位置固定
    ey = SPEC["eyebrow"]
    d.text((ey[0], ey[1]), args.competition, font=pick(FONT_BODY, ey[2]), fill=pal["gold"])
    if skip_text:
        return bg.convert("RGB"), title_layer
    st = (SPEC["subtitle"][0], SPEC["subtitle"][1] + dy, SPEC["subtitle"][2])
    d.rectangle([st[0], st[1] - 44, st[0] + 380, st[1] - 40], fill=pal["gold"])
    d.text((st[0], st[1]), args.subtitle, font=pick(FONT_SUB, st[2]), fill=pal["text"])
    tg = (SPEC["tagline"][0], SPEC["tagline"][1] + dy, SPEC["tagline"][2])
    d.text((tg[0], tg[1]), args.tagline, font=pick(FONT_BODY, tg[2]), fill=pal["muted"])
    ib = (SPEC["info_bar"][0], SPEC["info_bar"][1] + dy, SPEC["info_bar"][2], SPEC["info_bar"][3] + dy)
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
    return bg, title_layer


def build_pptx(cover_path, out, bg_path=None, args=None, pal=None):
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
    for i, f in enumerate(frames[:-1]):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        s.shapes.add_picture(f, 0, 0, Cm(W_CM), Cm(H_CM))
        old = s._element.find(qn("p:transition"))
        if old is not None:
            s._element.remove(old)
        t = etree.SubElement(s._element, qn("p:transition")); t.set("spd", "slow")
        etree.SubElement(t, qn("p:fade"))

    # 最后一页：背景图 + 每个字一个独立图片框（可拖动）+ 文字框
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(bg_path or cover_path, 0, 0, Cm(W_CM), Cm(H_CM))
    chars = list(args.title)[:6]
    sizes = [int(v * args.title_scale) for v in [360, 340, 322, 306, 292, 280]]
    jit = [float(v) for v in args.char_jitter.split(",")] if args.char_jitter else SPEC["char_offsets"]
    pngs = [os.path.join("chars", f"{i}_{ch}.png") for i, ch in enumerate(chars)]
    for p, x, y, w, h, _im in title_layout(pngs, sizes, jit, 190, PW - 840, 4, 185, args.title_y or 430):
        if not os.path.exists(p):
            continue
        pic = s.shapes.add_picture(p, Cm(x / PW * W_CM), Cm(y / PH * H_CM),
                                   Cm(w / PW * W_CM), Cm(h / PH * H_CM))
        pic.name = "标题-" + os.path.basename(p).split("_")[-1].split(".")[0]
    # 文字框（可编辑）
    def tb(txt_, x, y, size, color, font, bold=False):
        box = s.shapes.add_textbox(Cm(x / PW * W_CM), Cm(y / PH * H_CM),
                                   Cm(1200 / PW * W_CM), Cm(90 / PH * H_CM))
        tf = box.text_frame; tf.word_wrap = True
        para = tf.paragraphs[0]; run = para.add_run(); run.text = txt_
        run.font.size = Pt(size * W_CM / PW * 2.5)
        run.font.bold = bold; run.font.name = font
        run.font.color.rgb = RGBColor(*color)
        return box
    tb(args.subtitle, SPEC["subtitle"][0], SPEC["subtitle"][1] + 8, SPEC["subtitle"][2],
       pal["text"], "思源宋体 CN Heavy")
    tb(args.tagline, SPEC["tagline"][0], SPEC["tagline"][1], SPEC["tagline"][2],
       pal["muted"], "阿里巴巴普惠体 L")
    tb(args.info, SPEC["info_bar"][0] + 26, SPEC["info_bar"][1] + 20, SPEC["info_size"],
       pal["text"], "阿里巴巴普惠体 L")
    old = s._element.find(qn("p:transition"))
    if old is not None:
        s._element.remove(old)
    t = etree.SubElement(s._element, qn("p:transition")); t.set("spd", "slow")
    etree.SubElement(t, qn("p:morph"))
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
    ap.add_argument("--char-colors", default="",
                    help='逐字颜色，逗号分隔，如 "gold,text,gold2,gold" 或 "#D9B875,#F4EFDF"')
    ap.add_argument("--char-fonts", default="",
                    help='逐字字体，逗号分隔，可选 流云楷/许静行楷/夏行楷/镇魂行楷/悠然小楷')
    ap.add_argument("--bg-pan", type=float, default=0.32,
                    help="背景纵向平移 0=顶部 0.5=居中 1=底部（调高主体用更小值）")
    ap.add_argument("--title-scale", type=float, default=1.18, help="标题字号放大系数")
    ap.add_argument("--title-y", type=int, default=0, help="标题垂直中心（px，0=自动居中）")
    ap.add_argument("--char-jitter", default="",
                    help='逐字上下微调（px），逗号分隔，如 "0,-18,12,-8"')
    ap.add_argument("--out", default="封面.pptx")
    ap.add_argument("--dump-spec", action="store_true", help="输出参数指纹 JSON")
    a = ap.parse_args()
    if a.dump_spec:
        print(json.dumps(SPEC, ensure_ascii=False, indent=2))
        return
    pal = PALETTES[a.palette]
    os.makedirs("assets", exist_ok=True)
    cover = "assets/cover_composed.jpg"
    full, title_layer = compose(a, pal)

    # 背景层（不含标题）——给 PPT 用，标题单独放
    bg_only = compose(a, pal, skip_title=True)[0]
    bg_only_path = "assets/cover_bg.jpg"
    bg_only.save(bg_only_path, quality=93)

    # 合成预览（含标题）
    merged = full.copy().convert("RGBA")
    merged.alpha_composite(title_layer)
    merged.convert("RGB").save(cover, quality=93)

    build_pptx(cover, a.out, bg_path=bg_only_path, args=a, pal=pal)
    print(f"已生成 {a.out}（4 页：3 页开屏 + 正式封面）；合成图 {cover}")


if __name__ == "__main__":
    main()
