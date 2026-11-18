#!/usr/bin/env python3
"""生成 01–05 共用的 16:9 封面，默认输出一页可编辑 PPTX。

示例：python scripts/make_cover.py --type 04 --title 通信之路
      --subtitle 通信工程专业职业生涯规划 --bg lab.jpg --out 封面.pptx
选传 --hero 可把独立人物/产品放到右侧。也可直接使用已经构图的整张场景图。
"""
import argparse
from io import BytesIO
import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
PW, PH = 1920, 1080
W, H = Inches(40 / 3), Inches(7.5)
GOLD = (229, 184, 99)
IVORY = (255, 250, 234)
PALETTES = {
    "green": {"dark": (7, 19, 14), "muted": (206, 223, 211)},
    "blue": {"dark": (6, 21, 34), "muted": (204, 223, 233)},
    "red": {"dark": (35, 15, 16), "muted": (231, 216, 208)},
}
TYPE_PRESETS = {
    "01": ("大学生创新创业项目", "green", "culture"),
    "02": ("商业计划书", "blue", "business"),
    "03": ("融资路演", "blue", "tech"),
    "04": ("答辩汇报", "blue", "business"),
    "05": ("课程汇报", "green", "culture"),
}
BODY_STYLES = {
    "culture": "STSong", "tech": "Microsoft YaHei",
    "business": "Microsoft YaHei", "royal": "STSong",
}
SPEC = {
    "title_left": 222, "title_right": 1080, "title_center_y": 417, "title_height": 240,
    "char_offsets": [10, -12, 11, -9], "char_sizes": [1.0, 0.96, 0.92, 0.98],
    "eyebrow": [249, 85, 30], "rule": [248, 696, 380, 5],
    "subtitle": [248, 741, 56], "tagline": [250, 846, 30],
    "info_bar": [249, 912, 1152, 74], "info_size": 26,
    "grad_top": list(IVORY), "grad_bottom": [222, 166, 71],
    "hero_scale": 0.95, "open_steps": [0.96, 0.72, 0.43],
}
FONT_DIRS = (ROOT / "templates/fonts", Path("/usr/share/fonts/truetype/custom"))
FONT_NAMES = {
    "流云楷": "演示流云楷.ttf",
    "许静行楷": "汉仪许静行楷W.ttf",
    "夏行楷": "演示夏行楷.ttf",
    "镇魂行楷": "演示镇魂行楷.ttf",
    "悠然小楷": "悠然小楷.ttf",
}
TITLE_FONTS = [
    *(folder / filename for folder in FONT_DIRS for filename in
      ("演示流云楷.ttf", "演示流云楷--淘宝店：烦烦优创馆.ttf", "汉仪许静行楷W.ttf")),
    Path("C:/Windows/Fonts/STXINGKA.TTF"),
    Path("C:/Windows/Fonts/simkai.ttf"),
    Path("/System/Library/Fonts/STHeiti Light.ttc"),
    Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
    Path("/system/fonts/NotoSansCJK-Regular.ttc"),
]

def select_font(preferred=None):
    if preferred:
        filename = FONT_NAMES.get(preferred, preferred)
        candidates = [Path(preferred), *(folder / filename for folder in FONT_DIRS)]
    else:
        candidates = TITLE_FONTS
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise ValueError("找不到标题字体，请用 --title-font 指定支持题名字形的字体文件。")

def render_char(char, font_path, bottom):
    """用实际笔画的包围盒归一化渐变，让每个字都有从白到金的完整过渡。"""
    font = ImageFont.truetype(str(font_path), 850)
    mask = font.getmask(char, mode="L")
    pixels = Image.frombytes("L", mask.size, bytes(mask))
    bounds = pixels.getbbox()
    if not bounds:
        raise ValueError(f"无法渲染标题字 {char!r}，请检查标题字体。")
    pixels = pixels.crop(bounds)
    height = max(1, pixels.height - 1)
    gradient = Image.new("RGB", (1, pixels.height))
    gradient.putdata([
        tuple(round(start + (end - start) * y / height) for start, end in zip(IVORY, bottom))
        for y in range(pixels.height)
    ])
    glyph = gradient.resize(pixels.size, Image.Resampling.NEAREST).convert("RGBA")
    glyph.putalpha(pixels)
    return glyph

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

def compose_background(args, palette):
    with Image.open(args.bg) as source:
        fitted = ImageOps.fit(
            ImageOps.exif_transpose(source).convert("RGB"), (PW, PH),
            method=Image.Resampling.LANCZOS, centering=(0.5, args.bg_pan),
        )
    bg = fitted.convert("RGBA")
    shade = Image.new("L", (PW, 1))
    shade.putdata([round(170 * max(0, 1 - x / (PW * 0.68))) for x in range(PW)])
    tint = Image.new("RGBA", bg.size, palette["dark"])
    tint.putalpha(shade.resize(bg.size))
    bg = Image.alpha_composite(bg, tint)
    if not args.hero:
        return bg.convert("RGB")
    with Image.open(args.hero) as source:
        hero = ImageOps.contain(
            ImageOps.exif_transpose(source).convert("RGBA"),
            (int(PW * .45), int(PH * SPEC["hero_scale"])),
            method=Image.Resampling.LANCZOS,
        )
    alpha = hero.getchannel("A")
    if alpha.getextrema()[0] >= 250:
        alpha = hero_mask(hero.size)
    matched = color_match(hero, bg, .20)
    matched = ImageEnhance.Brightness(matched).enhance(1.02).convert("RGBA")
    matched.putalpha(alpha)
    matched = rim_light(matched, GOLD, .35)
    bg.alpha_composite(matched, (PW - matched.width + 12, PH - matched.height + 8))
    return bg.convert("RGB")

def title_layout(args):
    letters = [char for char in args.title if not char.isspace()]
    sizes = SPEC["char_sizes"]
    jitter = [float(value) for value in args.char_jitter.split(",")] if args.char_jitter else SPEC["char_offsets"]
    overrides = args.char_fonts.split(",") if args.char_fonts else []
    fonts = [select_font(value.strip()) for value in overrides]
    font = select_font(args.title_font) if args.title_font or not fonts else fonts[0]
    endings = [SPEC["grad_bottom"]]
    if args.char_colors:
        known = {"gold": GOLD, "text": IVORY, "gold2": tuple(SPEC["grad_bottom"])}
        endings = [
            tuple(bytes.fromhex(token[1:])) if token.startswith("#") and len(token) == 7 else known[token]
            for token in (part.strip() for part in args.char_colors.split(","))
        ]
    glyphs = []
    for i, char in enumerate(letters):
        glyph = render_char(char, fonts[i % len(fonts)] if fonts else font, endings[i % len(endings)])
        height = max(1, round(SPEC["title_height"] * args.title_scale * sizes[i % len(sizes)]))
        glyph.thumbnail((round(glyph.width * height / glyph.height), height), Image.Resampling.LANCZOS)
        glyphs.append(glyph)
    gap = min(19, 150 // len(glyphs))
    width = sum(im.width for im in glyphs) + gap * (len(glyphs) - 1)
    available = SPEC["title_right"] - SPEC["title_left"]
    scale = min(1, available / width)
    center_y = args.title_y or SPEC["title_center_y"]
    x = SPEC["title_left"]
    layout = []
    for i, (char, im) in enumerate(zip(letters, glyphs)):
        size = (max(1, round(im.width * scale)), max(1, round(im.height * scale)))
        resized = im.resize(size, Image.Resampling.LANCZOS)
        y = round(center_y - resized.height / 2 + jitter[i % len(jitter)] * scale)
        layout.append((char, resized, round(x), y))
        x += resized.width + gap * scale
    return layout

def emu_x(px):
    return round(W * px / PW)

def emu_y(px):
    return round(H * px / PH)

def add_image(slide, im, x, y, w, h, alt):
    stream = BytesIO()
    if im.mode == "RGBA":
        im.save(stream, format="PNG")
    else:
        im.save(stream, format="JPEG", quality=92, optimize=True)
    stream.seek(0)
    pic = slide.shapes.add_picture(stream, emu_x(x), emu_y(y), emu_x(w), emu_y(h))
    pic.name = alt
    pic._element.nvPicPr.cNvPr.set("descr", alt)
    return pic

def add_text(slide, value, x, y, px_size, width, tint, font, *, bold=False):
    if not value:
        return
    # 约束单行长度，允许较长的用户标题有足够横向空间。
    units = sum(1 if ord(char) > 127 else .58 for char in value)
    px_size = min(px_size, int(width / max(units, 1)))
    box = slide.shapes.add_textbox(emu_x(x), emu_y(y), emu_x(width), emu_y(px_size * 1.35))
    frame = box.text_frame
    frame.word_wrap = False
    frame.auto_size = MSO_AUTO_SIZE.NONE
    frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    run = frame.paragraphs[0].add_run()
    run.text = value
    run.font.name = font
    run.font.size = Pt(px_size / 2)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*tint)
    props = run._r.get_or_add_rPr()
    east_asian = OxmlElement("a:ea")
    east_asian.set("typeface", font)
    props.append(east_asian)

def add_block(slide, area, tint, *, opacity=1.0, border=False):
    x, y, w, h = area
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if border else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, emu_x(x), emu_y(y), emu_x(w), emu_y(h))
    shape._element.remove(shape._element.find(qn("p:style")))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*tint)
    if opacity < 1:
        alpha = OxmlElement("a:alpha")
        alpha.set("val", str(round(opacity * 100000)))
        shape._element.find(".//" + qn("a:srgbClr")).append(alpha)
    if border:
        shape.adjustments[0] = .12
        shape.line.color.rgb = RGBColor(*GOLD)
        shape.line.width = Pt(.7)
    else:
        shape.line.fill.background()
    return shape

def build_pptx(args):
    palette = PALETTES[args.palette]
    scene = compose_background(args, palette)
    letters = title_layout(args)
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    prs.core_properties.title = args.title
    font = BODY_STYLES[args.style]
    for shade in (*SPEC["open_steps"], 0) if args.opening else (0,):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_image(slide, scene, 0, 0, PW, PH, "全幅场景")
        for char, image, x, y in letters:
            add_image(slide, image, x, y, image.width, image.height, "标题-" + char)
        add_text(slide, args.competition, *SPEC["eyebrow"], 1100, GOLD, font)
        add_block(slide, SPEC["rule"], GOLD)
        add_text(slide, args.subtitle, *SPEC["subtitle"], 1030, IVORY, font, bold=True)
        add_text(slide, args.tagline, *SPEC["tagline"], 1100, palette["muted"], font)
        if args.info:
            add_block(slide, SPEC["info_bar"], palette["dark"], opacity=.70, border=True)
            add_text(slide, args.info, 273, 927, SPEC["info_size"], 1100, IVORY, font)
        if shade:
            add_block(slide, (0, 0, PW, PH), palette["dark"], opacity=shade)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(args.out)
    return len(prs.slides)

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--type", choices=TYPE_PRESETS, default="01")
    parser.add_argument("--title", help="左侧艺术主标题")
    parser.add_argument("--subtitle", default="", help="完整题名或一句定位")
    parser.add_argument("--tagline", default="", help="副标题下方的简短说明")
    parser.add_argument("--competition", help="顶部真实场合，省略时使用通用类型名")
    parser.add_argument("--info", default="", help="底部信息条中的已确认字段")
    parser.add_argument("--bg", type=Path, help="铺满画布的场景图")
    parser.add_argument("--hero", type=Path, help="可选的右侧人物或产品图")
    parser.add_argument("--palette", choices=PALETTES)
    parser.add_argument("--style", choices=BODY_STYLES)
    parser.add_argument("--title-font", help="艺术标题字体文件路径")
    parser.add_argument("--char-colors", default="", help="可选的旧版逐字渐变末端色；默认所有字同一白金渐变")
    parser.add_argument("--char-fonts", default="", help="可选的逐字字体名或路径，逗号分隔")
    parser.add_argument("--bg-pan", type=float, default=.50, help="背景纵向裁切位置，0 至 1")
    parser.add_argument("--title-scale", type=float, default=1.0)
    parser.add_argument("--title-y", type=int, default=0, help="标题垂直中心的像素位置")
    parser.add_argument("--char-jitter", default="", help="逐字上下微调的像素值，逗号分隔")
    parser.add_argument("--opening", action="store_true", help="另生成三页递进开屏")
    parser.add_argument("--out", type=Path, default=Path("封面.pptx"))
    parser.add_argument("--dump-spec", action="store_true")
    args = parser.parse_args()
    if args.dump_spec:
        print(json.dumps(SPEC, ensure_ascii=False, indent=2))
        return
    if not args.title or not args.title.strip() or not args.bg:
        parser.error("生成封面时需要 --title 和 --bg。")
    if not 0 <= args.bg_pan <= 1 or args.title_scale <= 0:
        parser.error("--bg-pan 必须在 0 至 1 之间，--title-scale 必须大于 0。")
    eyebrow, palette, style = TYPE_PRESETS[args.type]
    args.competition = eyebrow if args.competition is None else args.competition
    args.palette = args.palette or palette
    args.style = args.style or style
    try:
        count = build_pptx(args)
    except (OSError, ValueError, KeyError) as error:
        parser.error(str(error))
    print(f"已生成 {args.out}（{count} 页）")

if __name__ == "__main__":
    main()
