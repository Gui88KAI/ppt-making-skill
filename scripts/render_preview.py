import os, sys
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation

EMU = 360000.0
NS = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
CJK = "/system/fonts/NotoSansCJK-Regular.ttc"
CJKB = "/system/fonts/NotoSansCJK-Bold.ttc"



# —— 字体名 → 真实字体文件（预览用真实字体，不再一律 Noto）——
import glob as _glob
_FT = "/usr/share/fonts/truetype/custom/"
_FONT_MAP = [
    ("思源宋体 CN Heavy", _FT + "思源宋体 CN Heavy.otf"),
    ("思源宋体 CN Light", _FT + "思源宋体 CN Light.otf"),
    ("思源宋体 CN", _FT + "思源宋体 CN.otf"),
    ("华文楷体", _FT + "华文楷体.TTF"),
    ("汉仪雅酷黑 75W", _FT + "汉仪雅酷黑 75W.ttf"),
    ("阿里巴巴普惠体 B", _FT + "阿里巴巴普惠体 B.ttf"),
    ("阿里巴巴普惠体 L", _FT + "阿里巴巴普惠体 L.ttf"),
    ("阿里巴巴普惠体 R", _FT + "阿里巴巴普惠体 R.ttf"),
    ("演示流云楷", _FT + "演示流云楷--淘宝店：烦烦优创馆.ttf"),
    ("演示镇魂行楷", _FT + "演示镇魂行楷--淘宝店：烦烦优创馆.ttf"),
    ("Montserrat", _FT + "Montserrat--淘宝店：烦烦优创馆.ttf"),
    ("Bahnschrift", _FT + "Bahnschrift.ttf"),
]


def _resolve_font(name, bold):
    if not name:
        return None
    base = name.replace(" ", "").lower()
    for k, path in _FONT_MAP:
        if k.replace(" ", "").lower() in base.replace(" ", ""):
            if os.path.exists(path):
                return path
    # 粗体变体优先
    if bold:
        for cand in (_FT + "阿里巴巴普惠体 B.ttf", CJKB):
            if os.path.exists(cand):
                return cand
    return CJK


def fnt(pt, bold, ppc, name=None):
    px = max(int(round(pt / 72 * 2.54 * ppc)), 5)
    path = _resolve_font(name, bold) or (CJKB if bold and os.path.exists(CJKB) else CJK)
    try:
        return ImageFont.truetype(path, px)
    except Exception:
        return ImageFont.truetype(CJKB if bold and os.path.exists(CJKB) else CJK, px)


def _fnt_old(pt, bold, ppc):
    px = max(int(round(pt / 72 * 2.54 * ppc)), 5)
    return ImageFont.truetype(CJKB if bold and os.path.exists(CJKB) else CJK, px)


def rgb(c):
    try:
        return (c[0], c[1], c[2])
    except Exception:
        return None


def fillc(sh):
    try:
        f = sh.fill
        if f.type is None:
            return None
        try:
            return rgb(f.fore_color.rgb)
        except Exception:
            try:
                st = f.gradient_stops
                return rgb(st[0].color.rgb)
            except Exception:
                return None
    except Exception:
        return None


def alpha_of(sh):
    try:
        al = sh._element.find('.//' + NS + 'solidFill/' + NS + 'srgbClr/' + NS + 'alpha')
        if al is not None:
            return int(int(al.get('val')) / 100000 * 255)
    except Exception:
        pass
    return 255


def geom(sh):
    """返回几何类型：ellipse / roundRect / rect / 其它。"""
    try:
        g = sh._element.find('.//' + NS + 'prstGeom')
        if g is not None:
            return g.get('prst') or 'rect'
    except Exception:
        pass
    return 'rect'


def rotation(sh):
    try:
        x = sh._element.find('.//' + NS + 'xfrm')
        if x is not None and x.get('rot'):
            return int(x.get('rot')) / 60000.0
    except Exception:
        pass
    return 0.0


def line_color(sh):
    try:
        ln = sh._element.find('.//' + NS + 'ln')
        if ln is None:
            return None
        f = ln.find(NS + 'solidFill')
        if f is None:
            return None
        c = f.find(NS + 'srgbClr')
        if c is None:
            return None
        return rgb_css(c.get('val'))
    except Exception:
        return None


def rgb_css(v):
    try:
        return (int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))
    except Exception:
        return None


def xf(sh):
    try:
        x = sh._element.find('.//' + NS + 'xfrm')
        if x is None:
            return None
        off = x.find(NS + 'off'); ext = x.find(NS + 'ext')
        co = x.find(NS + 'chOff'); ce = x.find(NS + 'chExt')
        if ext is None or ce is None:
            return None
        return (int(off.get('x')) / EMU, int(off.get('y')) / EMU,
                int(ext.get('cx')) / EMU, int(ext.get('cy')) / EMU,
                int(co.get('x')) / EMU, int(co.get('y')) / EMU,
                max(int(ce.get('cx')) / EMU, 0.01), max(int(ce.get('cy')) / EMU, 0.01))
    except Exception:
        return None


def text_lines(d, tf, box, ppc, y0=None):
    """在 box 内绘制文本框内容。"""
    if not (tf.text.strip()):
        return
    if y0 is None:
        y0 = box[1]
        try:
            if 'MIDDLE' in str(tf.vertical_anchor):
                n = max(1, sum(1 for pa in tf.paragraphs if pa.runs))
                sz = 18
                for pa in tf.paragraphs:
                    for r in pa.runs:
                        if r.font.size:
                            sz = r.font.size.pt
                            break
                    if sz != 18:
                        break
                lh = sz / 72 * 2.54 * ppc * 1.3
                y0 = box[1] + ((box[3] - box[1]) - n * lh) / 2
        except Exception:
            pass
    for pa in tf.paragraphs:
        if not pa.runs:
            continue
        sz = 18; bold = False; col = (40, 40, 40); fname = None
        for r in pa.runs:
            if r.font.size:
                sz = r.font.size.pt
            bold = bool(r.font.bold) or bold
            try:
                fname = r.font.name or fname
            except Exception:
                pass
            try:
                if r.font.color and r.font.color.type is not None:
                    col = rgb(r.font.color.rgb)
            except Exception:
                pass
            break
        f = fnt(sz, bold, ppc, fname)
        lh = int(sz / 72 * 2.54 * ppc * 1.3)
        txt = ''.join(r.text for r in pa.runs)
        w = box[2] - box[0]
        lines = []; cur = ''
        for ch in txt:
            if d.textlength(cur + ch, font=f) > w:
                lines.append(cur); cur = ch
            else:
                cur += ch
        lines.append(cur)
        al = str(pa.alignment) if pa.alignment is not None else "LEFT"
        for ln in lines:
            if y0 > box[3]:
                break
            tw = d.textlength(ln, font=f)
            if "CENTER" in al:
                tx = box[0] + (box[2] - box[0] - tw) / 2
            elif "RIGHT" in al:
                tx = box[2] - tw
            else:
                tx = box[0]
            d.text((tx, y0), ln, font=f, fill=col + (255,))
            y0 += lh


def draw_node(layer, sh, box, ppc):
    """在 layer 上画一个形状（不含旋转）。"""
    d = ImageDraw.Draw(layer, 'RGBA')
    g = geom(sh)
    c = fillc(sh)
    a = alpha_of(sh)
    lc = line_color(sh)
    outline = (lc + (255,)) if lc else None
    if c:
        try:
            if g in ('ellipse', 'oval'):
                d.ellipse(box, fill=c + (a,))
            elif g in ('roundRect',):
                r = max(2, int(min(box[2] - box[0], box[3] - box[1]) * 0.12))
                d.rounded_rectangle(box, radius=r, fill=c + (a,))
            else:
                d.rectangle(box, fill=c + (a,))
        except Exception:
            pass
    if outline:
        try:
            if g in ('ellipse', 'oval'):
                d.ellipse(box, outline=outline, width=2)
            elif g in ('roundRect',):
                d.rounded_rectangle(box, radius=max(2, int(min(box[2]-box[0], box[3]-box[1]) * 0.12)),
                                    outline=outline, width=2)
            else:
                d.rectangle(box, outline=outline, width=2)
        except Exception:
            pass
    if sh.has_text_frame:
        text_lines(d, sh.text_frame, box, ppc)


def draw(img, shapes, ppc, T=lambda x, y: (x, y)):
    for sh in shapes:
        if sh.shape_type == 6:  # GROUP
            t = xf(sh)
            if t is None:
                draw(img, sh.shapes, ppc, T); continue
            ox, oy, ex, ey, cx, cy, cex, cey = t
            sx, sy = ex / cex, ey / cey

            def T2(x, y, T=T):
                ax, ay = T(ox, oy)
                return (ax + (x - cx) * sx, ay + (y - cy) * sy)
            draw(img, sh.shapes, ppc, T2)
            continue
        try:
            x, y = T(sh.left / EMU, sh.top / EMU)
            x2, _ = T(sh.left / EMU + sh.width / EMU, sh.top / EMU)
            _, y2 = T(sh.left / EMU, sh.top / EMU + sh.height / EMU)
            w = max(x2 - x, 0.02); h = max(y2 - y, 0.02)
        except Exception:
            continue
        box = [int(x * ppc), int(y * ppc), int((x + w) * ppc), int((y + h) * ppc)]
        bw = max(box[2] - box[0], 1); bh = max(box[3] - box[1], 1)
        rot = rotation(sh)

        # 图片
        if sh.shape_type == 13:
            try:
                pic = Image.open(BytesIO(sh.image.blob)).convert('RGB')
                iw, ih = pic.size
                cl, cr = sh.crop_left or 0, sh.crop_right or 0
                ct, cb = sh.crop_top or 0, sh.crop_bottom or 0
                pic = pic.crop((int(iw * cl), int(ih * ct), int(iw * (1 - cr)), int(ih * (1 - cb))))
                pic = pic.resize((bw, bh))
                if rot:
                    pic = pic.rotate(-rot, expand=True)
                mask = None
                if geom(sh) in ('ellipse', 'oval'):
                    mask = Image.new('L', pic.size, 0)
                    ImageDraw.Draw(mask).ellipse([0, 0, pic.width - 1, pic.height - 1], fill=255)
                cx0 = box[0] - (pic.width - bw) // 2
                cy0 = box[1] - (pic.height - bh) // 2
                if mask:
                    img.paste(pic, (cx0, cy0), mask)
                else:
                    img.paste(pic, (cx0, cy0))
            except Exception:
                pass
            continue

        # 形状
        if rot:
            pad = int((bw + bh) * 0.5) + 8
            layer = Image.new('RGBA', (bw + pad * 2, bh + pad * 2), (0, 0, 0, 0))
            draw_node(layer, sh, [pad, pad, pad + bw, pad + bh], ppc)
            layer = layer.rotate(-rot, expand=False, resample=Image.BICUBIC)
            img.paste(layer, (box[0] - pad, box[1] - pad), layer)
        else:
            draw_node(img, sh, box, ppc)


def render(path, outdir, width=1400, pages=None):
    prs = Presentation(path)
    ppc = width / (prs.slide_width / EMU)
    Hh = int(prs.slide_height / EMU * ppc)
    os.makedirs(outdir, exist_ok=True)
    outs = []
    for i, s in enumerate(prs.slides, 1):
        if pages and i not in pages:
            continue
        img = Image.new('RGB', (width, Hh), (255, 255, 255))
        try:
            bg = s.background
        except Exception:
            bg = None
        draw(img, s.shapes, ppc)
        p = os.path.join(outdir, f'p{i:02d}.png'); img.save(p); outs.append(p)
    return outs


if __name__ == '__main__':
    f = sys.argv[1]; out = sys.argv[2]
    pages = {int(x) for x in sys.argv[3].split(',')} if len(sys.argv) > 3 else None
    fs = render(f, out, pages=pages)
    print(len(fs), '张 ->', out)
