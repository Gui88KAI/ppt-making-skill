#!/usr/bin/env python3
"""blend.py — 让人物融进背景的五层处理。

1. 色彩匹配：把人物的色温/亮度往背景靠
2. 边缘光：给人像加暖色轮廓光（模拟环境光）
3. 地面阴影：脚下加柔和投影
4. 空气雾：人物与背景之间叠一层雾，削弱"贴上去"的边界
5. 统一后处理：全局调色 + 暗角 + 胶片颗粒
"""
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter
import numpy as np


def color_transfer(src, ref, strength=0.6):
    """把 src 的均值/方差往 ref 靠（Reinhard 简化版）。"""
    s = np.asarray(src.convert("RGB")).astype(np.float32)
    r = np.asarray(ref.convert("RGB")).astype(np.float32)
    for c in range(3):
        s_mean, s_std = s[:, :, c].mean(), s[:, :, c].std()
        r_mean, r_std = r[:, :, c].mean(), r[:, :, c].std()
        if s_std < 1e-3:
            continue
        s[:, :, c] = (s[:, :, c] - s_mean) / s_std * r_std + r_mean
    out = src.convert("RGB").copy()
    res = Image.fromarray(np.clip(s, 0, 255).astype(np.uint8))
    return Image.blend(src.convert("RGB"), res, strength)


def rim_light(hero, color=(255, 210, 140), width=14, strength=0.55):
    """给人像边缘加一圈光，模拟环境光反射。"""
    alpha = hero.split()[3]
    # 边缘 = 膨胀 - 原图
    dilated = alpha.filter(ImageFilter.MaxFilter(width * 2 + 1))
    edge = Image.fromarray(
        np.clip(np.asarray(dilated, dtype=np.int16) - np.asarray(alpha, dtype=np.int16), 0, 255).astype(np.uint8))
    edge = edge.filter(ImageFilter.GaussianBlur(width * 0.6))
    rim = Image.new("RGBA", hero.size, color + (0,))
    rim.putalpha(edge.point(lambda v: int(v * strength)))
    return Image.alpha_composite(hero, rim)


def soft_shadow(size, pos, radius=180, alpha=90):
    """脚下柔和投影。"""
    w, h = size
    sh = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(sh)
    x, y = pos
    d.ellipse([x - radius, y - radius * 0.25, x + radius, y + radius * 0.25],
              fill=(0, 0, 0, alpha))
    return sh.filter(ImageFilter.GaussianBlur(radius * 0.35))


def mist_overlay(size, y_from=0.55, color=(190, 210, 200), alpha=42):
    """下半部分加薄雾，拉开前后层次。"""
    w, h = size
    mist = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(mist)
    for y in range(int(h * y_from), h):
        t = (y - h * y_from) / (h * (1 - y_from))
        a = int(alpha * t)
        d.line([(0, y), (w, y)], fill=color + (a,))
    return mist.filter(ImageFilter.GaussianBlur(20))


def unify(img, vignette=0.55, grain=6):
    """全局统一：暗角 + 颗粒，让不同素材看起来像同一张照片。"""
    w, h = img.size
    # 暗角
    vig = Image.new("L", (w, h), 0)
    ImageDraw.Draw(vig).ellipse([-w * .2, -h * .3, w * 1.2, h * 1.3], fill=255)
    vig = vig.filter(ImageFilter.GaussianBlur(w // 10))
    dark = Image.new("RGB", (w, h), (4, 12, 10))
    img = Image.composite(img, dark, vig.point(lambda v: int(255 - (255 - v) * (1 - vignette))))
    # 颗粒
    if grain > 0:
        arr = np.asarray(img).astype(np.int16)
        noise = np.random.normal(0, grain, arr.shape[:2])[:, :, None]
        arr = np.clip(arr + noise, 0, 255).astype(np.uint8)
        img = Image.fromarray(arr)
    return img


def blend_hero(bg, hero, x, y, scale=1.0, feather=40, match=0.55, rim=True, shadow=True, mist=True):
    """完整融合流程。bg 为 RGB，hero 为 RGBA（已抠图）。"""
    bg = bg.convert("RGB")
    hero = hero.convert("RGBA")
    if scale != 1.0:
        hero = hero.resize((int(hero.width * scale), int(hero.height * scale)), Image.LANCZOS)

    # 羽化 alpha（保留原有 alpha 的软边）
    a = hero.split()[3].filter(ImageFilter.GaussianBlur(feather))
    hero.putalpha(a)

    # 取人物将占据的背景区域做参考
    px, py = int(x), int(y)
    pw, ph = hero.size
    ref_box = (max(px, 0), max(py, 0), min(px + pw, bg.width), min(py + ph, bg.height))
    ref = bg.crop(ref_box) if ref_box[2] > ref_box[0] and ref_box[3] > ref_box[1] else bg

    # 1) 色彩匹配
    hero_rgb = color_transfer(hero.convert("RGB"), ref, match).convert("RGBA")
    hero_rgb.putalpha(a)

    # 2) 边缘光
    if rim:
        hero_rgb = rim_light(hero_rgb)

    bg = bg.convert("RGBA")

    # 3) 阴影
    if shadow:
        sh = soft_shadow(bg.size, (px + pw // 2, py + ph - int(ph * 0.02)))
        bg = Image.alpha_composite(bg, sh)

    # 4) 合成
    bg.alpha_composite(hero_rgb, (px, py))

    # 5) 雾
    if mist:
        bg = Image.alpha_composite(bg, mist_overlay(bg.size))

    return bg.convert("RGB")
