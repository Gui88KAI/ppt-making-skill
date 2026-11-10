#!/usr/bin/env python3
"""gen_image.py — 让 AI 根据 PPT 主题自己生成配图。

素材库里的 PPT 之所以高级，是因为**背景是照片**（机器人、无人机、稻田、工坊），
不是色块。这个脚本让任何能生图的 AI 都能自己造出这些图。

三种模式：
  background  整幅背景（带留白区，给标题让位）
  hero        主体图（产品/场景/概念）
  texture     纹理/抽象图形（光带、线条、粒子）

用法:
  # 单张
  python3 gen_image.py --kind background --theme "智慧农业无人机" --out bg.png

  # 批量（推荐）：按 manifest 生成整套
  python3 gen_image.py --manifest prompts.json --outdir assets/

  # 用其他生图服务
  python3 gen_image.py --manifest prompts.json --provider openai     # 需 OPENAI_API_KEY
  python3 gen_image.py --manifest prompts.json --provider custom --endpoint URL --api-key KEY

支持的 provider:
  pollinations  免密钥（默认），宽高 ≤ 768 最稳
  openai        DALL·E / gpt-image，需 OPENAI_API_KEY
  custom        任意兼容 OpenAI /images/generations 接口的服务
"""

import argparse
import base64
import json
import os
import sys
import time
import urllib.parse
import urllib.request

NEGATIVE = ("no text, no letters, no words, no watermark, no logo, no signature, "
            "no chinese characters, clean composition")

STYLE = {
    "photo": "professional photograph, cinematic lighting, high detail, shallow depth of field",
    "tech": "dark technology aesthetic, glowing light effects, futuristic HUD, sci-fi",
    "business": "clean corporate style, soft lighting, minimal, modern",
    "culture": "traditional Chinese aesthetic, warm tones, artistic, textured",
    "illustration": "flat vector illustration, soft gradients, clean shapes",
}

MAX_SIDE = 768   # pollinations 免密钥稳定上限


# ---------------- provider 实现 ----------------
def gen_pollinations(prompt, w, h, seed=None, model="flux", timeout=180):
    q = urllib.parse.quote(prompt)
    # 超过稳定上限就等比缩小
    scale = min(1.0, MAX_SIDE / max(w, h))
    gw, gh = max(int(w * scale), 64), max(int(h * scale), 64)
    url = f"https://image.pollinations.ai/prompt/{q}?width={gw}&height={gh}&nologo=true&model={model}"
    if seed is not None:
        url += f"&seed={seed}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def gen_openai(prompt, w, h, api_key=None, model="dall-e-3", timeout=300):
    key = api_key or os.environ.get("OPENAI_API_KEY")
    if not key:
        raise RuntimeError("需要 OPENAI_API_KEY")
    # dall-e-3 只支持固定尺寸
    size = "1792x1024" if w / h > 1.4 else ("1024x1792" if h / w > 1.4 else "1024x1024")
    body = json.dumps({"model": model, "prompt": prompt, "size": size,
                       "n": 1, "response_format": "b64_json"}).encode()
    req = urllib.request.Request("https://api.openai.com/v1/images/generations", data=body,
                                 headers={"Authorization": f"Bearer {key}",
                                          "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        data = json.load(r)
    return base64.b64decode(data["data"][0]["b64_json"])


def gen_custom(prompt, w, h, endpoint=None, api_key=None, model=None, timeout=300):
    if not endpoint:
        raise RuntimeError("需要 --endpoint")
    body = json.dumps({"model": model or "image", "prompt": prompt,
                       "size": f"{w}x{h}", "n": 1,
                       "response_format": "b64_json"}).encode()
    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"
    req = urllib.request.Request(endpoint, data=body, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        data = json.load(r)
    item = data["data"][0]
    if "b64_json" in item:
        return base64.b64decode(item["b64_json"])
    if "url" in item:
        with urllib.request.urlopen(item["url"], timeout=timeout) as im:
            return im.read()
    raise RuntimeError("无法解析返回数据")


PROVIDERS = {"pollinations": gen_pollinations, "openai": gen_openai, "custom": gen_custom}


# ---------------- 提示词构造 ----------------
def build_prompt(spec, theme, palette, default_style):
    kind = spec.get("kind", "background")
    parts = []

    if kind == "background":
        parts.append("cinematic full-bleed background image for a presentation slide")
        area = spec.get("text_area", "")
        if area == "left":
            parts.append("the LEFT half is darker and visually empty so white text can be placed there; "
                         "the RIGHT half has the interesting detail")
        elif area == "right":
            parts.append("the RIGHT half is darker and empty for text; the LEFT half has detail")
        elif area == "bottom":
            parts.append("lower third is darker for text, upper two thirds has detail")
        else:
            parts.append("even composition with empty space for text overlay")
        parts.append("wide establishing shot, no centered focal subject")
    elif kind == "hero":
        parts.append("a single clear subject, centered, isolated against a clean dark background, "
                     "product photography lighting")
    else:
        parts.append("abstract decorative graphic, subtle, low contrast, no focal subject")

    parts.append(spec["prompt"])
    if theme:
        parts.append(f"context: {theme}")
    if palette:
        parts.append(f"color palette: {palette}")
    parts.append(STYLE.get(spec.get("style", default_style), STYLE["photo"]))
    parts.append(NEGATIVE)
    return ", ".join(parts)


def generate_one(spec, args, theme, palette):
    w, h = spec.get("size", [1536, 864])
    prompt = build_prompt(spec, theme, palette, args.style)
    fn = PROVIDERS[args.provider]
    print(f"  → {spec.get('key','?')}  {w}x{h}  [{spec.get('kind','background')}]")
    print(f"     {prompt[:120]}...")
    last = ""
    for attempt in range(3):
        try:
            if args.provider == "pollinations":
                data = fn(prompt, w, h, spec.get("seed"), args.model)
            elif args.provider == "openai":
                data = fn(prompt, w, h, args.api_key, args.model)
            else:
                data = fn(prompt, w, h, args.endpoint, args.api_key, args.model)
            if len(data) < 4000:
                raise RuntimeError(f"图片过小 {len(data)}B（可能被限流）")
            return data, prompt, None
        except Exception as e:
            last = str(e)[:100]
            print(f"     重试 {attempt+1}/3: {last}")
            time.sleep(2 + attempt * 2)
    return None, prompt, last


# ---------------- 后处理 ----------------
def postprocess(path, spec, theme_palette_hex=None):
    """统一处理：裁切到目标比例 + 按需加暗遮罩（让文字可读）。"""
    try:
        from PIL import Image, ImageEnhance
    except ImportError:
        return
    im = Image.open(path).convert("RGB")
    tw, th = spec.get("size", [1536, 864])
    # cover 裁切
    target = tw / th
    iw, ih = im.size
    cur = iw / ih
    if cur > target:
        nw = int(ih * target)
        im = im.crop(((iw - nw) // 2, 0, (iw + nw) // 2, ih))
    elif cur < target:
        nh = int(iw / target)
        im = im.crop((0, (ih - nh) // 2, iw, (ih + nh) // 2))
    im = im.resize((min(tw, 1600), int(min(tw, 1600) / target)), Image.LANCZOS)
    # 背景图默认压暗，保证白字可读
    if spec.get("kind") == "background" and spec.get("darken", True):
        factor = spec.get("darken_factor", 0.62)
        im = ImageEnhance.Brightness(im).enhance(factor)
    im.save(path, quality=90)


def main():
    ap = argparse.ArgumentParser(description="AI 生成 PPT 配图")
    ap.add_argument("--kind", default="background", choices=["background", "hero", "texture"])
    ap.add_argument("--prompt", default="")
    ap.add_argument("--theme", default="", help="整套主题，如：智慧农业无人机巡检")
    ap.add_argument("--palette", default="", help="配色描述")
    ap.add_argument("--style", default="photo", choices=list(STYLE))
    ap.add_argument("--text-area", default="", choices=["left", "right", "bottom", ""])
    ap.add_argument("--size", default="1536x864")
    ap.add_argument("--out", default="image.jpg")
    ap.add_argument("--manifest", default="")
    ap.add_argument("--outdir", default="assets")
    ap.add_argument("--provider", default="pollinations", choices=list(PROVIDERS))
    ap.add_argument("--model", default="flux")
    ap.add_argument("--endpoint", default="")
    ap.add_argument("--api-key", default="")
    ap.add_argument("--seed", type=int, default=None)
    ap.add_argument("--no-darken", action="store_true")
    a = ap.parse_args()

    if a.manifest:
        m = json.load(open(a.manifest, encoding="utf-8"))
        theme = m.get("theme", a.theme)
        palette = m.get("palette", a.palette)
        os.makedirs(a.outdir, exist_ok=True)
        report, ok = [], 0
        for spec in m.get("images", []):
            data, prompt, err = generate_one(spec, a, theme, palette)
            rec = dict(key=spec.get("key"), prompt=prompt)
            if data is None:
                rec["status"] = "failed"
                rec["error"] = err
            else:
                path = os.path.join(a.outdir, spec["key"] + ".jpg")
                open(path, "wb").write(data)
                if a.no_darken:
                    spec["darken"] = False
                postprocess(path, spec)
                ok += 1
                rec.update(status="ok", path=path, kb=os.path.getsize(path) // 1024)
                print(f"     ✓ {path}")
            report.append(rec)
        json.dump(report, open(os.path.join(a.outdir, "generation_log.json"), "w"),
                  ensure_ascii=False, indent=1)
        print(f"\n完成 {ok}/{len(m.get('images', []))}，日志 {a.outdir}/generation_log.json")
        if ok < len(m.get("images", [])):
            sys.exit(1)
        return

    if not a.prompt:
        sys.exit("需要 --prompt 或 --manifest")
    w, h = (int(x) for x in a.size.lower().split("x"))
    spec = dict(key="x", kind=a.kind, prompt=a.prompt, size=[w, h], seed=a.seed,
                text_area=a.text_area, style=a.style, darken=not a.no_darken)
    data, prompt, err = generate_one(spec, a, a.theme, a.palette)
    if data is None:
        sys.exit(f"生成失败：{err}")
    d = os.path.dirname(os.path.abspath(a.out))
    if d:
        os.makedirs(d, exist_ok=True)
    open(a.out, "wb").write(data)
    postprocess(a.out, spec)
    print(f"已保存 {a.out}（{os.path.getsize(a.out)//1024} KB）")


if __name__ == "__main__":
    main()
