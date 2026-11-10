#!/usr/bin/env python3
"""fill_template.py — 把内容填进现成模板，保留原有设计（最重要的路线）。

做法：克隆素材库里的漂亮模板，只替换文字和图片，版式/配色/装饰全部不动。

用法:
    # 1. 先看模板结构
    python3 template_map.py 模板.pptx --slide 1

    # 2. 写映射文件（YAML），元素编号来自 template_map
    # slides:
    #   - index: 1
    #     texts: {8: "我的项目名", 9: "汇报人：张三"}
    #     images: {12: "cover.jpg"}
    # delete: [19, 20]        # 删掉多出来的页

    # 3. 填内容
    python3 fill_template.py 模板.pptx mapping.yaml -o 输出.pptx

映射字段说明:
    index   第几页（1 开始，对应 template_map 的页码）
    texts   {元素编号: "新文字"}，保留原元素的字体/字号/颜色/位置
    images  {元素编号: "图片路径"}，按原位置和大小铺满（等比裁切）
    delete  删除这些页（从大到小自动处理）
    keep_text  设为 true 时，没在 texts 里列出的元素保持原样（默认就是 true）
"""
import argparse
import os
import sys

import yaml
from pptx import Presentation
from pptx.util import Cm

MSO_PICTURE = 13


def set_text(shape, new_text):
    """替换文字，保留第一个 run 的字体格式。"""
    tf = shape.text_frame
    # 记录第一段第一个 run 的格式
    src = None
    for p in tf.paragraphs:
        for r in p.runs:
            if r.text.strip():
                src = r
                break
        if src:
            break
    if src is None:
        tf.text = new_text
        return
    font = src.font
    try:
        align = tf.paragraphs[0].alignment
        line_spacing = tf.paragraphs[0].line_spacing
    except Exception:
        align, line_spacing = None, None

    # 清空
    for p in list(tf.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    first = tf.paragraphs[0]
    for r in list(first.runs)[1:]:
        r._r.getparent().remove(r._r)
    first.runs[0].text = "" if not first.runs else first.runs[0].text

    lines = str(new_text).split("\n")
    if not first.runs:
        run = first.add_run()
    else:
        run = first.runs[0]
        run.text = lines[0]
    run.font.size = font.size
    run.font.bold = font.bold
    run.font.name = font.name
    try:
        if font.color and font.color.type is not None:
            run.font.color.rgb = font.color.rgb
    except Exception:
        pass
    for line in lines[1:]:
        p = tf.add_paragraph()
        if align is not None:
            p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        r = p.add_run()
        r.text = line
        r.font.size = font.size
        r.font.bold = font.bold
        r.font.name = font.name
        try:
            if font.color and font.color.type is not None:
                r.font.color.rgb = font.color.rgb
        except Exception:
            pass


def cover_crop(pic, box_w_cm, box_h_cm):
    """按 cover 方式裁切图片，填满原框。"""
    from PIL import Image
    iw, ih = Image.open(pic.image.blob and __import__("io").BytesIO(pic.image.blob)).size
    br, ir = box_w_cm / box_h_cm, iw / ih
    if ir > br:
        over = 1 - br / ir
        pic.crop_left = pic.crop_right = over / 2
    else:
        over = 1 - ir / br
        pic.crop_top = pic.crop_bottom = over / 2


def replace_image(prs, slide, shape, path):
    x, y, w, h = shape.left, shape.top, shape.width, shape.height
    new = slide.shapes.add_picture(path, x, y, w, h)
    try:
        cover_crop(new, w / 360000, h / 360000)
    except Exception:
        pass
    shape._element.getparent().remove(shape._element)
    return new


def delete_slides(prs, idxs):
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    for i in sorted([x - 1 for x in idxs], reverse=True):
        if 0 <= i < len(slides):
            rId = slides[i].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            prs.part.drop_rel(rId)
            xml_slides.remove(slides[i])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("template")
    ap.add_argument("mapping")
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()

    m = yaml.safe_load(open(a.mapping, encoding="utf-8")) or {}
    prs = Presentation(a.template)
    slides = list(prs.slides)

    done = {"text": 0, "image": 0}
    for item in m.get("slides") or []:
        idx = int(item["index"]) - 1
        if idx >= len(slides):
            print(f"⚠ 跳过第 {item['index']} 页：模板只有 {len(slides)} 页")
            continue
        slide = slides[idx]
        shapes = list(slide.shapes)
        for sidx, txt in (item.get("texts") or {}).items():
            sidx = int(sidx)
            if sidx < len(shapes):
                set_text(shapes[sidx], txt)
                done["text"] += 1
        for sidx, path in (item.get("images") or {}).items():
            sidx = int(sidx)
            if sidx < len(shapes) and os.path.exists(path):
                replace_image(prs, slide, shapes[sidx], path)
                done["image"] += 1

    if m.get("delete"):
        delete_slides(prs, [int(x) for x in m["delete"]])
        print("已删除页面:", m["delete"])

    for old, new in (m.get("replace") or {}).items():
        for slide in prs.slides:
            for sh in slide.shapes:
                if sh.has_text_frame and old in sh.text_frame.text:
                    set_text(sh, sh.text_frame.text.replace(old, new))
                    done["text"] += 1

    d = os.path.dirname(os.path.abspath(a.out))
    if d:
        os.makedirs(d, exist_ok=True)
    prs.save(a.out)
    print(f"完成：替换文字 {done['text']} 处、图片 {done['image']} 张 → {a.out}")


if __name__ == "__main__":
    main()
