#!/usr/bin/env python3
"""fill_template.py — 用户指定模板时，按映射填充一份副本。

编号来自 template_map.py。顶层使用整数，组内使用点分路径，如 "4.2.0"。

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
"""
import argparse
import os
from copy import deepcopy

import yaml
from pptx import Presentation
from template_map import iter_shape_paths
from wps_compat import explicit_fonts, inspect as inspect_wps
from motion import transition

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
    run_properties = deepcopy(src._r.rPr)
    try:
        align = tf.paragraphs[0].alignment
        line_spacing = tf.paragraphs[0].line_spacing
    except Exception:
        align, line_spacing = None, None

    # 清空
    for p in list(tf.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    first = tf.paragraphs[0]
    for r in list(first.runs):
        r._r.getparent().remove(r._r)

    lines = str(new_text).split("\n")
    for index, line in enumerate(lines):
        p = tf.add_paragraph() if index else first
        if align is not None:
            p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        r = p.add_run()
        r.text = line
        if run_properties is not None:
            r._r.insert(0, deepcopy(run_properties))


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


def replace_image(slide, shape, path):
    x, y, w, h = shape.left, shape.top, shape.width, shape.height
    new = slide.shapes.add_picture(path, x, y, w, h)
    try:
        cover_crop(new, w / 360000, h / 360000)
    except Exception:
        pass
    # Keep both z-order and the coordinate system of the original group.
    shape._element.addnext(new._element)
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
    ap.add_argument("--target", choices=["wps", "powerpoint"], default="powerpoint")
    ap.add_argument("--check", action="store_true")
    a = ap.parse_args()

    with open(a.mapping, encoding="utf-8") as stream:
        m = yaml.safe_load(stream) or {}
    prs = Presentation(a.template)
    slides = list(prs.slides)

    done = {"text": 0, "image": 0}
    for item in m.get("slides") or []:
        idx = int(item["index"]) - 1
        if not 0 <= idx < len(slides):
            print(f"⚠ 跳过第 {item['index']} 页：模板只有 {len(slides)} 页")
            continue
        slide = slides[idx]
        shapes = dict(iter_shape_paths(slide.shapes))
        for sidx, txt in (item.get("texts") or {}).items():
            shape = shapes.get(str(sidx))
            if shape is None or not shape.has_text_frame:
                ap.error(f"第 {idx + 1} 页的元素 {sidx} 不是可填充的文本框")
            set_text(shape, txt)
            done["text"] += 1
        for sidx, path in (item.get("images") or {}).items():
            shape = shapes.get(str(sidx))
            if shape is None or not os.path.isfile(path):
                ap.error(f"第 {idx + 1} 页的图片映射 {sidx} 无效：{path}")
            replace_image(slide, shape, path)
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
    if a.target == "wps":
        explicit_fonts(prs)
        for slide in prs.slides:
            transition(slide, "fade")
    prs.save(a.out)
    print(f"完成：替换文字 {done['text']} 处、图片 {done['image']} 张 → {a.out}")
    if a.target == "wps" and a.check:
        import json
        report = inspect_wps(a.out)
        report_path = os.path.splitext(a.out)[0] + ".wps-report.json"
        with open(report_path, "w", encoding="utf-8") as stream:
            json.dump(report, stream, ensure_ascii=False, indent=2)
        print(f"WPS 静态报告: {report_path}；目标 WPS 实机验证尚未执行")
        if report["errors"]:
            raise SystemExit(1)


if __name__ == "__main__":
    main()

