#!/usr/bin/env python3
"""给 python-pptx 生成的幻灯片补上稳定的对象名和 Morph/淡出转场。"""
from lxml import etree
from pptx.oxml.ns import qn

def transition(slide, kind="morph", speed="med"):
    old = slide._element.find(qn("p:transition"))
    if old is not None:
        slide._element.remove(old)
    node = etree.SubElement(slide._element, qn("p:transition"))
    node.set("spd", speed)
    if kind == "morph":
        etree.SubElement(node, qn("p:morph"), {"option": "byObject"})
    else:
        etree.SubElement(node, qn("p:fade"))

def _set_name(shape, name):
    try:
        shape.name = name
        return
    except Exception:
        pass
    try:
        c_nv_pr = shape._element.nvSpPr.cNvPr
        c_nv_pr.set("name", name)
    except Exception:
        pass

def tag_shapes(slide, spec):
    """按约定给标题、图片、锚点和轨迹命名，供 Morph 跨页匹配。"""
    title = str(spec.get("title", "")).strip()
    anchor = str(spec.get("visual_anchor", "")).strip()
    pictures = []
    text_shapes = []
    for shape in slide.shapes:
        if getattr(shape, "shape_type", None) == 13:
            pictures.append(shape)
        if getattr(shape, "has_text_frame", False) and shape.text.strip():
            text_shapes.append(shape)
    for shape in text_shapes:
        value = shape.text.strip()
        if title and title in value:
            _set_name(shape, "motion:title")
            break
    if pictures:
        _set_name(pictures[0], "motion:hero")
    if anchor:
        for shape in text_shapes:
            if anchor in shape.text:
                _set_name(shape, "motion:anchor")
                break
    lines = [s for s in slide.shapes if getattr(s, "shape_type", None) == 9]
    if lines:
        _set_name(lines[0], "motion:trail")

def apply_slide_motion(slide, spec):
    motion = spec.get("motion") or {}
    transition(slide, motion.get("transition", "morph"), motion.get("speed", "med"))
    tag_shapes(slide, spec)
