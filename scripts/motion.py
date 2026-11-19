#!/usr/bin/env python3
"""Conservative WPS transitions; standards-based Morph with a fade fallback."""
from lxml import etree
from pptx.oxml.ns import qn
from wps_compat import motion_kind

MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
P159 = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"


def transition(slide, kind="morph", speed="med"):
    if kind not in ("morph", "fade", "none"):
        raise ValueError(f"Unknown transition: {kind}")
    if speed not in ("slow", "med", "fast"):
        raise ValueError(f"Unknown transition speed: {speed}")
    root = slide._element
    for child in list(root):
        if child.tag == qn("p:transition") or (
            child.tag == f"{{{MC}}}AlternateContent"
            and child.find(".//" + qn("p:transition")) is not None
        ):
            root.remove(child)
    if kind == "none":
        return
    def basic(parent, effect):
        node = etree.SubElement(parent, qn("p:transition"), {"spd": speed})
        etree.SubElement(node, effect)
        return node
    if kind == "morph":
        wrapper = etree.Element(f"{{{MC}}}AlternateContent", nsmap={"mc": MC, "p159": P159})
        choice = etree.SubElement(wrapper, f"{{{MC}}}Choice", {"Requires": "p159"})
        node = basic(choice, f"{{{P159}}}morph")
        node[0].set("option", "byObject")
        basic(etree.SubElement(wrapper, f"{{{MC}}}Fallback"), qn("p:fade"))
    else:
        wrapper = etree.Element(qn("p:transition"), {"spd": speed})
        etree.SubElement(wrapper, qn("p:fade"))
    # Slide order: cSld, clrMapOvr, transition, timing, extLst.
    insert = next((i for i, node in enumerate(root) if node.tag in (qn("p:timing"), qn("p:extLst"))), len(root))
    root.insert(insert, wrapper)


def tag_shapes(slide, spec):
    title = str(spec.get("title", "")).strip()
    anchor = str(spec.get("visual_anchor", "")).strip()
    text_shapes = [s for s in slide.shapes if getattr(s, "has_text_frame", False) and s.text.strip()]
    title_shape = next((s for s in text_shapes if title and title in s.text.strip()), None)
    if title_shape is not None:
        title_shape.name = "!!motion:title"
    pictures = [s for s in slide.shapes if getattr(s, "shape_type", None) == 13]
    if pictures:
        pictures[0].name = "!!motion:hero"
    anchor_shape = next((s for s in text_shapes if anchor and anchor in s.text and s is not title_shape), None)
    if anchor_shape is not None:
        anchor_shape.name = "!!motion:anchor"
    lines = [s for s in slide.shapes if getattr(s, "shape_type", None) == 9]
    if lines:
        lines[0].name = "!!motion:trail"


def apply_slide_motion(slide, spec, target="powerpoint", default=None):
    motion = spec.get("motion") or {}
    kind = motion_kind(spec, target, default)
    transition(slide, kind, motion.get("speed", "med"))
    tag_shapes(slide, spec)
    return kind

