"""Non-destructive picture styling: preserve embedded pixels, remove frames,
and write native DrawingML soft edges. Does not redraw or process raster assets.
"""
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
from pptx.util import Pt


def soft_edge(picture, radius_pt):
    if picture.shape_type != 13 or not 0 <= radius_pt <= 100:
        raise ValueError('Expected a picture and radius in 0–100pt')
    props=picture._element.spPr
    effects=props.find(qn('a:effectLst'))
    if effects is None:
        effects=OxmlElement('a:effectLst')
        ext=props.find(qn('a:extLst'))
        if ext is None:
            props.append(effects)
        else:
            ext.addprevious(effects)
    for old in list(effects.findall(qn('a:softEdge'))):
        effects.remove(old)
    if radius_pt:
        node=OxmlElement('a:softEdge')
        node.set('rad',str(int(Pt(radius_pt))))
        effects.append(node)
    return picture


def remove_picture_frames(slide, tolerance_emu=5000):
    """Remove only the known thin cream/tan frames around authored pictures.

    Avoid touching content panels or other shapes with similar color.
    """
    removed=[]
    pictures=[p for p in slide.shapes if p.shape_type==13]
    for shape in list(slide.shapes):
        if shape.shape_type!=1 or (shape.has_text_frame and shape.text.strip()):
            continue
        fill=shape._element.find('.//'+qn('a:solidFill')+'/'+qn('a:srgbClr'))
        if fill is None or fill.get('val') not in {'F2E0C4','FFF7EC','F6EBDD'}:
            continue
        for pic in pictures:
            gaps=[pic.left-shape.left,pic.top-shape.top,
                  shape.left+shape.width-pic.left-pic.width,
                  shape.top+shape.height-pic.top-pic.height]
            if all(0<=g<=50000 for g in gaps) and max(gaps)-min(gaps)<=tolerance_emu:
                removed.append(shape.shape_id)
                shape._element.getparent().remove(shape._element)
                break
    return removed
