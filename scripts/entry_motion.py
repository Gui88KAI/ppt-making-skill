"""Write basic click-controlled OOXML fade groups into a new slide.

Uses PresentationML timing/shape targets. Static checks do not prove WPS playback.
Reference: https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-animation
"""
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn


def click_fades(slide, groups, duration_ms=400):
    if slide._element.find(qn('p:timing')) is not None:
        raise ValueError('Slide already contains timing; preserve or explicitly edit it')
    if not 1 <= len(groups) <= 3 or not 150 <= duration_ms <= 1500:
        raise ValueError('Use 1–3 nonempty groups and a reasonable duration')
    known={s.shape_id for s in slide.shapes}
    flat=[int(sid) for group in groups for sid in group]
    if any(not group for group in groups) or len(flat)!=len(set(flat)) or not set(flat)<=known:
        raise ValueError('Groups must contain distinct existing shape IDs')
    next_id=3
    def uid():
        nonlocal next_id
        value=next_id;next_id+=1
        return value
    steps=[]
    for group in groups:
        outer,inner=uid(),uid()
        rows=[]
        for n,sid in enumerate(group):
            effect,visible,fade=uid(),uid(),uid()
            rows.append(f'''<p:par><p:cTn id="{effect}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="{'clickEffect' if n==0 else 'withEffect'}">
            <p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
            <p:set><p:cBhvr><p:cTn id="{visible}" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
            <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{fade}" dur="{duration_ms}"/><p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl></p:cBhvr></p:animEffect>
            </p:childTnLst></p:cTn></p:par>''')
        steps.append(f'''<p:par><p:cTn id="{outer}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="{inner}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{''.join(rows)}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>''')
    xml=f'''<p:timing {nsdecls('p')}><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>{''.join(steps)}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'''
    timing=parse_xml(xml)
    ext=slide._element.find(qn('p:extLst'))
    if ext is None:
        slide._element.append(timing)
    else:
        ext.addprevious(timing)
    return {'groups':groups, 'effect':'fade', 'duration_ms':duration_ms,
            'file_status':'written', 'native_wps_validation':'not_performed'}
