#!/usr/bin/env python3
"""Check actual embedded pictures for identical content, including renamed/re-encoded copies.

This does not detect every crop, recolor, mirror or visually similar composition;
those still require visual review. No network or cloud downloads are performed.
"""
import argparse
import hashlib
import json
from collections import defaultdict
from io import BytesIO
from pathlib import Path
from PIL import Image, ImageOps
from pptx import Presentation


def pixel_hash(blob):
    with Image.open(BytesIO(blob)) as raw:
        im = ImageOps.exif_transpose(raw).convert('RGBA')
        return hashlib.sha256(str(im.size).encode()+im.tobytes()).hexdigest()


def inspect(path, excluded_prefixes=(), baseline=None):
    seen = defaultdict(list)
    old = set()
    if baseline:
        for f in Path(baseline).rglob('*'):
            if f.is_file() and f.suffix.lower() in ('.png','.jpg','.jpeg','.webp'):
                old.add(pixel_hash(f.read_bytes()))
    baseline_matches = []
    excluded = []
    def walk(shapes):
        for shape in shapes:
            yield shape
            if shape.shape_type == 6:
                yield from walk(shape.shapes)
    for page, slide in enumerate(Presentation(path).slides, 1):
        for shape in walk(slide.shapes):
            if shape.shape_type != 13:
                continue
            row = {'page':page, 'name':shape.name, 'shape_id':shape.shape_id}
            if any(shape.name.startswith(p) for p in excluded_prefixes):
                excluded.append(row)
                continue
            digest = pixel_hash(shape.image.blob)
            seen[digest].append(row)
            if digest in old:
                baseline_matches.append(row)
    duplicates = [rows for rows in seen.values() if len(rows)>1]
    return {'checked_pictures':sum(map(len,seen.values())),
            'unique_picture_contents':len(seen), 'duplicates':duplicates,
            'baseline_matches':baseline_matches, 'excluded':excluded,
            'status':'passed' if not duplicates and not baseline_matches else 'failed',
            'limitation':'完全相同像素检测；裁切、镜像、改色和相似构图仍须逐图目视核对。'}


if __name__ == '__main__':
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('pptx')
    ap.add_argument('--baseline', help='若要求全部新图，可指定旧配图目录')
    ap.add_argument('--exclude-name-prefix', action='append', default=[],
                    help='仅用于已记录理由的标识等结构素材，不可排除重复主题图')
    ap.add_argument('--report')
    args=ap.parse_args()
    result=inspect(args.pptx,args.exclude_name_prefix,args.baseline)
    body=json.dumps(result,ensure_ascii=False,indent=2)
    if args.report:
        Path(args.report).write_text(body,encoding='utf-8')
    print(body)
    raise SystemExit(0 if result['status']=='passed' else 1)
