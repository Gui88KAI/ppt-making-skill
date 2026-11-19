#!/usr/bin/env python3
"""只检索已有轻量索引；云端素材只读，不连接、不下载、不解压。"""
import argparse
import json
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

def read_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8'))

def search(data, words=(), category=None, min_pptx=0, limit=10):
    return sorted([
        p for p in data.get('packs', [])
        if (not category or p.get('cat') == category)
        and all(w.casefold() in str(p.get('path', '')).casefold() for w in words)
        and p.get('pptx', 0) >= min_pptx
    ], key=lambda p: -p.get('mb', 0))[:limit]

def source_config(path=None):
    local = ROOT/'library'/'sources.local.json'
    selected = Path(path or os.environ.get('PPT_LIBRARY_CONFIG') or
                    (local if local.is_file() else ROOT/'library'/'sources.json'))
    return read_json(selected)

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('words', nargs='*')
    ap.add_argument('--cat')
    ap.add_argument('--limit', type=int, default=10)
    ap.add_argument('--min-pptx', type=int, default=0)
    ap.add_argument('--files', nargs='?', const=1, type=int, metavar='N',
                    help='非交互列出第N个命中包的历史PPTX路径')
    ap.add_argument('--json', action='store_true')
    ap.add_argument('--index', default=os.environ.get('PPT_LIBRARY_INDEX') or str(ROOT/'assets'/'模板索引.json'))
    ap.add_argument('--source-config')
    args = ap.parse_args()
    if args.limit < 1 or args.min_pptx < 0:
        ap.error('--limit必须为正数；--min-pptx不能为负数')
    data = read_json(args.index)
    cfg = source_config(args.source_config)
    hits = search(data, args.words, args.cat, args.min_pptx, args.limit)
    selected_files = []
    if args.files is not None:
        if not 1 <= args.files <= len(hits):
            ap.error('--files N必须指向实际命中结果；无结果时不读取文件清单')
        prefix = hits[args.files-1]['path'].rstrip('/')+'/'
        selected_files = [f for f in data.get('pptx', []) if f['p'].startswith(prefix)]
    result = {
        'source_provider': cfg.get('provider', 'google_drive'),
        'access_mode': 'read_only_online',
        'index_status': cfg.get('index', {}).get('status', 'unverified'),
        'folder_url': cfg.get('folder_url'),
        'folder_id': cfg.get('folder_id'),
        'hits': hits, 'files': selected_files,
        'next_step': '用已连接的云端只读工具搜索并在线预览；不得下载、导出、同步或缓存。目录未定位时先只读搜索或补问链接/ID。',
        'cloud_access_performed': False,
    }
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return
    print(f'历史索引命中 {len(hits)} 个包；状态 {result["index_status"]}，不代表当前云端文件已核实。')
    for n, pack in enumerate(hits, 1):
        print(f'{n}. [{pack.get("cat", "")}] {pack.get("name", "")}')
        print(f'   参考路径：{pack["path"]}')
    for f in selected_files:
        print(f'   PPTX线索：{f["p"]}')
    print(result['next_step'])

if __name__ == '__main__':
    main()
