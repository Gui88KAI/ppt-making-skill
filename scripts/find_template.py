#!/usr/bin/env python3
"""find_template.py — 在素材库索引里检索模板包，避免加载 103GB 整个素材库。

用法:
    python3 find_template.py 蓝色 科技
    python3 find_template.py --cat 红旅非遗 --limit 5
    python3 find_template.py 国赛 --min-pptx 40
    python3 find_template.py --files 2          # 列出第 2 个结果的 pptx 清单
    python3 find_template.py 乡村振兴 --json    # 机器可读输出

索引: ../assets/模板索引.json（由素材库 zip 生成，记录了 468 个模板包 / 2697 个 pptx）
素材库: /sdcard/Download/DSH-Go/挑战杯PPT制作素材库.zip

典型流程: 检索 → 只解压命中的那一个包的预览图/字体 → 再解压需要的那一个 pptx。
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
INDEX = os.path.join(HERE, "..", "assets", "模板索引.json")
ARCHIVE = os.environ.get("PPT_LIBRARY_ZIP",
                         "/sdcard/Download/DSH-Go/挑战杯PPT制作素材库.zip")


def load():
    if not os.path.exists(INDEX):
        sys.exit(f"找不到索引 {INDEX}，请先用素材库生成")
    return json.load(open(INDEX, encoding="utf-8"))


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return
    want_files = "--files" in args
    as_json = "--json" in args
    limit = 10
    cat = None
    min_pptx = 0
    words = []
    i = 0
    while i < len(args):
        a = args[i]
        if a == "--cat":
            i += 1
            cat = args[i]
        elif a == "--limit":
            i += 1
            limit = int(args[i])
        elif a == "--min-pptx":
            i += 1
            min_pptx = int(args[i])
        elif a.startswith("--"):
            pass
        else:
            words.append(a)
        i += 1

    data = load()
    packs = data["packs"]
    if cat:
        packs = [p for p in packs if p["cat"] == cat]
    if words:
        packs = [p for p in packs if all(w.lower() in p["path"].lower() for w in words)]
    packs = [p for p in packs if p["pptx"] >= min_pptx]
    packs.sort(key=lambda p: -p["mb"])
    hits = packs[:limit]

    if as_json:
        json.dump(hits, sys.stdout, ensure_ascii=False, indent=1)
        return

    print(f"命中 {len(packs)} 个包，显示前 {len(hits)} 个（素材库 {data['total_mb']/1024:.1f} GB / {data['pptx_count']} 个 pptx）\n")
    for n, p in enumerate(hits, 1):
        print(f"{n}. [{p['cat']}] {p['name']}")
        print(f"   pptx {p['pptx']} 个, 包 {p['mb']/1024:.1f} GB, 预览图 {p['images']} 张, 字体 {p['fonts']} 个")
        print(f"   {p['path']}")

    if want_files:
        try:
            idx = int(input("\n列出第几个包的 pptx？(回车跳过) ") or 0)
        except Exception:
            idx = 0
        if 1 <= idx <= len(hits):
            p = hits[idx - 1]
            print()
            for f in data["pptx"]:
                if f["p"].startswith(p["path"] + "/"):
                    print(f"   {f['mb']:7.1f} MB  {f['p']}")
            print(f"\n解压命令:\n  unzip \"{ARCHIVE}\" \"{p['path']}/*\" -d /tmp/ppt")
            print(f"只要预览图:\n  unzip \"{ARCHIVE}\" \"{p['path']}/*预览*\" -d /tmp/ppt")

    print(f"\n下一步（只解压命中的包，不要全量解压）:\n  unzip \"{ARCHIVE}\" \"{hits[0]['path']}/*\" -d /tmp/ppt")


if __name__ == "__main__":
    main()
