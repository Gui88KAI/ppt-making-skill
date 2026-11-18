#!/usr/bin/env python3
"""封面示例入口。参数与 scripts/make_cover.py 一致。"""
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
from make_cover import main

if __name__ == "__main__":
    main()
