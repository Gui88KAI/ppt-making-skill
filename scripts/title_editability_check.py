#!/usr/bin/env python3
"""Check named, ungrouped, single-character title textboxes in an actual PPTX."""
import argparse
import json
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def inspect_title(pptx_path, expected, slide_number=1, name_prefix="标题-", only_title_text=False):
    errors = []
    try:
        deck = Presentation(str(pptx_path))
    except Exception as exc:
        return {"status": "failed", "errors": [f"无法读取PPTX: {exc}"], "characters": []}
    if not 1 <= slide_number <= len(deck.slides):
        return {"status": "failed", "errors": ["页码超出PPTX范围"], "characters": []}
    if not name_prefix:
        return {"status": "failed", "errors": ["名称前缀不能为空"], "characters": []}
    expected_chars = [ch for ch in expected if not ch.isspace()]
    if not expected_chars:
        return {"status": "failed", "errors": ["标题不能为空"], "characters": []}

    slide = deck.slides[slide_number - 1]
    named = [(index, shape) for index, shape in enumerate(slide.shapes) if shape.name.startswith(name_prefix)]

    def sort_key(item):
        index, shape = item
        suffix = shape.name[len(name_prefix):]
        number = re.match(r"(\d+)(?:\D|$)", suffix)
        return (0, int(number.group(1)), index) if number else (1, index, index)

    named.sort(key=sort_key)
    rows = []
    for _, shape in named:
        value = shape.text if shape.has_text_frame else ""
        locks = shape._element.xpath(".//a:spLocks")
        locked = any(lock.get("noMove") in ("1", "true") or lock.get("noSelect") in ("1", "true") for lock in locks)
        rows.append({"id": shape.shape_id, "name": shape.name, "text": value,
                     "x_emu": shape.left, "y_emu": shape.top,
                     "type": str(shape.shape_type), "locked": locked})
        if shape.shape_type != MSO_SHAPE_TYPE.TEXT_BOX:
            errors.append(f"{shape.name}: 不是独立原生文本框")
        if len(value) != 1 or value.isspace():
            errors.append(f"{shape.name}: 文本框必须只包含一个非空字符")
        if locked:
            errors.append(f"{shape.name}: 对象被锁定，不能确认可单独移动")
    if len(named) != len(expected_chars):
        errors.append(f"匹配到 {len(named)} 个标题对象，预期 {len(expected_chars)} 个")
    if "".join(row["text"] for row in rows) != "".join(expected_chars):
        errors.append("标题文字或顺序与预期不一致")
    if only_title_text:
        others = [shape.name for shape in slide.shapes
                  if not shape.name.startswith(name_prefix)
                  and shape.has_text_frame and shape.text.strip()]
        if others:
            errors.append("封面存在其他文本框: " + ", ".join(others))
    return {
        "file": str(Path(pptx_path)), "slide": slide_number,
        "status": "failed" if errors else "passed", "errors": errors,
        "characters": rows,
        "limitations": ["仅检查命名对象的文件结构与锁定状态，不验证WPS实际编辑、字体显示或审美"],
    }


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx")
    parser.add_argument("--title", required=True, help="完整标题，空白字符不要求单独文本框")
    parser.add_argument("--slide", type=int, default=1)
    parser.add_argument("--name-prefix", default="标题-")
    parser.add_argument("--only-title-text", action="store_true")
    args = parser.parse_args()
    result = inspect_title(args.pptx, args.title, args.slide, args.name_prefix, args.only_title_text)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
