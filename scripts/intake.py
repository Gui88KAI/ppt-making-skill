#!/usr/bin/env python3
"""Check a structured brief before authoring; this does not interpret free text or confirm facts for the user."""
import argparse
import json
from pathlib import Path

TYPES = {
    "01": ("创赛（互联网+ / 挑战杯 / 大创）", "types/01-创赛互联网+.md", [25, 39]),
    "02": ("商业计划书 BP", "types/02-商业计划书.md", [15, 22]),
    "03": ("路演融资", "types/03-路演融资.md", [10, 12]),
    "04": ("答辩 / 评审汇报", "types/04-答辩汇报.md", [15, 20]),
    "05": ("课程汇报 / 作业 / 主题班会", "types/05-课程汇报.md", [10, 15]),
}


def positive_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and value > 0


def positive_integer(value):
    return isinstance(value, int) and not isinstance(value, bool) and value > 0


def check_pages(value):
    if positive_integer(value):
        return {"mode": "fixed", "value": value}, None
    if not isinstance(value, dict):
        return None, "page_count 需要正整数，或 fixed / maximum / range / delegate 对象"
    mode = value.get("mode")
    if mode == "delegate":
        return {"mode": "delegate"}, None
    if mode in ("fixed", "maximum") and positive_integer(value.get("value")):
        return {"mode": mode, "value": value["value"]}, None
    if mode == "range" and positive_integer(value.get("min")) and positive_integer(value.get("max")) and value["min"] <= value["max"]:
        return {"mode": mode, "min": value["min"], "max": value["max"]}, None
    return None, "页数配置无效：填写正整数；range 下限不能大于上限；delegate 必须来自用户授权"


def assess(brief):
    if not isinstance(brief, dict):
        return {"status": "invalid", "can_author": False, "errors": ["brief 必须是对象"], "missing": [], "questions": []}
    missing, errors, questions = [], [], []
    raw_type = brief.get("type")
    selected = str(raw_type).zfill(2) if raw_type is not None and not isinstance(raw_type, bool) else ""
    if not selected:
        missing.append("type")
        questions.append({"field": "type", "prompt": "你要做哪一类 PPT？", "options": [{"id": k, "label": v[0], "suggested_pages": v[2]} for k, v in TYPES.items()]})
    elif selected not in TYPES:
        errors.append("类型必须是 01–05；不匹配时先请用户说明场景，不能默认创赛")
    for key, question in (
        ("topic", "这套 PPT 的主题是什么？"),
        ("audience", "主要给谁看？"),
        ("occasion", "在什么场合使用？"),
    ):
        if not isinstance(brief.get(key), str) or not brief[key].strip():
            missing.append(key)
            questions.append({"field": key, "prompt": question})
    duration = brief.get("duration_minutes")
    if duration is None:
        missing.append("duration_minutes")
        questions.append({"field": "duration_minutes", "prompt": "计划讲多久？"})
    elif not positive_number(duration):
        errors.append("duration_minutes 必须是大于零的分钟数")
    pages = None
    if brief.get("page_count") is None:
        missing.append("page_count")
        questions.append({"field": "page_count", "prompt": "总页数希望固定多少、最多多少、限定范围，还是按时长交给我安排？总数包含封面、封底和开屏。"})
    else:
        pages, error = check_pages(brief["page_count"])
        if error:
            errors.append(error)
    ready = not missing and not errors
    result = {
        "status": "invalid" if errors else "needs_input" if missing else "ready",
        "can_author": ready, "missing": missing, "errors": errors, "questions": questions,
        "type": selected or None, "page_count": pages,
        "next_action": "ask_missing_decisions" if not ready else "read_selected_type_file",
        "confirmation_boundary": "本脚本只检查结构；字段值是否确由用户给定或授权，须由调用AI依据对话核对",
    }
    if selected in TYPES:
        label, route, suggested = TYPES[selected]
        result.update(type_name=label, read_next=route, suggested_pages=suggested)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("brief", help="已从对话提取的 brief.json；不要猜值凑齐字段")
    args = parser.parse_args()
    result = assess(json.loads(Path(args.brief).read_text(encoding="utf-8-sig")))
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["can_author"] else 1 if result["errors"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
