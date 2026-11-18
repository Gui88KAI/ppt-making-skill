#!/usr/bin/env python3
"""Check an already-extracted brief. This does not interpret conversation or verify facts."""
import argparse
import json
import math
import sys
from pathlib import Path

TYPES = {
    "01": ("创赛（互联网+ / 挑战杯 / 大创）", "directions/01-创赛/GUIDE.md", None),
    "02": ("商业计划书 BP", "directions/02-商业计划书/GUIDE.md", None),
    "03": ("融资路演", "directions/03-融资路演/GUIDE.md", None),
    "04": ("答辩 / 评审汇报", "directions/04-答辩汇报/GUIDE.md", None),
    "05": ("课程汇报 / 作业 / 主题班会", "directions/05-课程与主题班会/GUIDE.md", None),
}


def positive_number(value):
    return (isinstance(value, (int, float)) and not isinstance(value, bool)
            and (not isinstance(value, float) or math.isfinite(value)) and value > 0)


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
    return None, "页数配置无效；delegate 必须来自用户授权"


def has_text(value):
    return isinstance(value, str) and bool(value.strip())


def specialized_questions(brief, selected):
    """Return structural questions only; an AI still needs to read the selected guide."""
    if selected != "01":
        return [], []
    questions, unresolved = [], []
    unknown_values = {"未知", "不知道", "不清楚", "待确认", "unknown", "unsure"}

    def need(field, prompt):
        value = brief.get(field)
        if not has_text(value) and not (field == "competition_year" and positive_integer(value)):
            questions.append({"field": field, "prompt": prompt})
        elif has_text(value) and value.strip().lower() in unknown_values:
            unresolved.append(field)

    need("competition", "具体参加什么赛事？")
    name = brief.get("competition")
    if not has_text(name) or not any(term in name for term in ("挑战杯", "大挑", "小挑")):
        return questions, unresolved
    need("competition_year", "是哪一年或哪一届？不清楚可记录未知并核对通知。")
    need("competition_stage", "当前是校赛、省赛、国赛还是其他阶段？")
    kind = brief.get("challenge_kind")
    if not has_text(kind):
        if "大挑" in name and "小挑" not in name:
            kind = "academic"
        elif "小挑" in name and "大挑" not in name:
            kind = "entrepreneurship"
        else:
            need("challenge_kind", "是课外学术科技作品竞赛、创业计划竞赛，还是专项赛？")
    elif kind.strip().lower() in unknown_values:
        need("challenge_kind", "是课外学术科技作品竞赛、创业计划竞赛，还是专项赛？")
    if has_text(kind) and kind.strip().lower() not in unknown_values:
        if kind.strip().lower() in ("academic", "课外学术科技作品竞赛", "课外学术科技", "大挑"):
            need("project_category", "作品属于自然科学论文、社科调查、科技发明，还是其他类别？")
        elif kind.strip().lower() in ("entrepreneurship", "special", "创业计划竞赛", "创业计划", "小挑", "专项赛", "专项"):
            need("track", "参赛赛道或专项榜题是什么？")
        else:
            questions.append({"field": "challenge_kind", "prompt": "请确认挑战杯分支：课外学术科技、创业计划还是专项赛？"})
    return questions, unresolved


def assess(brief):
    if not isinstance(brief, dict):
        return {"status": "invalid", "route_ready": False, "can_author": False,
                "content_ready": None, "delivery_status": "not_checked",
                "errors": ["brief 必须是对象"], "missing": [], "questions": []}
    missing, errors, questions = [], [], []
    raw_type = brief.get("type")
    selected = str(raw_type).zfill(2) if raw_type is not None and not isinstance(raw_type, bool) else ""
    if not selected:
        missing.append("type")
        questions.append({"field": "type", "prompt": "你要做哪一类 PPT？",
                          "options": [{"id": k, "label": v[0]} for k, v in TYPES.items()]})
    elif selected not in TYPES:
        errors.append("类型必须是 01–05；不明确时先确认用途，不能默认创赛")
    for key, question in (
        ("topic", "这套 PPT 的主题是什么？"),
        ("audience", "主要给谁看？"),
        ("occasion", "在什么场合使用？"),
    ):
        if not has_text(brief.get(key)):
            missing.append(key)
            questions.append({"field": key, "prompt": question})
    duration = brief.get("duration_minutes")
    if duration is None:
        missing.append("duration_minutes")
        questions.append({"field": "duration_minutes", "prompt": "计划讲多久？"})
    elif not positive_number(duration):
        errors.append("duration_minutes 必须是有限且大于零的分钟数")
    pages = None
    if brief.get("page_count") is None:
        missing.append("page_count")
        questions.append({"field": "page_count", "prompt": "总页数固定多少、最多多少、限定范围，还是授权我按时长安排？总数含封面、封底和开屏。"})
    else:
        pages, error = check_pages(brief["page_count"])
        if error:
            errors.append(error)

    route_ready = not missing and not errors
    extra, unresolved = specialized_questions(brief, selected) if selected in TYPES else ([], [])
    # Read the selected guide as soon as common intake is complete. Additional
    # questions prevent claiming structured intake is complete for a challenge.
    ready = route_ready and not extra and not unresolved
    if route_ready:
        questions.extend(extra)
    result = {
        "status": "invalid" if errors else "needs_input" if missing else "needs_specialized_input" if extra else "unresolved" if unresolved else "ready",
        "route_ready": route_ready,
        "can_author": ready,
        "content_ready": None,
        "delivery_status": "not_checked",
        "missing": missing,
        "specialized_missing": [q["field"] for q in extra],
        "specialized_unresolved": unresolved,
        "errors": errors,
        "questions": questions,
        "type": selected or None,
        "page_count": pages,
        "next_action": "ask_missing_decisions" if not route_ready else "read_selected_guide_and_ask_specialized" if extra else "read_selected_guide_and_continue_draft" if unresolved else "read_selected_guide",
        "confirmation_boundary": "route_ready 只表示可读所选方法；can_author 只表示脚本检查的结构性缺项已填写，不证明事实、权限、视觉质量或交付完成。",
    }
    if selected in TYPES:
        label, route, suggested = TYPES[selected]
        result.update(type_name=label, read_next=route, suggested_pages=suggested,
                      method_status="已有历史实践，独立方法待验证" if selected in ("01", "05") else "已有规则，新方法成品待验证")
    return result


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("brief", help="已从对话提取的 brief.json；不要猜值凑齐字段")
    args = parser.parse_args()
    result = assess(json.loads(Path(args.brief).read_text(encoding="utf-8-sig")))
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["can_author"] else 1 if result["errors"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
