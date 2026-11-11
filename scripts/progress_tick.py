#!/usr/bin/env python3
"""progress_tick.py — PROGRESS.md 的轮次计数（每 10 轮才更新一次）。

用法: 每轮结束时调用一次
    python3 scripts/progress_tick.py

输出:
    "tick 3/10 — 不更新"
    "tick 10/10 — 该更新 PROGRESS.md 了"

约定（写进 SKILL/记忆）：
  - 平时轮次**不要**动 PROGRESS.md
  - 只有计数到 10 的倍数时，才重新生成/补充 PROGRESS.md
  - 计数文件: .progress_round（仓库外也可，放本地 .dsh 下更稳）
"""
import json
import os
import sys

STATE = os.path.expanduser("~/.dsh/progress_round.json")
EVERY = 10


def tick():
    st = {"count": 0, "last_update_round": 0}
    if os.path.exists(STATE):
        try:
            st = json.load(open(STATE))
        except Exception:
            pass
    st["count"] = st.get("count", 0) + 1
    n = st["count"]
    due = n % EVERY == 0
    if due:
        st["last_update_round"] = n
    os.makedirs(os.path.dirname(STATE), exist_ok=True)
    json.dump(st, open(STATE, "w"))
    return n, due


if __name__ == "__main__":
    n, due = tick()
    if due:
        print(f"tick {n}/{EVERY} — 该更新 PROGRESS.md 了")
    else:
        print(f"tick {n}/{EVERY} — 不更新（还剩 {EVERY - n % EVERY} 轮）")
