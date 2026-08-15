# -*- coding: utf-8 -*-
"""调休轮转状态持久化。

状态文件（默认 项目根/状态/rest_state.json）记录每人「上次调休日」，
跨多次运行保持轮转连续性：每运行读取 -> 按『距上次休息间隔最大者优先』排 ~3 人/日
（周二~周日轮转、周一不排、排除当日请假/补休者）-> 回写。

首次运行从最新历史文件的『调休』备注播种；之后每次运行累积更新。
"""
import os
import json
import datetime

from io_roster import find_project_root  # noqa: E402

_HERE = os.path.dirname(os.path.abspath(__file__))


def default_state_path():
    root = find_project_root(_HERE)
    return os.path.join(root, "状态", "rest_state.json")


# 离岗四分类：任何一类离岗都视为「当日离岗」，计入上次休息日基线。
# 只读「调休」会漏掉补休/年休假者（如震生 8.4-8.8 补休年休假），
# 导致其 last_rest 为空、被当作「从未休过」而误排。
OFF_KEYS = ["rest", "leave", "buxiu", "sick", "annual"]


def _seed_from_history(history):
    """从历史事件记录播种 last_rest / rest_history。
    聚合 调休/请假/补休/病假/年休假 四类离岗，消除冷启动盲区。"""
    last_rest = {}
    rest_history = {}
    for h in history:
        hd = datetime.date.fromisoformat(h["date"])
        for key in OFF_KEYS:
            for n in h.get(key, []):
                rest_history.setdefault(n, [])
                if hd.isoformat() not in rest_history[n]:
                    rest_history[n].append(hd.isoformat())
                if n not in last_rest or hd > datetime.date.fromisoformat(last_rest[n]):
                    last_rest[n] = hd.isoformat()
    return {"version": 1, "last_rest": last_rest, "rest_history": rest_history}


def load_or_seed(state_path=None, history=None):
    if state_path is None:
        state_path = default_state_path()
    if os.path.exists(state_path):
        with open(state_path, encoding="utf-8") as f:
            state = json.load(f)
        state.setdefault("last_rest", {})
        state.setdefault("rest_history", {})
        return state
    if history is None:
        raise RuntimeError("无状态文件且无历史记录，无法播种调休状态")
    state = _seed_from_history(history)
    save(state, state_path)
    return state


def save(state, state_path=None):
    if state_path is None:
        state_path = default_state_path()
    d = os.path.dirname(state_path)
    if d and not os.path.exists(d):
        os.makedirs(d, exist_ok=True)
    state["updated_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(state_path, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def plan_rest(state, names, future_dates, unavail_by_day, forced=None, daily_target=3, tags_of=None):
    """按间隔最大优先排调休，更新 state 的 last_rest / rest_history，返回 rest_by_day。

    unavail_by_day: {iso: set(names)} 当日请假/补休/病假/年休假者（不参与轮休候选）。
    forced: {iso: set(names)} 用户强制调休（优先列入，不计入 daily_target 上限之外）。
    tags_of: {name: set(tags)} 人员标签，用于关键标签保护（见下）。

    关键标签保护：负责人 / 技术 是稀缺专项岗（各仅 3 人），且岗位要求每天需 2 名
    不同的负责人（行政上午+下午）与 2 名不同的技术（音响上午+下午）。为兼容
    「每人一天不跨段」的硬约束，必须保证同一天「轮休 + 请假/补休」合计
    不抽空任一关键标签（即每标签当天不可用者 ≤1，保留 ≥2 可用）。否则该天
    行政/音响必有一时段无人可填而阻断。
    """
    last = state["last_rest"]
    hist = state.setdefault("rest_history", {})
    forced = forced or {}
    tags_of = tags_of or {}
    CRIT = {"负责人", "技术"}
    rest_by_day = {}
    for fd in future_dates:
        d = fd["date"]
        iso = d.isoformat()
        un = unavail_by_day.get(iso, set())
        # 当天已请假/补休占用的关键标签数（用户强制调休也计入占用）
        crit_used = {t: 0 for t in CRIT}
        for n in un:
            for t in CRIT:
                if t in tags_of.get(n, ()):
                    crit_used[t] += 1
        cands = [n for n in names if n not in un]
        # 间隔越大越优先；新人/久未休 = 999（最大值）优先
        def _gap(n):
            ln = last.get(n)
            if not ln:
                return 999
            if isinstance(ln, datetime.date):
                ln = ln.isoformat()
            g = (d - datetime.date.fromisoformat(ln)).days
            # 状态被污染写入未来日期（重跑同目标日）-> 间隔为负；
            # 视作「刚休完」，钳为 0（最低优先级），杜绝负间隔被反复选中
            return g if g >= 0 else 0
        ranked = sorted(cands, key=lambda n: (-_gap(n), n))
        chosen = []
        fset = forced.get(iso, set())

        def _admit(n):
            """n 是否可排入（关键标签占用检查，强制者除外）。"""
            if n in fset:
                return True
            for t in CRIT:
                if t in tags_of.get(n, ()) and crit_used[t] >= 1:
                    return False
            return True

        def _commit(n):
            chosen.append(n)
            last[n] = d.isoformat()
            hist.setdefault(n, [])
            if iso not in hist[n]:
                hist[n].append(iso)
            for t in CRIT:
                if t in tags_of.get(n, ()):
                    crit_used[t] += 1

        for n in fset:
            if n in cands and n not in chosen:
                _commit(n)
        for n in ranked:
            if len(chosen) >= daily_target:
                break
            if n in chosen:
                continue
            if not _admit(n):
                continue
            _commit(n)
        rest_by_day[iso] = chosen
    return rest_by_day


if __name__ == "__main__":
    from io_roster import parse_personnel, parse_history
    root = find_project_root(_HERE)
    people = parse_personnel(os.path.join(root, "输入", "人员", "人员.xls"))
    hist = parse_history(os.path.join(root, "输入", "过往三天"), people)
    st = load_or_seed(default_state_path(), hist)
    print("SEEDED last_rest:", st["last_rest"])
    fd = []
    d = datetime.date(2026, 8, 13)
    for _ in range(3):
        d += datetime.timedelta(days=1)
        if d.weekday() != 0:
            fd.append({"date": d, "weekday": "x"})
    rb = plan_rest(st, [p["name"] for p in people], fd, {})
    print("PLANNED rest_by_day:", rb)
