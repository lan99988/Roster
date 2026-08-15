# -*- coding: utf-8 -*-
"""排班求解器：按 11 条权威规则做岗位分配、负荷均衡、自动轮转调休，
并在硬约束无法满足时返回 3 套解决方案（优缺点 + 沟通对象 + 话术）。

纯逻辑、确定性强（无随机），便于单元测试与低 token 复现。
"""
import os
import json
import datetime

_HERE = os.path.dirname(os.path.abspath(__file__))
from io_roster import parse_personnel, parse_history, parse_buxiu, find_project_root  # noqa: E402

HOST_KEYS = {"演艺上午", "演艺下午", "开衙", "讲解11:30", "讲解16:30"}
WEEKDAY_CN = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]


def build_groups(personnel):
    tags_of = {p["name"]: set(p["tags"]) for p in personnel}
    by_tag = {}
    for p in personnel:
        for t in p["tags"]:
            by_tag.setdefault(t, []).append(p["name"])
    women = by_tag.get("女", [])
    hosts = by_tag.get("主持人", [])
    tech = by_tag.get("技术", [])
    leads = by_tag.get("负责人", [])
    pr = by_tag.get("宣传", [])
    all_names = [p["name"] for p in personnel]
    special = set(women) | set(hosts) | set(tech) | set(leads) | set(pr)
    general = [n for n in all_names if n not in special]
    return {
        "tags_of": tags_of, "by_tag": by_tag, "women": women, "hosts": hosts,
        "tech": tech, "leads": leads, "pr": pr, "general": general, "all": all_names,
    }


def _pool(groups, tags):
    if not tags:
        return list(groups["all"])
    pools = [set(groups["by_tag"].get(t, [])) for t in tags]
    base = pools[0]
    for s in pools[1:]:
        base = base & s
    return list(base)


def compute_future_dates(template_date, n=3):
    """模板日期之后的连续 n 个工作日（周二~周日，跳过周一）。"""
    d = template_date + datetime.timedelta(days=1)
    out = []
    while len(out) < n:
        if d.weekday() != 0:  # 0 = Monday
            out.append({"date": d, "weekday": WEEKDAY_CN[d.weekday()]})
        d += datetime.timedelta(days=1)
    return out


def _merge_unavailable(changes, buxiu, rest_by_day, future_dates):
    """返回 {iso: {'full':set,'am':set,'pm':set}}。"""
    un = {}
    for fd in future_dates:
        iso = fd["date"].isoformat()
        un[iso] = {"full": set(), "am": set(), "pm": set()}
    for u in changes.get("unavailable", []):
        iso = u["date"]
        if iso not in un:
            continue
        scope = (u.get("scope") or "全天")
        if scope in ("上午", "am", "morning", "早"):
            un[iso]["am"].add(u["name"])
        elif scope in ("下午", "pm", "afternoon", "晚"):
            un[iso]["pm"].add(u["name"])
        else:  # 全天 / day / 全 / 请假 / 病假 / 年休假 / 补休 等均按全天
            un[iso]["full"].add(u["name"])
    # 补休：按日期
    for name, days in buxiu.items():
        for fd in future_dates:
            if fd["date"].day in days:
                un[fd["date"].isoformat()]["full"].add(name)
    # 自动/强制调休
    for iso, names in rest_by_day.items():
        if iso in un:
            un[iso]["full"].update(names)
    return un


def compute_rest(personnel_names, history, future_dates, changes):
    """按历史休息间隔均衡轮转，每天约 3 人调休（2-4 浮动由缺口自然形成）。"""
    last_rest = {}
    for h in history:
        hd = datetime.date.fromisoformat(h["date"])
        for n in h["rest"]:
            if n not in last_rest or hd > last_rest[n]:
                last_rest[n] = hd
    forced = {}
    for r in changes.get("rest", []):
        forced.setdefault(r["date"], set()).add(r["name"])
    rest_by_day = {}
    for fd in future_dates:
        d = fd["date"]
        cands = []
        for n in personnel_names:
            lr = last_rest.get(n)
            deficit = (d - lr).days if lr else 999
            cands.append((-deficit, n))
        cands.sort()
        chosen = []
        for _, n in cands:
            if n in forced.get(d.isoformat(), set()):
                if n not in chosen:
                    chosen.append(n)
                    last_rest[n] = d
                continue
            if len(chosen) >= 3:
                break
            chosen.append(n)
            last_rest[n] = d
        # 用户强制的也计入（可能超过3，属人工干预）
        for n in forced.get(d.isoformat(), set()):
            if n not in chosen:
                chosen.append(n)
        rest_by_day[d.isoformat()] = chosen
    return rest_by_day


def _available(groups, un_day):
    full = un_day["full"]
    return [n for n in groups["all"] if n not in full]


def schedule_day(demands, groups, un_day, load):
    result = {}
    usage = {n: 0 for n in groups["all"]}
    host_usage = {n: 0 for n in groups["all"]}
    # 段互斥：每人当天只能落在上午段(B/C)或下午段(D)之一，禁止跨段
    am_assigned = set()
    pm_assigned = set()
    hard = []
    soft = []
    # 主持人段预算：am 段 distinct 主持人上限 = 可用主持人 - 1（预留 1 人给 pm 段）。
    # 解决"主持人/技术双紧缺"时单段被抢空导致的死锁（如 8.16 休 2 主持人、8.18 映纯休后技术紧）。
    avail_hosts = [h for h in groups["hosts"] if h not in un_day["full"]]
    pm_host_needed = any(dd.get("block") == "pm" and dd["key"] in HOST_KEYS for dd in demands)
    am_host_budget = max(1, len(avail_hosts) - (1 if pm_host_needed else 0))
    am_host_distinct = set()

    def block_unavail(block):
        if block == "am":
            return un_day["am"]
        if block == "pm":
            return un_day["pm"]
        return set()

    def pick(d, pool, allow_stack, need_override=None):
        need = d["count"] if need_override is None else need_override
        chosen = []
        bu = block_unavail(d["block"])
        # 跨段排除：上午段岗跳过已排下午段的人，下午段岗跳过已排上午段的人
        seg_blocked = pm_assigned if d["block"] == "am" else am_assigned
        is_am_host = allow_stack and d["block"] == "am"
        if allow_stack:
            cands = sorted([n for n in pool
                            if n not in un_day["full"] and n not in bu and n not in seg_blocked],
                           key=lambda n: ((0 if (is_am_host and n in am_host_distinct) else 1),
                                          host_usage[n], usage[n], load.get(n, 0), n))
        else:
            cands = sorted([n for n in pool
                            if n not in un_day["full"] and n not in bu and n not in seg_blocked],
                           key=lambda n: (usage[n], load.get(n, 0), n))
        for n in cands:
            if len(chosen) >= need:
                break
            # am 段主持人预算：新 distinct 人超出预算则跳过（优先复用已在 am 段的人）
            if is_am_host and n not in am_host_distinct and len(am_host_distinct) >= am_host_budget:
                continue
            cap = 3 if allow_stack else 2
            if usage[n] >= cap:
                continue
            chosen.append(n)
            usage[n] += 1
            if allow_stack:
                host_usage[n] += 1
            if is_am_host:
                am_host_distinct.add(n)
            if d["block"] == "am":
                am_assigned.add(n)
            else:
                pm_assigned.add(n)
        return chosen

    for d in demands:
        key = d["key"]
        if d.get("prefer_general"):
            pool = list(groups["general"])
        else:
            pool = _pool(groups, d["tags"])
        chosen = []
        # 打包岗：复用同段配对岗已排且满足本岗标签的人（行政=音响同段同人）
        pair_key = d.get("pair_with")
        if pair_key and pair_key in result:
            for n in result[pair_key]:
                if n in pool and n not in chosen:
                    bu = block_unavail(d["block"])
                    if n in un_day["full"] or n in bu:
                        continue
                    seg_blocked = pm_assigned if d["block"] == "am" else am_assigned
                    if n in seg_blocked:
                        continue
                    cap = 3 if (key in HOST_KEYS) else 2
                    if usage[n] >= cap:
                        continue
                    chosen.append(n)
                    usage[n] += 1
                    if key in HOST_KEYS:
                        host_usage[n] += 1
                    if d["block"] == "am":
                        am_assigned.add(n)
                    else:
                        pm_assigned.add(n)
        # 剩余需求走常规 pick
        if len(chosen) < d["count"]:
            rem = d["count"] - len(chosen)
            chosen += pick(d, pool, allow_stack=(key in HOST_KEYS), need_override=rem)
        if len(chosen) < d["count"] and d.get("fallback") is not None:
            rem = d["count"] - len(chosen)
            fallback_pool = _pool(groups, d["fallback"])
            extra = pick(d, fallback_pool, allow_stack=(key in HOST_KEYS), need_override=rem)
            chosen += extra[:rem]
        result[key] = chosen
        if len(chosen) < d["count"]:
            if d.get("optional"):
                soft.append(f"{d['match']} 空缺（{d['count'] - len(chosen)} 人未排，单人专属岗且当日不可用，按历史惯例留空并提示）")
            else:
                hard.append(f"{d['match']} 缺 {d['count'] - len(chosen)} 人（需求 {d['count']}，已排 {len(chosen)}）")
    # 段互斥校验（防御性）：检查是否有人同时落在上午段(am)与下午段(pm)
    am_keys = {d["key"] for d in demands if d.get("block") == "am"}
    pm_keys = {d["key"] for d in demands if d.get("block") == "pm"}
    cross = []
    for n in groups["all"]:
        in_am = any(n in result.get(k, []) for k in am_keys)
        in_pm = any(n in result.get(k, []) for k in pm_keys)
        if in_am and in_pm:
            cross.append(n)
    if cross:
        hard.append("段互斥违反（同人跨上午/下午排班）：" + "、".join(cross))

    # 主持人 distinct 检查
    host_names = set()
    for k in HOST_KEYS:
        host_names.update(result.get(k, []))
    if len(host_names) < 2:
        hard.append(f"主持人当天同时在场仅 {len(host_names)} 人，低于最少 2 人要求")
    return result, usage, hard, soft


def _build_options(issues, un, future_dates):
    """根据失败原因生成 3 套方案。"""
    women_short = any("女员工" in i for i in issues)
    host_short = any("主持人" in i for i in issues)
    tech_short = any("技术岗" in i or "正堂音响" in i for i in issues)
    lead_short = any("负责人" in i for i in issues)
    gap = "; ".join(issues)
    specifics = []
    if women_short:
        specifics.append("女员工缺口")
    if host_short:
        specifics.append("主持人缺口")
    if tech_short:
        specifics.append("技术/正堂音响缺口")
    if lead_short:
        specifics.append("负责人缺口")

    opt1_title = "方案一：补齐人员或调整不可用安排（推荐）"
    opt1_pros = "能严格满足全部硬约束，产出可直接发布，公平性最好。"
    opt1_cons = "需要协调相关人员在指定日期到岗或改期，沟通成本略高。"
    opt1_comm = "联系场馆负责人 / 排班维护人，确认缺口人员（" + "、".join(specifics) + "）能否调班、借调或改期。"
    opt1_word = (f"自动排班暂未发布：检测到「{gap}」。请确认以下人员能否在对应日期到岗或调换："
                 + "、".join(specifics) + "。补齐后无需改配置，直接重新运行即可生成可发布排班。")

    opt2_title = "方案二：经负责人批准后放宽硬约束"
    opt2_pros = "可立即出表，避免停摆；在人力实在不足时保运转。"
    opt2_cons = "偏离既定规则（如驿站须女、主持人最少2人），需负责人书面确认，风险自担。"
    opt2_comm = "向场馆负责人说明人力缺口与公平性风险，由其决定是否批准临时放宽（如驿站允许男员工顶替、跨组支援、减少单岗人数）。"
    opt2_word = (f"若今天必须出表，是否批准临时放宽硬约束以填补「{gap}」？"
                 "例如驿站允许男员工顶替或减员、主持人跨组支援。请明确批准范围，我将标注「临时放宽」后生成。")

    opt3_title = "方案三：生成临时草案并标注待人工核对"
    opt3_pros = "最快拿到可讨论的草稿，争取时间。"
    opt3_cons = "无法保证合规与负荷均衡，不能直接执行，必须经负责人逐人复核签字。"
    opt3_comm = "向场馆负责人说明当前证据不足，由其决定是否接受草案并组织逐人复核。"
    opt3_word = (f"当前无法满足硬约束（{gap}），可仅按人员标签与已知请假生成「临时草案」，"
                 "标注「未校验，待负责人复核」后由您逐人确认方可执行。")

    return [
        {"title": opt1_title, "pros": opt1_pros, "cons": opt1_cons,
         "communicate_with": opt1_comm, "suggested_wording": opt1_word},
        {"title": opt2_title, "pros": opt2_pros, "cons": opt2_cons,
         "communicate_with": opt2_comm, "suggested_wording": opt2_word},
        {"title": opt3_title, "pros": opt3_pros, "cons": opt3_cons,
         "communicate_with": opt3_comm, "suggested_wording": opt3_word},
    ]


def run_schedule(personnel, history, demands, changes, template_date, future_dates=None, buxiu=None, rest_by_day=None):
    groups = build_groups(personnel)
    if future_dates is None:
        future_dates = compute_future_dates(template_date, 3)
    if buxiu is None:
        buxiu = {}
    # 调休轮转由外部状态文件驱动；仅当未提供时才从历史回退（用于离线测试）
    if rest_by_day is None:
        rest_by_day = compute_rest(groups["all"], history, future_dates, changes)
    un = _merge_unavailable(changes, buxiu, rest_by_day, future_dates)

    # 历史负荷（近 3 天出现次数）作为均衡基线
    load = {n: 0 for n in groups["all"]}
    for h in history:
        for label, cells in h["assignments"].items():
            for col, names in cells.items():
                for n in names:
                    if n in load:
                        load[n] += 1

    days = []
    all_issues = []
    infeasible = False
    for fd in future_dates:
        iso = fd["date"].isoformat()
        un_day = un[iso]
        avail_women = [w for w in groups["women"] if w not in un_day["full"]]
        avail_hosts = [h for h in groups["hosts"] if h not in un_day["full"]]
        avail_tech = [t for t in groups["tech"] if t not in un_day["full"]]
        avail_leads = [l for l in groups["leads"] if l not in un_day["full"]]
        day_issues = []
        prefill_hard = []
        if len(avail_women) < 4:
            prefill_hard.append(f"女员工不足4人（仅 {len(avail_women)}），无法填满游客服务驿站(4人/日)")
        if len(avail_hosts) < 2:
            prefill_hard.append(f"主持人不足2人（仅 {len(avail_hosts)}）")
        if len(avail_tech) < 1:
            prefill_hard.append("技术岗无人可用，正堂音响/园区广播无法排")
        if len(avail_leads) < 1:
            prefill_hard.append("负责人无人可用，行政事务/巡园/机动无法排")

        res, _, fill_hard, fill_soft = schedule_day(demands, groups, un_day, load)
        day_issues = prefill_hard + fill_hard + fill_soft
        if prefill_hard or fill_hard:
            infeasible = True
        # 把当天结果计入后续负荷基线（跨天均衡）
        for k, names in res.items():
            for n in names:
                load[n] = load.get(n, 0) + 1
        days.append({
            "date": iso, "weekday": fd["weekday"],
            "assignments": res,
            "rest": rest_by_day.get(iso, []),
            "unavailable": {"full": sorted(un_day["full"]),
                            "am": sorted(un_day["am"]), "pm": sorted(un_day["pm"])},
            "issues": day_issues,
        })
        all_issues += [f"[{iso}] {i}" for i in day_issues]

    if infeasible:
        return {
            "status": "blocked",
            "days": days,
            "issues": all_issues,
            "options": _build_options(all_issues, un, future_dates),
            "rest_by_day": rest_by_day,
        }
    return {
        "status": "ok",
        "days": days,
        "issues": all_issues,
        "options": [],
        "rest_by_day": rest_by_day,
    }


if __name__ == "__main__":
    root = find_project_root(_HERE)
    personnel = parse_personnel(os.path.join(root, "输入", "人员", "人员.xls"))
    history = parse_history(os.path.join(root, "输入", "过往三天"), personnel)
    buxiu = parse_buxiu(os.path.join(root, "输入", "过往三天", "8.13值班表.xls"), personnel)
    with open(os.path.join(_HERE, "template_map.json"), encoding="utf-8") as f:
        tmap = json.load(f)
    template_date = datetime.date.fromisoformat(history[-1]["date"])
    res = run_schedule(personnel, history, tmap["demands"], {"unavailable": [], "rest": []}, template_date, buxiu=buxiu)
    print("STATUS:", res["status"])
    for d in res["days"]:
        print(d["date"], d["weekday"], "rest=", d["rest"])
        for k, v in d["assignments"].items():
            print("   ", k, "->", v)
        if d["issues"]:
            print("   ISSUES:", d["issues"])
