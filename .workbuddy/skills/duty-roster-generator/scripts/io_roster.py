# -*- coding: utf-8 -*-
"""读取排班系统的真实 Excel 输入：人员名册、过往三天实际值班表、补休日历。

只用本地脚本完成解析，对话层只读取本模块产出的结构化结果，从而把 token 消耗压到最低。
"""
import os
import re
import sys
import glob
import datetime

_HERE = os.path.dirname(os.path.abspath(__file__))
_LIB = os.path.join(os.path.dirname(_HERE), ".lib")
if _LIB not in sys.path:
    sys.path.insert(0, _LIB)
import xlrd  # noqa: E402

# 历史排班里出现的别名 -> 名册规范名（名册唯一来源 = 人员表）。
# 注意：林璜 ≠ 林煌，是独立的人，且（连同慧銮）均已离职——
# 不在 19 人名册的历史人员一律忽略，不再参与解析与轮转基线。
NAME_ALIASES = {}

WEEKDAY_CN = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]


def _col_letter(idx):
    """0-based 列号 -> 列字母。"""
    s = ""
    idx += 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


# 已知角色标签（用于识别相邻单元格里的纯标签写法）
TAG_SET = {"女", "主持人", "技术", "负责人", "宣传", "通用男"}


def _cell_tags(text):
    """从单个单元格文本抽取角色标签（兼容 `（标签）` 与纯 `标签` 两种写法）。"""
    if not isinstance(text, str):
        return []
    s = re.sub(r"（[^）]*）", " ", text)   # 去掉括号后缀
    out = []
    for t in re.split(r"\s+", s.strip()):
        t = t.strip()
        if not t:
            continue
        if t in TAG_SET:
            out.append(t)
        else:
            for tag in TAG_SET:
                if tag in t:
                    out.append(tag)
    return out


def parse_personnel(path):
    """解析 人员.xls。兼容两种布局：
    (a) 单列 `Name（标签）（标签）`；
    (b) A 列放名字、同行 B/C/D… 相邻格放纯标签（名字是变量，标签是常量锚点）。
    返回 [{name, tags}]。
    """
    book = xlrd.open_workbook(path)
    ws = book.sheet_by_index(0)
    people = []
    for r in range(ws.nrows):
        v = ws.cell_value(r, 0)
        if not isinstance(v, str) or not v.strip():
            continue
        text = v.strip()
        tags = re.findall(r"（([^）]+)）", text)
        name = re.sub(r"（[^）]*）", "", text).strip()
        if not name:
            continue
        # 兼容布局(b)：扫描相邻单元格里的纯标签
        for c in range(1, min(ws.ncols, 6)):
            cv = ws.cell_value(r, c)
            if cv:
                tags.extend(_cell_tags(cv))
        # 去重保序
        seen = set()
        tags = [t for t in tags if not (t in seen or seen.add(t))]
        people.append({"name": name, "tags": tags})
    return people


def _extract_names(text, valid_names):
    """从单元格/备注文本里抽取名册中的人名（含别名归一）。"""
    if not text or not isinstance(text, str):
        return []
    s = re.sub(r"（[^）]*）", " ", text)              # 去掉角色后缀
    s = re.sub(r"\d{1,2}:\d{2}\s*[—-]\s*\d{1,2}:\d{2}", " ", s)  # 时段
    s = re.sub(r"\d{1,2}:\d{2}", " ", s)
    tokens = re.split(r"[、，,；;/\s]+", s)
    out = []
    for t in tokens:
        t = t.strip()
        if not t:
            continue
        if t in NAME_ALIASES:
            m = NAME_ALIASES[t]
        elif t in valid_names:
            m = t
        else:
            m = next((n for n in valid_names if n in t), None)
        if m and m not in out:
            out.append(m)
    return out


def _extract_date(sheet):
    """在表头区域找 `X月X日周X`。"""
    for r in range(min(sheet.nrows, 6)):
        for c in range(min(sheet.ncols, 8)):
            v = sheet.cell_value(r, c)
            if isinstance(v, str):
                m = re.search(r"(\d{1,2})月(\d{1,2})日(周[一二三四五六日])", v)
                if m:
                    return int(m.group(1)), int(m.group(2)), m.group(3)
    return None


def _extract_unavailable(note, valid_names):
    """从 `其他` 备注抽取 调休/请假/补休/病假/年休假 名单。

    支持同一编号条目内含多类离岗，如「6、瑞琼补休、毓廷请假」——按各类
    关键词出现位置切分名字，避免旧实现一段只抓第一个类型而漏掉后续人。
    """
    res = {"调休": [], "请假": [], "补休": [], "病假": [], "年休假": []}
    KIND_ORDER = ["病假", "年休假", "补休", "请假", "调休"]
    for seg in re.split(r"[；;]", note):
        seg = re.sub(r"^\d+\s*", "", seg)          # 去掉条目编号
        seg = re.sub(r"^\d+[、.]\s*", "", seg)
        # 找出该段内所有离岗关键词及其位置，按出现顺序切分名字
        hits = [(seg.find(k), k) for k in KIND_ORDER if k in seg]
        if not hits:
            continue
        hits.sort(key=lambda x: x[0])
        for i, (pos, kind) in enumerate(hits):
            start = 0 if i == 0 else hits[i - 1][0] + len(hits[i - 1][1])
            pre = seg[start:pos]
            names = _extract_names(pre, valid_names)
            res[kind].extend(names)
    # 去重
    for k in res:
        res[k] = list(dict.fromkeys(res[k]))
    return res


def parse_history(history_dir, personnel):
    """解析 过往三天/*.xls。返回按日期升序的日记录列表。"""
    names = [p["name"] for p in personnel]
    records = []
    for path in sorted(glob.glob(os.path.join(history_dir, "*.xls"))):
        book = xlrd.open_workbook(path)
        if "8月排班表" not in book.sheet_names():
            continue
        ws = book.sheet_by_name("8月排班表")
        d = _extract_date(ws)
        if not d:
            continue
        month, day, wd = d
        date = datetime.date(2026, month, day)
        # 岗位行：列 A 为标签，B/C/D/G 为人员
        assignments = {}
        note = ""
        for r in range(ws.nrows):
            label = ws.cell_value(r, 0)
            if not isinstance(label, str):
                continue
            label = label.strip()
            if not label:
                continue
            if label.startswith("其他"):
                note = " ".join(
                    str(ws.cell_value(r, c)) for c in range(ws.ncols)
                    if isinstance(ws.cell_value(r, c), str) and ws.cell_value(r, c).strip()
                )
                continue
            cells = {}
            for c in range(1, ws.ncols):
                v = ws.cell_value(r, c)
                if isinstance(v, str) and v.strip():
                    cells[_col_letter(c)] = _extract_names(v, names)
            if cells:
                assignments[label] = cells
        una = _extract_unavailable(note, names) if note else {"调休": [], "请假": [], "补休": [], "病假": [], "年休假": []}
        records.append({
            "path": path,
            "date": date.isoformat(),
            "weekday": wd,
            "assignments": assignments,
            "rest": una["调休"],
            "leave": una["请假"],
            "buxiu": una["补休"],
            "sick": una["病假"],
            "annual": una["年休假"],
            "note": note,
        })
    records.sort(key=lambda x: x["date"])
    return records


def parse_buxiu(path, personnel):
    """解析 补休 日历 sheet：返回 {name: set(day_int)}（仅 2026.8月 列）。"""
    names = [p["name"] for p in personnel]
    book = xlrd.open_workbook(path)
    if "补休" not in book.sheet_names():
        return {}
    ws = book.sheet_by_name("补休")
    # 找到 2026.8月 列
    aug_col = None
    for c in range(ws.ncols):
        v = ws.cell_value(0, c)
        if isinstance(v, str) and "8月" in v and "2026" in v:
            aug_col = c
            break
    if aug_col is None:
        return {}
    out = {}
    for r in range(1, ws.nrows):
        name = ws.cell_value(r, 0)
        if not isinstance(name, str):
            continue
        name = re.sub(r"（[^）]*）", "", name).strip()
        if name not in names:
            continue
        v = ws.cell_value(r, aug_col)
        if not isinstance(v, str) or not v.strip():
            continue
        days = [int(x) for x in re.findall(r"(\d{1,2})日", v)]
        if days:
            out[name] = set(days)
    return out


def find_project_root(start):
    cur = os.path.abspath(start)
    for _ in range(6):
        if os.path.isdir(os.path.join(cur, "输入")):
            return cur
        parent = os.path.dirname(cur)
        if parent == cur:
            break
        cur = parent
    return os.path.abspath(start)


if __name__ == "__main__":
    base = find_project_root(_HERE)
    people = parse_personnel(os.path.join(base, "输入", "人员", "人员.xls"))
    print("PERSONNEL:", len(people))
    for p in people:
        print("  ", p["name"], p["tags"])
    hist = parse_history(os.path.join(base, "输入", "过往三天"), people)
    print("HISTORY days:", [h["date"] for h in hist])
    for h in hist:
        print(f"  {h['date']} {h['weekday']} rest={h['rest']} leave={h['leave']} buxiu={h['buxiu']}")
    bx = parse_buxiu(os.path.join(base, "输入", "过往三天", "8.13值班表.xls"), people)
    print("BUXIU aug:", bx)
