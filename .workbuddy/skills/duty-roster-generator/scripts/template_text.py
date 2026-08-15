# -*- coding: utf-8 -*-
"""把求解器产出的「岗位 -> 人名列表」渲染成模板单元格文本（含角色后缀）。

只做纯文本拼接，确定性、可单测，不触碰任何单元格格式。
"""


def _join(names):
    return "、".join(names)


RENDERERS = {
    "驿站上午": lambda n: _join(n) + "（开六房、机动）" if n else "",
    "驿站下午": lambda n: _join(n) + "（机动）" if n else "",
    "闸机开馆": lambda n: (n[0] + "（开馆）") if n else "",
    "音响上午": lambda n: (n[0] + "（开六房）") if n else "",
    "音响下午": lambda n: (n[0] + "（关六房）") if n else "",
    "演艺上午": lambda n: n[0] if n else "",
    "演艺下午": lambda n: n[0] if n else "",
    "宣传": lambda n: n[0] if n else "",
    "开衙": lambda n: _join(n) if n else "",
    "讲解11:30": lambda n: ("11:30 " + n[0]) if n else "",
    "讲解16:30": lambda n: ("16:30 " + n[0]) if n else "",
    "侨批上午": lambda n: (n[0] + "（开馆）") if n else "",
    "侨批协同": lambda n: n[0] if n else "",
    "侨批下午": lambda n: ("15:00—17:30 " + n[0]) if n else "",
    "行政上午": lambda n: n[0] if n else "",
    "行政下午": lambda n: n[0] if n else "",
}


def render_cell(key, names, tail=""):
    """渲染单个需求单元格。闸机关馆需拼接模板固定尾部（安保/闭馆说明）。"""
    if key == "闸机关馆":
        if names:
            return names[0] + "（关六房）" + tail
        return tail
    fn = RENDERERS.get(key)
    if not fn:
        return _join(names)
    return fn(names)


def render_date(date_iso, weekday_cn):
    """2026-08-14 + 周五 -> 8月14日周五。"""
    import datetime
    d = datetime.date.fromisoformat(date_iso)
    return f"{d.month}月{d.day}日{weekday_cn}"


def rebuild_note(template_note, rest, leave, buxiu, sick, annual):
    """保留模板备注的常设条目(1-4,7-9)，用当天的调休/请假/补休/病假/年休假替换原 5、6 条。"""
    import re
    if not template_note:
        return ""
    segs = re.split(r"(?=\d+[、.])", template_note)
    keep = [s for s in segs if not re.match(r"\s*[56][、.]", s)]
    dyn = ""
    if rest:
        dyn += "5、" + "、".join(rest) + "调休；"
    parts = []
    if leave:
        parts.append("、".join(leave) + "请假")
    if buxiu:
        parts.append("、".join(buxiu) + "补休")
    if sick:
        parts.append("、".join(sick) + "病假")
    if annual:
        parts.append("、".join(annual) + "年休假")
    if parts:
        dyn += "6、" + "；".join(parts) + "；"
    out = []
    inserted = False
    for s in keep:
        if re.match(r"\s*7[、.]", s) and dyn and not inserted:
            out.append(dyn)
            inserted = True
        out.append(s)
    if dyn and not inserted:
        out.append(dyn)
    return "".join(out)
