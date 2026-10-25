#!/usr/bin/env python3
"""animate.py v2 — 给构建好的 .pptx 注入动效（pptxgenjs 不支持动画，本脚本做 OOXML 后处理）。

用法: python3 animate.py <deck.pptx> <anim.json> [--out <output.pptx>]
  默认原地覆盖（build.js 重建后需重跑）。
  WPS 兼容：图表的进退场默认跳过（全屏放映不触发会隐藏图表）；顶层 "animate_charts": true 可开启（PowerPoint 生效）。

anim.json 结构:
{
  "transition":  {"type": "fade", "speed": "med"},        // 全场默认切换
  "transitions": {"1": {"type": "zoom", "dir": "in"}},    // 按页覆盖
  "slides": {
    "2": [ {"target": "*", "effect": "fade", "trigger": "afterPrev", "dur": 400, "stagger": 120} ],
    "4": [ {"target": "chart-revenue", "effect": "wipe", "dir": "right",
            "chart": "category", "trigger": "afterPrev", "dur": 700} ],
    "6": [ {"target": "bullets", "effect": "fade", "byPara": true,
            "trigger": "afterPrev", "dur": 350, "stagger": 250} ]
  }
}

target: "*" 该页全部形状按位置级联 | objectName（build.js 的 objectName）
trigger: afterPrev(自动接续) | withPrev(同时) | click(点击，开新组)
通用参数: dur(ms) delay(ms) stagger(ms) autoRev repeat(N 或 "indefinite") accel decel(0-100)

效果一览:
  auto  {"target": "*", "effect": "auto"} —— 角色感知自动编排（推荐默认）：
        标题/副标题/页脚/图注/装饰/辉光等"铬件"不动画；截图整块 fade、大数字 zoomIn、
        正文按行成组级联；单页进场总时长收敛在 ~3.5s 内。可用 dur/stagger 覆盖节奏
        （高端档建议 {"effect":"auto","dur":650,"stagger":200}）。
  进场  fade appear wipe(dir) flyIn(dir) zoomIn(from%)
        wheel wheel2 wheel4 wheel8 checkerboard(dir) blinds(dir) circle diamond
        dissolveIn strips(dir) wedge randomBar(dir)
  退场  fadeOut zoomOut wipeOut(dir) flyOut(dir)
        （任一滤镜效果加 "out": true 即为退场版）
  强调  pulse(scale%) grow(scale%) shrink(scale%) spin(deg) colorTo(RRGGBB)
  路径  move(dx, dy)（slide 宽高小数，y 向下为正）
  构建  "chart": "category"|"series"  图表按类/按系列逐一构建
        "byPara": true                 文本逐段进场（target 须为具体名称）

编排原则（写 anim.json 时遵守）: 铬件不动、焦点先行、同行成组、有理由才强调；
点名效果(name)负责"特殊时刻"，其余交给 auto——不要用 "*" 手工扫场。
"""
import json
import sys

from pptx import Presentation
from pptx.oxml.ns import qn
from lxml import etree

FLY_FROM = {"top": (1, "0-#ppt_h/2", "y"), "bottom": (4, "1+#ppt_h/2", "y"),
            "left": (8, "0-#ppt_w/2", "x"), "right": (2, "1+#ppt_w/2", "x")}
WIPE_FILTER = {"right": "wipe(right)", "left": "wipe(left)", "down": "wipe(down)", "up": "wipe(up)"}
WIPE_SUBTYPE = {"right": 8, "left": 2, "down": 1, "up": 4}
TRANSITION_SPEED = {"slow": "slow", "med": "med", "fast": "fast"}

FILTER_FAMILY = {
    # name: (filter 模板, presetID, 默认 dir 值域)
    "wheel":  ("wheel({d})", 21, [1]),
    "wheel2": ("wheel(2)", 21, [2]),
    "wheel4": ("wheel(4)", 21, [4]),
    "wheel8": ("wheel(8)", 21, [8]),
    "checkerboard": ("checkerboard({d})", 5, ["across", "down"]),
    "blinds": ("blinds({d})", 3, ["horizontal", "vertical"]),
    "randomBar":    ("randombar({d})", 4, ["horizontal", "vertical"]),
    "strips": ("strips({d})", 31, ["downRight", "downLeft", "upRight", "upLeft"]),
    "circle": ("circle", 15, None),
    "diamond": ("diamond(in)", 18, None),
    "dissolveIn": ("dissolve", 9, None),
    "wedge": ("wedge", 27, None),
    "fade": ("fade", 10, None),
    "plus": ("plus", 13, None),
}

ENTRANCE = {"fade", "appear", "wipe", "flyIn", "zoomIn", "dissolveIn", "circle", "diamond",
            "wedge", "plus", "wheel", "wheel2", "wheel4", "wheel8", "checkerboard",
            "blinds", "randomBar", "strips"}
EXIT_SET = {"fadeOut", "zoomOut", "wipeOut", "flyOut"}
EMPHASIS = {"pulse", "grow", "shrink", "spin", "colorTo"}


def make(parent, tag, attrib=None, *children):
    e = etree.Element(qn(tag))
    if attrib:
        for k, v in attrib.items():
            e.set(k, str(v))
    for c in children:
        e.append(c)
    if parent is not None:
        parent.append(e)
    return e


class Ids:
    def __init__(self):
        self.n = 2
    def next(self):
        self.n += 1
        return self.n


def tune(ctn, eff):
    """把 autoRev/repeat/accel/decel 写到行为时间节点上。"""
    if eff.get("autoRev"):
        ctn.set("autoRev", "1")
    rep = eff.get("repeat")
    if rep is not None:
        ctn.set("repeatCount", rep if isinstance(rep, str) else f"{rep}x")
    for k in ("accel", "decel"):
        if eff.get(k) is not None:
            ctn.set(k, str(int(eff[k] * 1000)))
    return ctn


def shapes_of(slide):
    """顶层形状 [(spid, name, kind, x_emu, y_emu)]，kind: sp|pic|chart|table|group。"""
    out = []
    tree = slide._element.find(qn("p:cSld")).find(qn("p:spTree"))
    for node in tree:
        tag = etree.QName(node).localname
        if tag not in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp"):
            continue
        cnv = node.find(".//" + qn("p:cNvPr"))
        if cnv is None:
            continue
        sid, name = cnv.get("id"), cnv.get("name") or ""
        kind = tag
        if tag == "graphicFrame":
            kind = "chart" if node.find(".//" + qn("c:chart")) is not None else "table"
        x = y = 0
        xfrm = node.find(".//" + qn("a:xfrm"))
        if xfrm is not None:
            off = xfrm.find(qn("a:off"))
            if off is not None:
                x, y = int(off.get("x", 0)), int(off.get("y", 0))
        out.append((int(sid), name, kind, x, y))
    return out


def walk_pptx_shapes(shapes):
    for s in shapes:
        if s.shape_type == 6:
            yield from walk_pptx_shapes(s.shapes)
        else:
            yield s


def para_indexes(slide, spid):
    """有文字的段落索引列表。"""
    for s in walk_pptx_shapes(slide.shapes):
        if s.shape_id == spid and getattr(s, "has_text_frame", False):
            return [i for i, p in enumerate(s.text_frame.paragraphs) if p.text.strip()]
    return []


def shape_by_id(slide, spid):
    for s in walk_pptx_shapes(slide.shapes):
        if s.shape_id == spid:
            return s
    return None


def find_containers(slide, shapes, spids, used):
    """容器归属：无字静态大形状几何包含动画文字 → {text_spid: container_spid}。
    卡片框与其内容必须同批动画（框先、文字紧随），否则观感割裂。"""
    res = {}
    geom = {}
    texts = {}
    for spid, name, kind, x, y in shapes:
        sh = shape_by_id(slide, spid)
        g = None
        try:
            if sh is not None and sh.left is not None:
                g = (sh.left / EMU, sh.top / EMU, sh.width / EMU, sh.height / EMU)
        except Exception:
            g = None
        geom[spid] = g
        texts[spid] = (sh.text_frame.text.strip()
                       if sh is not None and getattr(sh, "has_text_frame", False) else "")
    for spid in spids:
        if spid in used or spid in res:
            continue
        g = geom.get(spid)
        if not g:
            continue
        cx, cy = g[0] + g[2] / 2, g[1] + g[3] / 2
        best = None
        for spid2, name2, kind2, x2, y2 in shapes:
            if spid2 in used or spid2 in res or spid2 == spid:
                continue
            g2 = geom.get(spid2)
            if not g2 or texts.get(spid2):
                continue  # 只考虑无字形状作为容器
            l, tp, w, h = g2
            if w * h < 0.55 or w < 0.9:
                continue
            if l - 0.06 <= cx <= l + w + 0.06 and tp - 0.06 <= cy <= tp + h + 0.06:
                if best is None or w * h < geom[best][2] * geom[best][3]:
                    best = spid2  # 最小包含容器（最具体）
        if best is not None:
            res[spid] = best
    return res


def max_font_pt(sh, default=18.0):
    best = 0.0
    if sh is not None and getattr(sh, "has_text_frame", False):
        for pa in sh.text_frame.paragraphs:
            for r in pa.runs:
                if r.font.size is not None:
                    best = max(best, r.font.size.pt)
    return best or default


EMU = 914400


def classify(slide, shapes):
    """按版式角色分类: static(铬件) | focal(焦点) | body(正文) | chart。
    判定用文字内容/字号/位置——标题区、页脚带、图注、无字装饰与辉光图都不动画。"""
    roles = {}
    for spid, name, kind, x, y in shapes:
        y_in = y / EMU
        sh = shape_by_id(slide, spid)
        txt = (sh.text_frame.text.strip()
               if sh is not None and getattr(sh, "has_text_frame", False) else "")
        if kind in ("chart", "table"):
            roles[spid] = "chart"
        elif kind == "pic":
            roles[spid] = "static" if "glow" in (name or "").lower() else "focal"
        elif not txt:
            roles[spid] = "static"                       # 线条/卡片底/圆环/装饰
        else:
            pt = max_font_pt(sh)
            if y_in > 6.9:
                roles[spid] = "static"                   # 页脚/页码/来源行
            elif y_in < 1.45 and pt >= 20:
                roles[spid] = "static"                   # 标题
            elif y_in < 1.45 and pt <= 15:
                roles[spid] = "static"                   # 副标题
            elif pt <= 12.5 and y_in > 5.3:
                roles[spid] = "static"                   # 图注
            elif pt >= 38:
                roles[spid] = "focal"                    # 大数字
            else:
                roles[spid] = "body"
    return roles


def build_auto(slide, shapes, used, base, page_no, t0=0):
    """自动编排：返回 {'auto': True, 'items': [(eff, item, start)]}（start 含 t0 偏移）。
    规则: 铬件静态；截图 fade、大数字 zoomIn、正文 fade；同行(Δy<0.55)成组，
    组内小 stagger、组间顺序推进；总时长 >3.6s 自动压缩节奏。
    容器规则: 无字大卡片/相框与其内部动画文字**同批出现**（框先、文字紧随）——避免框文割裂。"""
    roles = classify(slide, shapes)
    items = []
    n_static = 0
    geom = {}
    for spid, name, kind, x, y in shapes:
        role = roles.get(spid)
        sh = shape_by_id(slide, spid)
        g = None
        try:
            if sh is not None and sh.left is not None:
                g = (sh.left / EMU, sh.top / EMU, sh.width / EMU, sh.height / EMU)
        except Exception:
            g = None
        geom[spid] = g
        if role in (None, "chart", "static") or spid in used:
            if role == "static":
                n_static += 1
            continue
        items.append([y / EMU, x / EMU, spid, kind, role])
    items.sort(key=lambda it: (it[0], it[1]))

    # —— 容器归属：无字静态大形状 ⊃ 动画文字 → 与其承载的第一条文字同批（框先、文字紧随）——
    cmap = find_containers(slide, shapes, {it[2] for it in items}, used)
    claimed = set()
    for text_spid, csp in cmap.items():
        if csp in used or csp in claimed:
            continue
        it = next(it for it in items if it[2] == text_spid)
        g = geom.get(csp)
        if not g:
            continue
        claimed.add(csp)
        used.add(csp)
        items.append([it[0], g[0] + 0.005, csp, "static", "container"])
    items.sort(key=lambda it: (it[0], it[1]))
    batches = []
    for y, x, spid, kind, role in items:
        y_in = y / EMU
        if batches and abs(y_in - batches[-1][0]) < 0.55 and len(batches[-1][1]) < 4:
            batches[-1][1].append((x, spid, kind, role))
        else:
            batches.append([y_in, [(x, spid, kind, role)]])
    stag = base.get("stagger", 90)
    gap = 90
    dur = base.get("dur", 380)
    projected = len(batches) * (dur + gap) + sum(len(b[1]) for b in batches) * stag
    if projected > 3600:
        k = 3600 / projected
        stag, gap = max(45, int(stag * k)), max(40, int(gap * k))
    out, t = [], 0
    for _y, its in batches:
        its.sort()
        for j, (_x, spid, kind, role) in enumerate(its):
            if role == "focal":
                e = {"effect": "fade" if kind == "pic" else "zoomIn",
                     "trigger": "afterPrev", "from": 55,
                     "dur": min(dur + 140, 900)}
            else:
                e = {"effect": "fade", "trigger": "afterPrev", "dur": dur}
            out.append((e, (e, spid, kind, None), t + j * stag))
        t += dur + gap + (len(its) - 1) * stag
    print(f"  ~ p{page_no} auto 编排: {len(out)} 动画 / {n_static} 静态 / {len(batches)} 批（预计 ~{t}ms）")
    return {"auto": True, "items": out}


def resolve_targets(effects, shapes, slide):
    """展开 target → [(eff, spid, kind, prg)]；prg=(st,end) 表示段落范围。"""
    resolved = []
    for eff in effects:
        if eff["target"] == "*":
            ordered = sorted(shapes, key=lambda s: (s[4], s[3]))
            for i, s in enumerate(ordered):
                e2 = dict(eff)
                if "stagger" in eff:
                    e2["delay"] = eff.get("delay", 0) + i * eff["stagger"]
                resolved.append((e2, s[0], s[2], None))
        else:
            m = [s for s in shapes if s[1] == eff["target"]]
            if not m:
                m = [s for s in shapes if eff["target"] in s[1]]
            if not m:
                print(f"  ! 未找到目标形状: {eff['target']}（跳过）")
                continue
            spid, kind = m[0][0], m[0][2]
            if eff.get("byPara"):
                idxs = para_indexes(slide, spid)
                if not idxs:
                    print(f"  ! {eff['target']} 无文本段落（跳过 byPara）")
                    continue
                for k, st in enumerate(idxs):
                    e2 = dict(eff)
                    e2["delay"] = eff.get("delay", 0) + k * eff.get("stagger", 220)
                    resolved.append((e2, spid, kind, (st, st)))
            else:
                resolved.append((dict(eff), spid, kind, None))
    return resolved


def tgt(item):
    spid, kind, prg = item[1], item[2], item[3]
    t = etree.Element(qn("p:tgtEl"))
    if prg is not None:
        tx = make(t, "p:txEl")
        make(tx, "p:pRg", {"st": prg[0], "end": prg[1]})
    else:
        make(t, "p:spTgt", {"spid": spid})
    return t


def bhvr(ids, item, attr=None, dur=None, extra_delay=None, additive=None, eff=None):
    b = etree.Element(qn("p:cBhvr"))
    if additive:
        b.set("additive", additive)
    a = {"id": ids.next()}
    if dur is not None:
        a["dur"] = dur
    if extra_delay is not None:
        a["fill"] = "hold"
    ctn = make(b, "p:cTn", a)
    if extra_delay is not None:
        st = make(ctn, "p:stCondLst")
        make(st, "p:cond", {"delay": extra_delay})
    if eff:
        tune(ctn, eff)
    b.append(tgt(item))
    if attr:
        al = make(b, "p:attrNameLst")
        make(al, "p:attrName").text = attr
    return b


def set_vis(ids, item, val, delay=0):
    s = etree.Element(qn("p:set"))
    b = bhvr(ids, item, attr="style.visibility", dur=1,
             extra_delay=delay if val == "hidden" else 0)
    s.append(b)
    make(s, "p:to").append(make(None, "p:strVal", {"val": val}))
    return s


def anim_effect(ids, item, transition, filt, dur, eff=None):
    ae = etree.Element(qn("p:animEffect"))
    ae.set("transition", transition)
    ae.set("filter", filt)
    ae.append(bhvr(ids, item, dur=dur, eff=eff))
    return ae


def anim_pos(ids, item, attr, from_v, to_v, dur, eff=None):
    an = etree.Element(qn("p:anim"))
    an.set("calcmode", "lin")
    an.set("valueType", "num")
    an.append(bhvr(ids, item, attr=attr, dur=dur, additive="base", eff=eff))
    tav = make(an, "p:tavLst")
    for tm, v in ((0, from_v), (100000, to_v)):
        t = make(tav, "p:tav", {"tm": tm})
        val = make(t, "p:val")
        make(val, "p:strVal", {"val": v})
    return an


def anim_scale(ids, item, from_pct, to_pct, dur, eff=None):
    sc = etree.Element(qn("p:animScale"))
    sc.append(bhvr(ids, item, dur=dur, eff=eff))
    if from_pct is not None:
        make(sc, "p:from", {"x": int(from_pct * 1000), "y": int(from_pct * 1000)})
    make(sc, "p:to", {"x": int(to_pct * 1000), "y": int(to_pct * 1000)})
    return sc


def anim_rot(ids, item, by_deg, dur, eff=None):
    r = etree.Element(qn("p:animRot"))
    r.set("by", str(int(by_deg * 60000)))
    r.append(bhvr(ids, item, attr="r", dur=dur, eff=eff))
    return r


def anim_clr(ids, item, hexcolor, dur, eff=None):
    c = etree.Element(qn("p:animClr"))
    c.set("clrSpc", "rgb")
    c.set("dir", "cw")
    c.append(bhvr(ids, item, attr="style.color", dur=dur, eff=eff))
    to = make(c, "p:to")
    make(to, "a:srgbClr", {"val": hexcolor})
    return c


def anim_motion(ids, item, dx, dy, dur, eff=None):
    m = etree.Element(qn("p:animMotion"))
    m.set("origin", "layout")
    m.set("pathEditMode", "relative")
    m.set("path", f"M 0 0 L {dx} {dy} E")
    m.set("ptsTypes", "AA")
    m.append(bhvr(ids, item, dur=dur, eff=eff))
    return m


def filter_string(name, eff):
    tmpl, _pid, dirs = FILTER_FAMILY[name]
    d = eff.get("dir", dirs[0] if dirs else None)
    return tmpl.format(d=d) if "{d}" in tmpl else tmpl


def effect_children(ids, eff, item, dur):
    """返回 (children, is_exit, needs_hidden_at_end)。"""
    e = eff["effect"]
    dirn = eff.get("dir", "right")
    ch = []
    if e in FILTER_FAMILY:  # 含 fade / dissolveIn / circle / 棋盘 / 百叶窗 / 轮子…
        ch.append(set_vis(ids, item, "visible"))
        ch.append(anim_effect(ids, item, "in", filter_string(e, eff), dur, eff))
        return ch, False, False
    if e == "appear":
        ch.append(set_vis(ids, item, "visible"))
        return ch, False, False
    if e == "wipe":
        ch.append(set_vis(ids, item, "visible"))
        ch.append(anim_effect(ids, item, "in", WIPE_FILTER[dirn], dur, eff))
        return ch, False, False
    if e == "flyIn":
        _sub, from_v, axis = FLY_FROM[dirn]
        other = "ppt_x" if axis == "y" else "ppt_y"
        ch.append(set_vis(ids, item, "visible"))
        ch.append(anim_pos(ids, item, other, f"#{other}", f"#{other}", dur, eff))
        ch.append(anim_pos(ids, item, f"ppt_{axis}", from_v, f"#ppt_{axis}", dur, eff))
        return ch, False, False
    if e == "zoomIn":
        ch.append(set_vis(ids, item, "visible"))
        ch.append(anim_effect(ids, item, "in", "fade", dur, eff))
        ch.append(anim_scale(ids, item, eff.get("from", 30), 100, dur, eff))
        return ch, False, False
    # —— 退场 ——
    if e == "fadeOut":
        ch.append(anim_effect(ids, item, "out", "fade", dur, eff))
        return ch, True, True
    if e == "zoomOut":
        ch.append(anim_effect(ids, item, "out", "fade", dur, eff))
        ch.append(anim_scale(ids, item, 100, eff.get("to", 30), dur, eff))
        return ch, True, True
    if e == "wipeOut":
        ch.append(anim_effect(ids, item, "out", WIPE_FILTER[dirn], dur, eff))
        return ch, True, True
    if e == "flyOut":
        _sub, to_v, axis = FLY_FROM[dirn]
        other = "ppt_x" if axis == "y" else "ppt_y"
        ch.append(anim_pos(ids, item, other, f"#{other}", f"#{other}", dur, eff))
        ch.append(anim_pos(ids, item, f"ppt_{axis}", f"#ppt_{axis}", to_v, dur, eff))
        return ch, True, True
    if eff.get("out") and e in FILTER_FAMILY:
        ch.append(anim_effect(ids, item, "out", filter_string(e, eff), dur, eff))
        return ch, True, True
    # —— 强调（不动可见性）——
    if e == "pulse":
        scale = eff.get("scale", 106)
        e2 = dict(eff)
        e2.setdefault("autoRev", True)
        ch.append(anim_scale(ids, item, 100, scale, dur, e2))
        return ch, False, False
    if e in ("grow", "shrink"):
        default = 120 if e == "grow" else 80
        ch.append(anim_scale(ids, item, None, eff.get("scale", default), dur, eff))
        return ch, False, False
    if e == "spin":
        ch.append(anim_rot(ids, item, eff.get("deg", 360), dur, eff))
        return ch, False, False
    if e == "colorTo":
        ch.append(anim_clr(ids, item, eff.get("color", "E8590C"), dur, eff))
        return ch, False, False
    if e == "move":
        ch.append(anim_motion(ids, item, eff.get("dx", 0.25), eff.get("dy", 0), dur, eff))
        return ch, False, False
    raise ValueError(f"未知效果: {e}")


def effect_par(ids, eff, item, start, node_type):
    e = eff["effect"]
    dur = eff.get("dur", 450)
    dirn = eff.get("dir", "right")
    if e in FILTER_FAMILY:
        pid = FILTER_FAMILY[e][1]
        pcls, psub = ("exit", 0) if eff.get("out") else ("entr", 0)
    elif e == "appear":
        pid, pcls, psub = 1, "entr", 0
    elif e == "wipe":
        pid, pcls, psub = 22, "entr", WIPE_SUBTYPE[dirn]
    elif e == "wipeOut":
        pid, pcls, psub = 22, "exit", WIPE_SUBTYPE[dirn]
    elif e in ("flyIn", "flyOut"):
        pid, pcls, psub = 2, "entr" if e == "flyIn" else "exit", FLY_FROM[dirn][0]
    elif e in ("zoomIn", "zoomOut"):
        pid, pcls, psub = 23, "entr" if e == "zoomIn" else "exit", 16
    elif e == "fadeOut":
        pid, pcls, psub = 10, "exit", 0
    elif e in EMPHASIS:
        pid = {"pulse": 26, "grow": 6, "shrink": 6, "spin": 8, "colorTo": 1}[e]
        pcls, psub = "emph", 0
    elif e == "move":
        pid, pcls, psub = 0, "path", 0
    else:
        pid, pcls, psub = 0, "entr", 0

    node_type = node_type or ("withEffect" if eff["trigger"] == "withPrev" else "afterEffect")
    par = etree.Element(qn("p:par"))
    ctn = make(par, "p:cTn", {"presetID": pid, "presetClass": pcls, "presetSubtype": psub,
                              "fill": "hold", "grpId": 0, "nodeType": node_type})
    st = make(ctn, "p:stCondLst")
    make(st, "p:cond", {"delay": start})
    child = make(ctn, "p:childTnLst")
    children, _is_exit, hide_at_end = effect_children(ids, eff, item, dur)
    for c in children:
        child.append(c)
    if hide_at_end:
        child.append(set_vis(ids, item, "hidden", delay=max(dur - 1, 0)))
    return par


def timing_tree(groups):
    """规范三层 par 结构（WPS/PowerPoint 通吃的 canonical 形态）:

    seq(mainSeq)
      └ par L1「点击组」  cTn stCondLst: indefinite(点击开始) | 0(进片自动开始)
        └ par L2「同时批次」 cTn stCondLst: delay=批次的组内累计起始
          └ par L3「效果」  cTn presetID/nodeType, stCondLst: delay=批内偏移
            └ behaviors(set/animEffect/anim*/animMotion...)
    自动接续(afterPrev) = 新的 L2 批次（delay 累计）；同时(withPrev) = 同批次内偏移。
    """
    ids = Ids()
    timing = etree.Element(qn("p:timing"))
    tnl = make(timing, "p:tnLst")
    root = make(tnl, "p:par")
    rc = make(root, "p:cTn", {"id": 1, "dur": "indefinite", "restart": "never",
                              "nodeType": "tmRoot"})
    rch = make(rc, "p:childTnLst")
    seq = make(rch, "p:seq", {"concurrent": 1, "nextAc": "seek"})
    main = make(seq, "p:cTn", {"id": 2, "dur": "indefinite", "nodeType": "mainSeq"})
    mch = make(main, "p:childTnLst")
    blds = {}
    for gi, g in enumerate(groups):
        l1 = make(mch, "p:par")
        l1c = make(l1, "p:cTn", {"id": ids.next(), "fill": "hold"})
        l1s = make(l1c, "p:stCondLst")
        make(l1s, "p:cond", {"delay": "indefinite"})
        if gi == 0 and g.get("auto"):
            # 开场自动组：引用 mainSeq(id=2) 的 onBegin —— 随页面进入即启动（PowerPoint 原生写法）
            cond = make(l1s, "p:cond", {"evt": "onBegin", "delay": 0})
            make(cond, "p:tn", {"val": 2})
        l1ch = make(l1c, "p:childTnLst")
        # —— L2 批次划分：trigger != withPrev 开新批 ——
        batches, cur_batch = [], None
        for i, (eff, item, start) in enumerate(g["items"]):
            if cur_batch is None or eff.get("trigger", "afterPrev") != "withPrev":
                cur_batch = {"start": start, "items": []}
                batches.append(cur_batch)
            cur_batch["items"].append((eff, item, start, i))
        for b in batches:
            l2 = make(l1ch, "p:par")
            l2c = make(l2, "p:cTn", {"id": ids.next(), "fill": "hold"})
            l2s = make(l2c, "p:stCondLst")
            make(l2s, "p:cond", {"delay": b["start"]})
            l2ch = make(l2c, "p:childTnLst")
            for eff, item, start, gi in b["items"]:
                rel = start - b["start"]
                if gi == 0 and not g["auto"]:
                    ntype = "clickEffect"
                elif eff.get("trigger", "afterPrev") == "withPrev":
                    ntype = "withEffect"
                else:
                    ntype = "afterEffect"
                l2ch.append(effect_par(ids, eff, item, rel, ntype))
                spid, kind, prg = item[1], item[2], item[3]
                if spid not in blds:
                    if kind == "chart" and eff.get("chart") in ("category", "series"):
                        blds[spid] = ("chart", eff["chart"])
                    elif prg is not None:
                        blds[spid] = ("para", None)
                    else:
                        blds[spid] = ("plain", None)
    pc = make(seq, "p:prevCondLst")
    make(make(pc, "p:cond", {"evt": "onPrev", "delay": 0}), "p:tgtEl").append(
        etree.Element(qn("p:sldTgt")))
    nc = make(seq, "p:nextCondLst")
    make(make(nc, "p:cond", {"evt": "onNext", "delay": 0}), "p:tgtEl").append(
        etree.Element(qn("p:sldTgt")))
    if blds:
        bl = make(timing, "p:bldLst")
        for spid, (btype, mode) in blds.items():
            if btype == "chart":
                bg = make(bl, "p:bldGraphic", {"spid": spid, "grpId": 0})
                sub = make(bg, "p:bldSub")
                make(sub, "a:bldChart", {"bld": mode, "animBg": "1"})
            elif btype == "para":
                make(bl, "p:bldP", {"spid": spid, "grpId": 0, "build": "p"})
            else:
                make(bl, "p:bldP", {"spid": spid, "grpId": 0})
    return timing


def transition_el(spec):
    t = etree.Element(qn("p:transition"))
    t.set("spd", TRANSITION_SPEED.get(spec.get("speed", "med"), "med"))
    typ = spec.get("type", "fade")
    d = spec.get("dir", "l")
    if typ == "push":
        make(t, "p:push", {"dir": d})
    elif typ == "wipe":
        make(t, "p:wipe", {"dir": d})
    elif typ == "cover":
        make(t, "p:cover", {"dir": d})
    elif typ == "uncover" or typ == "pull":
        make(t, "p:pull", {"dir": d})
    elif typ == "cut":
        make(t, "p:cut")
    elif typ == "dissolve":
        make(t, "p:dissolve")
    elif typ == "random":
        make(t, "p:random")
    elif typ == "zoom":
        make(t, "p:zoom", {"dir": d if d in ("in", "out") else "in"})
    elif typ == "split":
        make(t, "p:split", {"orient": spec.get("orient", "horz"),
                            "dir": d if d in ("in", "out") else "in"})
    elif typ == "wheel":
        make(t, "p:wheel", {"spokes": str(spec.get("spokes", 1))})
    elif typ == "comb":
        make(t, "p:comb", {"dir": spec.get("dir", "horz")})
    elif typ in ("newsflash", "plus", "wedge", "circle", "diamond"):
        make(t, f"p:{typ}")
    elif typ == "checkerboard":
        make(t, "p:checkerboard", {"dir": spec.get("dir", "vert")})
    elif typ == "blinds":
        make(t, "p:blinds", {"dir": spec.get("dir", "horz")})
    elif typ == "strips":
        make(t, "p:strips", {"dir": d})
    elif typ == "randomBar":
        make(t, "p:randomBar", {"dir": spec.get("dir", "horz")})
    else:
        make(t, "p:fade")
    return t


def apply(path, spec_path, out=None):
    spec = json.load(open(spec_path))
    prs = Presentation(path)
    slides = list(prs.slides)
    n_tr = n_an = 0
    for idx, slide in enumerate(slides, 1):
        sld = slide._element
        for tag in ("p:transition", "p:timing"):
            old = sld.find(qn(tag))
            if old is not None:
                sld.remove(old)
        tr = spec.get("transitions", {}).get(str(idx)) or spec.get("transition")
        if tr:
            sld.append(transition_el(tr))
            n_tr += 1
        effects = spec.get("slides", {}).get(str(idx))
        if not effects:
            continue
        shapes = shapes_of(slide)
        # 先解析全部点名效果（auto 需要排除已点名目标），再按声明顺序编排
        named_effs = [e for e in effects if e.get("effect") != "auto"]
        auto_seq = [e for e in effects if e.get("effect") == "auto"]
        named_res = []
        used = set()
        for eff in named_effs:
            for it in resolve_targets([eff], shapes, slide):
                named_res.append((eff, it))
                used.add(it[1])
        # WPS 兼容：图表挂进退场动画会在全屏放映时不触发（元素被隐藏）——
        # 默认跳过图表的进退场（随页面直接显示）；PowerPoint 场景可 spec 加 "animate_charts": true
        if not spec.get("animate_charts"):
            skip_vis = {"appear", "auto"} | ENTRANCE | EXIT_SET
            keep = []
            for eff, item in named_res:
                if item[2] == "chart" and (eff["effect"] in skip_vis or eff.get("out")):
                    print(f"  ~ p{idx} 图表跳过进退场（WPS 兼容）: {eff['target']}")
                    continue
                keep.append((eff, item))
            named_res = keep
        if not named_res and not auto_seq:
            continue
        # 命名效果的容器归属：文字在无字卡片内 → 卡片同批出现（框先、文字紧随）
        cmap_named = find_containers(slide, shapes, {it[1] for _e, it in named_res}, used)
        # 统一序列：按声明顺序展开（byPara/"*" 会一化多），auto 保留为标记
        seq = []
        named_iter = iter(named_res)
        for eff in effects:
            if eff.get("effect") == "auto":
                seq.append(("auto", eff, None))
            else:
                for _ in resolve_targets([eff], shapes, slide):
                    seq.append(("item", None, next(named_iter)))
        # 推进模型（canonical）：开场一个"自动组"承载 auto 与顺排点名（onBegin 引用 mainSeq
        # 随页面启动，效果间用组内累计延时串行）；只有 click 才开新的"点击组"（indefinite 等点击）。
        groups, cur, t, last_start = [], None, 0, 0
        for kind_of, aeff, named_item in seq:
            if kind_of == "auto":
                if cur is None:
                    cur = {"items": [], "auto": True}
                ag = build_auto(slide, shapes, used, aeff, idx, t)
                cur["items"].extend(ag["items"])
                if ag["items"]:
                    last = ag["items"][-1]
                    t = last[2] + last[0].get("dur", 380)
                continue
            eff2, (fe, spid, kind, prg) = named_item
            item = (fe, spid, kind, prg)
            trig = eff2.get("trigger", "afterPrev")
            if cur is not None and cur["items"] and trig == "click":
                groups.append(cur)
                cur, t, last_start = None, 0, 0
            if cur is None:
                cur = {"items": [], "auto": trig != "click"}
            delay = eff2.get("delay", 0)
            dur = eff2.get("dur", 450)
            if trig == "withPrev":
                start = last_start + delay
                last_start = start
            else:
                start = t + delay
                t = start + dur
                last_start = start
            csp = cmap_named.get(spid)
            if csp and csp not in used:
                ce = {"effect": "fade", "trigger": "withPrev", "dur": min(dur, 400)}
                cur["items"].append((ce, (ce, csp, kind, None), start))
                used.add(csp)
            cur["items"].append((eff2, item, start))
        if cur is not None and cur["items"]:
            groups.append(cur)
        if not groups:
            continue
        sld.append(timing_tree(groups))
        n_an += 1
    prs.save(out or path)
    print(f"animated: {out or path} — transition×{n_tr}, timing×{n_an}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(2)
    out = None
    if "--out" in sys.argv:
        out = sys.argv[sys.argv.index("--out") + 1]
    apply(sys.argv[1], sys.argv[2], out)
