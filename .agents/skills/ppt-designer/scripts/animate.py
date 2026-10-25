#!/usr/bin/env python3
"""animate.py v2 — 给构建好的 .pptx 注入动效（pptxgenjs 不支持动画，本脚本做 OOXML 后处理）。

用法: python3 animate.py <deck.pptx> <anim.json> [--out <output.pptx>]
  默认原地覆盖（build.js 重建后需重跑）。

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
  进场  fade appear wipe(dir) flyIn(dir) zoomIn(from%)
        wheel wheel2 wheel4 wheel8 checkerboard(dir) blinds(dir) circle diamond
        dissolveIn strips(dir) wedge randomBar(dir)
  退场  fadeOut zoomOut wipeOut(dir) flyOut(dir)
        （任一滤镜效果加 "out": true 即为退场版）
  强调  pulse(scale%) grow(scale%) shrink(scale%) spin(deg) colorTo(RRGGBB)
  路径  move(dx, dy)（slide 宽高小数，y 向下为正）
  构建  "chart": "category"|"series"  图表按类/按系列逐一构建
        "byPara": true                 文本逐段进场（target 须为具体名称）
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
    for g in groups:
        l1 = make(mch, "p:par")
        l1c = make(l1, "p:cTn", {"id": ids.next(), "fill": "hold"})
        l1s = make(l1c, "p:stCondLst")
        make(l1s, "p:cond", {"delay": "indefinite" if not g["auto"] else 0})
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
        resolved = resolve_targets(effects, shapes, slide)
        if not resolved:
            continue
        groups, cur, t, last_start = [], {"items": [], "auto": True}, 0, 0
        for eff, spid, kind, prg in resolved:
            item = (eff, spid, kind, prg)
            trig = eff.get("trigger", "afterPrev")
            if trig == "click" and cur["items"]:
                groups.append(cur)
                cur, t, last_start = {"items": [], "auto": False}, 0, 0
            if not cur["items"]:
                cur["auto"] = trig != "click"
            delay = eff.get("delay", 0)
            dur = eff.get("dur", 450)
            if trig == "withPrev":
                start = last_start + delay
                last_start = start
            else:
                start = t + delay
                t = start + dur
                last_start = start
            cur["items"].append((eff, item, start))
        if cur["items"]:
            groups.append(cur)
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
