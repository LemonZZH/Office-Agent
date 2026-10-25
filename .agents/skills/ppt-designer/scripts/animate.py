#!/usr/bin/env python3
"""animate.py — 给构建好的 .pptx 注入动效（pptxgenjs 不支持动画，本脚本做 OOXML 后处理）。

用法: python3 animate.py <deck.pptx> <anim.json> [--out <output.pptx>]
  默认原地覆盖（build.js 重建会重新生成，动效需重跑）。

anim.json 结构:
{
  "transition": {"type": "fade", "speed": "med"},      // 全部页切换动效 fade|push|wipe|cut|dissolve
  "slides": {
    "2":  [ {"target": "*", "effect": "fade", "trigger": "afterPrev",
             "dur": 400, "stagger": 120} ],            // "*": 该页全部顶层形状按位置排序级联
    "4":  [ {"target": "kpi-growth", "effect": "fade", "trigger": "afterPrev", "dur": 400},
            {"target": "chart-revenue", "effect": "wipe", "dir": "right",
             "trigger": "afterPrev", "dur": 700} ]     // target = build.js 里的 objectName
  }
}
effect: fade | wipe(dir=left/right/up/down) | flyIn(dir=top/bottom/left/right) | appear | fadeOut(退场)
trigger: afterPrev(自动接续) | withPrev(同时) | click(点击触发，新组)
"""
import json
import sys

from pptx import Presentation
from pptx.oxml.ns import qn
from lxml import etree

P = "http://schemas.openxmlformats.org/presentationml/2006/main"

FLY_FROM = {"top": (1, "0-#ppt_h/2", "y"), "bottom": (4, "1+#ppt_h/2", "y"),
            "left": (8, "0-#ppt_w/2", "x"), "right": (2, "1+#ppt_w/2", "x")}
WIPE_FILTER = {"right": "wipe(right)", "left": "wipe(left)", "down": "wipe(down)", "up": "wipe(up)"}
WIPE_SUBTYPE = {"right": 8, "left": 2, "down": 1, "up": 4}
TRANSITION_SPEED = {"slow": "slow", "med": "med", "fast": "fast"}


def make(parent, tag, attrib=None, *children):
    """在 parent 下创建命名元素（不追加，仅构造则 parent=None）。"""
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


def resolve_targets(effects, shapes):
    """展开 target；"*" 按位置排序（上→下、左→右）。同一形状可被多次动效（进场+退场）。"""
    resolved = []
    for eff in effects:
        if eff["target"] == "*":
            ordered = sorted(shapes, key=lambda s: (s[4], s[3]))
            for i, s in enumerate(ordered):
                e2 = dict(eff)
                if "stagger" in eff:
                    e2["delay"] = eff.get("delay", 0) + i * eff["stagger"]
                resolved.append((e2, s[0], s[2]))
        else:
            m = [s for s in shapes if s[1] == eff["target"]]
            if not m:  # 模糊匹配兜底
                m = [s for s in shapes if eff["target"] in s[1]]
            if not m:
                print(f"  ! 未找到目标形状: {eff['target']}（跳过）")
                continue
            resolved.append((dict(eff), m[0][0], m[0][2]))
    return resolved


def tgt(spid):
    t = etree.Element(qn("p:tgtEl"))
    make(t, "p:spTgt", {"spid": spid})
    return t


def bhvr(ids, spid, attr=None, dur=None, extra_delay=None, additive=None):
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
    b.append(tgt(spid))
    if attr:
        al = make(b, "p:attrNameLst")
        make(al, "p:attrName").text = attr
    return b


def set_vis(ids, spid, val, delay=0):
    s = etree.Element(qn("p:set"))
    b = bhvr(ids, spid, attr="style.visibility", dur=1, extra_delay=delay if val == "hidden" else None)
    s.append(b)
    make(s, "p:to").append(make(None, "p:strVal", {"val": val}))
    return s


def anim_effect(ids, spid, transition, filt, dur):
    ae = etree.Element(qn("p:animEffect"))
    ae.set("transition", transition)
    ae.set("filter", filt)
    ae.append(bhvr(ids, spid, dur=dur))
    return ae


def anim_pos(ids, spid, attr, from_v, to_v, dur):
    an = etree.Element(qn("p:anim"))
    an.set("calcmode", "lin")
    an.set("valueType", "num")
    an.append(bhvr(ids, spid, attr=attr, dur=dur, additive="base"))
    tav = make(an, "p:tavLst")
    for tm, v in ((0, from_v), (100000, to_v)):
        t = make(tav, "p:tav", {"tm": tm})
        val = make(t, "p:val")
        make(val, "p:strVal", {"val": v})
    return an


def effect_par(ids, eff, spid, kind, start, is_first_of_click):
    e = eff["effect"]
    dur = eff.get("dur", 450)
    dirn = eff.get("dir", "right")
    if e == "fade":
        preset = (10, "entr", 0)
    elif e == "appear":
        preset = (1, "entr", 0)
    elif e == "wipe":
        preset = (22, "entr", WIPE_SUBTYPE[dirn])
    elif e == "flyIn":
        preset = (2, "entr", FLY_FROM[dirn][0])
    else:
        preset = (10, "exit", 0)
    pid, pcls, psub = preset
    node_type = ("clickEffect" if is_first_of_click and eff["trigger"] == "click"
                 else "withEffect" if eff["trigger"] == "withPrev" else "afterEffect")
    par = etree.Element(qn("p:par"))
    ctn = make(par, "p:cTn", {"id": ids.next(), "presetID": pid, "presetClass": pcls,
                              "presetSubtype": psub, "fill": "hold", "grpId": 0,
                              "nodeType": node_type})
    st = make(ctn, "p:stCondLst")
    make(st, "p:cond", {"delay": start})
    ch = make(ctn, "p:childTnLst")
    if e in ("fade", "appear", "wipe", "flyIn"):
        ch.append(set_vis(ids, spid, "visible"))
        if e == "fade":
            ch.append(anim_effect(ids, spid, "in", "fade", dur))
        elif e == "wipe":
            ch.append(anim_effect(ids, spid, "in", WIPE_FILTER[dirn], dur))
        elif e == "flyIn":
            sub, from_v, axis = FLY_FROM[dirn]
            other = "ppt_x" if axis == "y" else "ppt_y"
            ch.append(anim_pos(ids, spid, other, f"#{other}", f"#{other}", dur))
            ch.append(anim_pos(ids, spid, f"ppt_{axis}", from_v, f"#ppt_{axis}", dur))
    elif e == "fadeOut":
        ch.append(anim_effect(ids, spid, "out", "fade", dur))
        ch.append(set_vis(ids, spid, "hidden", delay=max(dur - 1, 0)))
    return par, kind == "chart"


def timing_tree(groups):
    """groups: [ {auto: bool, effects: [(eff, spid, kind, start_ms)]} ]"""
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
    blds = []
    for g in groups:
        gpar = make(mch, "p:par")
        gctn = make(gpar, "p:cTn", {"id": ids.next(), "fill": "hold"})
        gst = make(gctn, "p:stCondLst")
        make(gst, "p:cond", {"delay": "indefinite"})
        if g["auto"]:
            cond = make(gst, "p:cond", {"evt": "onBegin", "delay": 0})
            make(cond, "p:tn", {"val": 2})
        gch = make(gctn, "p:childTnLst")
        ipar = make(gch, "p:par")
        ictn = make(ipar, "p:cTn", {"id": ids.next(), "fill": "hold"})
        ist = make(ictn, "p:stCondLst")
        make(ist, "p:cond", {"delay": 0})
        ich = make(ictn, "p:childTnLst")
        for i, (eff, spid, kind, start) in enumerate(g["effects"]):
            par, is_chart = effect_par(ids, eff, spid, kind, start,
                                       is_first_of_click=(i == 0))
            ich.append(par)
            blds.append((spid, is_chart))
    pc = make(seq, "p:prevCondLst")
    make(make(pc, "p:cond", {"evt": "onPrev", "delay": 0}), "p:tgtEl").append(
        etree.Element(qn("p:sldTgt")))
    nc = make(seq, "p:nextCondLst")
    make(make(nc, "p:cond", {"evt": "onNext", "delay": 0}), "p:tgtEl").append(
        etree.Element(qn("p:sldTgt")))
    if blds:
        bl = make(timing, "p:bldLst")
        for spid, is_chart in blds:
            if is_chart:
                bg = make(bl, "p:bldGraphic", {"spid": spid, "grpId": 0})
                make(bg, "p:bldAsOne")
            else:
                make(bl, "p:bldP", {"spid": spid, "grpId": 0})
    return timing


def transition_el(spec):
    t = etree.Element(qn("p:transition"))
    t.set("spd", TRANSITION_SPEED.get(spec.get("speed", "med"), "med"))
    typ = spec.get("type", "fade")
    if typ == "push":
        make(t, "p:push", {"dir": spec.get("dir", "l")})
    elif typ == "wipe":
        make(t, "p:wipe", {"dir": spec.get("dir", "l")})
    elif typ == "cut":
        make(t, "p:cut")
    elif typ == "dissolve":
        make(t, "p:dissolve")
    else:
        make(t, "p:fade")
    return t


def apply(path, spec_path, out=None):
    spec = json.load(open(spec_path))
    prs = Presentation(path)
    slides = list(prs.slides)
    tr = spec.get("transition")
    n_tr = n_an = 0
    for idx, slide in enumerate(slides, 1):
        sld = slide._element
        for tag in ("p:transition", "p:timing"):
            old = sld.find(qn(tag))
            if old is not None:
                sld.remove(old)
        if tr:
            sld.append(transition_el(tr))
            n_tr += 1
        effects = spec.get("slides", {}).get(str(idx))
        if not effects:
            continue
        shapes = shapes_of(slide)
        resolved = resolve_targets(effects, shapes)
        if not resolved:
            continue
        groups, cur, t = [], {"effects": [], "auto": True}, 0
        last_start = 0
        for eff, spid, kind in resolved:
            trig = eff.get("trigger", "afterPrev")
            if trig == "click" and cur["effects"]:
                groups.append(cur)
                cur, t, last_start = {"effects": [], "auto": False}, 0, 0
            if not cur["effects"]:
                cur["auto"] = trig != "click"
            delay = eff.get("delay", 0)
            dur = eff.get("dur", 450)
            if trig == "withPrev":          # 与上一个效果同时起（可加 delay 偏移）
                start = last_start + delay
                last_start = start
            else:                            # afterPrev / click：顺排
                start = t + delay
                t = start + dur
                last_start = start
            cur["effects"].append((eff, spid, kind, start))
        if cur["effects"]:
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
