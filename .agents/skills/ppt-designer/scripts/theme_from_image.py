#!/usr/bin/env python3
"""theme_from_image.py — 从风格参考图提取设计 token，生成可直接 THEME= 使用的 theme.js。

用法: python3 theme_from_image.py <ref.jpg|png|...> [--out theme-from-ref.js]

提取策略（色彩学规则，非玄学）:
  BG/DARK  参考图四周边框带（15%）的众数色——海报/封面的底色几乎总在边缘
  PRIMARY  覆盖率 >5% 且饱和度最高、色相不同于背景的色（主题色）
  ACCENT   饱和度高、覆盖率 0.5–5% 的最突出色（强调色）；不足则取 PRIMARY 高饱和变体
  TEXT/MUTED  背景亮→取最暗中性色系；背景暗→取最亮中性色系
  SERIES   PRIMARY 的明度梯度三档（design-tokens.md 规则：加色相不加新色）
  深浅判定  背景相对亮度 <0.35 → 深色 deck（BG=DARK=底色）；否则浅色三明治

输出: JSON 到 stdout（供 agent 判读微调）+ 可选 --out 直接写 theme.js。
注意: 色值是候选，字体/motif 仍需 agent 视觉判读参考图后补进 theme（见 SKILL.md §2）。
"""
import colorsys
import json
import sys

from PIL import Image


def hex_of(rgb):
    return "".join(f"{c:02X}" for c in rgb[:3])


def hsl(rgb):
    r, g, b = (c / 255 for c in rgb[:3])
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    return h * 360, s, l


def lum(rgb):
    r, g, b = (c / 255 for c in rgb[:3])
    lin = tuple(c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4 for c in (r, g, b))
    return 0.2126 * lin[0] + 0.7152 * lin[1] + 0.0722 * lin[2]


def shift_l(rgb, dl):
    h, s, l = hsl(rgb)
    l = max(0, min(1, l + dl))
    r, g, b = colorsys.hls_to_rgb(h / 360, l, s)
    return tuple(int(c * 255) for c in (r, g, b))


def hue_dist(a, b):
    d = abs(a - b) % 360
    return min(d, 360 - d)


def palette(img, n=12):
    q = img.quantize(colors=n, method=Image.Quantize.MEDIANCUT)
    counts = sorted(q.getcolors(200 * 200 * 4), reverse=True)  # [(count, palette_index)]
    pal = q.getpalette()
    out = []
    for cnt, idx in counts:
        rgb = (pal[idx * 3], pal[idx * 3 + 1], pal[idx * 3 + 2])
        out.append((cnt, rgb))
    return out


def bg_mode(border_img):
    """边缘带背景色：按 //16 桶聚类取最大簇的平均（不经过量化，避免深色被并入纯黑）。"""
    buckets = {}
    px = border_img.getdata()
    for r, g, b in px:
        k = (r // 16, g // 16, b // 16)
        c = buckets.get(k, (0, 0, 0, 0))
        buckets[k] = (c[0] + 1, c[1] + r, c[2] + g, c[3] + b)
    cnt, sr, sg, sb = max(buckets.values(), key=lambda v: v[0])
    return cnt, (sr // cnt, sg // cnt, sb // cnt)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    path = sys.argv[1]
    out_js = sys.argv[sys.argv.index("--out") + 1] if "--out" in sys.argv else None

    img = Image.open(path).convert("RGB")
    small = img.resize((200, 200))
    total = 200 * 200

    # —— 背景: 边缘带众数 ——
    w, h = small.size
    bw, bh = int(w * 0.15), int(h * 0.15)
    mask = Image.new("RGB", (w, h), (0, 0, 0))
    mask.paste(small.crop((0, 0, w, bh)), (0, 0))
    mask.paste(small.crop((0, h - bh, w, h)), (0, h - bh))
    mask.paste(small.crop((0, 0, bw, h)), (0, 0))
    mask.paste(small.crop((w - bw, 0, w, h)), (w - bw, 0))
    border_px = [px for px, m in zip(small.getdata(), mask.getdata()) if m != (0, 0, 0)]
    border_img = Image.new("RGB", (len(border_px), 1))
    border_img.putdata(border_px)
    bg_cnt, bg = bg_mode(border_img)
    dark = lum(bg) < 0.35

    # —— 候选分类 ——
    cols = palette(small, 12)
    bg_h, bg_s, bg_l = hsl(bg)
    primary = None
    accent = None
    distinct = []  # (饱和度, rgb) 异色相候选，供 ACCENT 去重
    neutrals = []
    for cnt, rgb in cols:
        frac = cnt / total
        h_, s_, l_ = hsl(rgb)
        if hex_of(rgb) == hex_of(bg):
            continue
        neutral = s_ < 0.18 or (hue_dist(h_, bg_h) < 15 and abs(l_ - bg_l) < 0.15)
        if neutral:
            neutrals.append((cnt, rgb))
            continue
        if frac > 0.05 and (primary is None or s_ > hsl(primary)[1]):
            primary = rgb
        elif s_ > 0.30:
            distinct.append((s_, cnt, rgb))
    if primary is None:  # 单色系参考图：用最深/最亮的非背景色
        cands = [c for c in cols if hex_of(c[1]) != hex_of(bg)]
        primary = (max(cands, key=lambda c: hsl(c[1])[1]) if cands else bg)
    # ACCENT：优先「高饱和 + 与 PRIMARY 异色相」的候选；双色参考图则退化为 PRIMARY 明度变体
    ph = hsl(primary)[0]
    for s_, cnt, rgb in sorted(distinct, reverse=True):
        if hue_dist(hsl(rgb)[0], ph) >= 25:
            accent = rgb
            break
    if accent is None:
        accent = shift_l(primary, 0.18 if dark else -0.22)

    # —— 中性色：排除背景同族（同色相且亮度接近），否则深底会把深色当正文 ——
    text_neutrals = [rgb for _c, rgb in neutrals
                     if not (hue_dist(hsl(rgb)[0], bg_h) < 15 and abs(lum(rgb) - lum(bg)) < 0.2)]
    if dark:
        cands = [n for n in text_neutrals if lum(n) > 0.5]
        text = max(cands, key=lum) if cands else (241, 239, 234)
        muted = shift_l(text, -0.20)
        dim = shift_l(text, -0.26)
    else:
        cands = [n for n in text_neutrals if lum(n) < 0.5]
        text = min(cands, key=lum) if cands else (30, 41, 59)
        muted = shift_l(text, 0.30)
        dim = muted

    T = {
        "W": 13.33, "H": 7.5, "M": 0.5,
        "BG": hex_of(bg if not dark else bg),
        "DARK": hex_of(bg if dark else shift_l(bg, -0.30)),
        "CARD": hex_of(shift_l(bg, 0.05 if dark else -0.02)),
        "LINE": hex_of(shift_l(bg, 0.14 if dark else -0.10)),
        "GRID": hex_of(shift_l(bg, 0.08 if dark else -0.06)),
        "PRIMARY": hex_of(primary),
        "ACCENT": hex_of(accent),
        "ACCENT2": hex_of(shift_l(accent, -0.12)),
        "TEXT": hex_of(text),
        "MUTED": hex_of(muted),
        "DIM": hex_of(dim),
        "ON_DARK": hex_of(text if dark else (255, 255, 255)),
        "SERIES": [hex_of(primary), hex_of(shift_l(primary, -0.12)), hex_of(shift_l(primary, 0.18))],
        "FONT": "微软雅黑",
    }
    meta = {
        "deck_is_dark": dark,
        "bg_coverage": round(bg_cnt / total, 2),
        "note": "色值为候选；请结合参考图视觉判读：确认 PRIMARY/ACCENT 语义、替换 FONT/motif，再入库 theme.js",
    }
    print(json.dumps({"tokens": T, "analysis": meta}, ensure_ascii=False, indent=2))
    if out_js:
        with open(out_js, "w") as f:
            f.write("// theme-from-image — 由参考图提取: %s\n" % path)
            f.write("// %s\n" % meta["note"])
            f.write("module.exports = ")
            f.write(json.dumps(T, ensure_ascii=False, indent=2))
            f.write(";\n")
        print(f"written: {out_js}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
