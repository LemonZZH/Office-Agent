#!/usr/bin/env python3
"""make_motion.py — 由 theme 本地生成循环动效 GIF（缓慢漂移光斑），零网络/零积分。

用法: python3 make_motion.py <theme.js> [--out assets/motion-bg.gif]
      [--frames 48] [--size 480x270] [--mode orbs|particles]

产出: 480×270 无缝循环 GIF（放映时全屏铺底，压暗后作氛围层；两段式正弦漂移保证首尾无缝）。
嵌入: build.js 里 s.addImage({ path: "assets/motion-bg.gif", x:0, y:0, w:W, h:H }) 置于最底层，
      其上压暗色矩形（transparency 30–45）再放标题。GIF 在 PowerPoint/WPS 放映时自动循环。
"""
import json
import math
import os
import subprocess
import sys

from PIL import Image, ImageDraw, ImageFilter

W, H = 480, 270


def load_theme(path):
    out = subprocess.check_output(
        ["node", "-e", f"console.log(JSON.stringify(require('./{os.path.basename(path)}')))"],
        cwd=os.path.dirname(os.path.abspath(path)) or ".")
    return json.loads(out)


def hex2rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def mix(a, b, t):
    return tuple(int(x + (y - x) * t) for x, y in zip(a, b))


def orb_layer(size, color, alpha_max):
    r = Image.radial_gradient("L").resize(size)
    a = r.point(lambda v: int((255 - v) * alpha_max))
    layer = Image.new("RGBA", size, color + (255,))
    layer.putalpha(a)
    return layer


def frame(t, base, orbs):
    img = base.copy().convert("RGBA")
    for (color, alpha, r, x0, y0, x1, y1, phase) in orbs:
        u = (math.sin(2 * math.pi * (t + phase)) + 1) / 2  # 0→1→0 平滑往返
        x = x0 + (x1 - x0) * u
        y = y0 + (y1 - y0) * u
        orb = orb_layer((r, r), color, alpha)
        img.alpha_composite(orb, (int(x - r / 2), int(y - r / 2)))
    return img.convert("RGB")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    theme_path = sys.argv[1]
    out = sys.argv[sys.argv.index("--out") + 1] if "--out" in sys.argv \
        else os.path.join(os.path.dirname(os.path.abspath(theme_path)), "assets", "motion-bg.gif")
    frames_n = int(sys.argv[sys.argv.index("--frames") + 1]) if "--frames" in sys.argv else 48
    os.makedirs(os.path.dirname(out), exist_ok=True)
    T = load_theme(theme_path)
    dark = hex2rgb(T.get("DARK", T["BG"]))
    accent = hex2rgb(T["ACCENT"])
    primary = hex2rgb(T["PRIMARY"])

    # 底：对角渐变
    base = Image.new("RGB", (W, H))
    px = base.load()
    for y in range(H):
        for x in range(W):
            t = (x / W + y / H) / 2
            px[x, y] = mix(mix(dark, (0, 0, 0), 0.12), mix(dark, (255, 255, 255), 0.04), t)

    orbs = [
        (accent, 0.34, 360, 0.15 * W, 0.25 * H, 0.55 * W, 0.55 * H, 0.00),
        (primary, 0.30, 420, 0.80 * W, 0.70 * H, 0.45 * W, 0.40 * H, 0.35),
        (accent, 0.16, 260, 0.50 * W, 0.85 * H, 0.85 * W, 0.25 * H, 0.62),
    ]
    fr = []
    for i in range(frames_n):
        img = frame(i / frames_n, base, orbs)
        fr.append(img.filter(ImageFilter.GaussianBlur(1.2)))
    fr = [f.convert("P", palette=Image.ADAPTIVE, colors=96) for f in fr]
    fr[0].save(out, save_all=True, append_images=fr[1:], duration=90, loop=0,
               optimize=True)
    print(f"saved: {out} ({os.path.getsize(out) / 1e6:.1f}MB, {frames_n} 帧, 无缝循环)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
