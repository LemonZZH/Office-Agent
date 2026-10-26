#!/usr/bin/env python3
"""make_bg.py — 由 theme.js 本地生成质感素材（PIL，无需 AI 积分）：渐变背景 + 辉光。

用法: python3 make_bg.py <theme.js> [--out <assets_dir>]   默认输出到 theme 同级 assets/

产出:
  bg-dark.jpg   1920×1080 深色渐变（DARK 基调纵向渐变 + ACCENT/PRIMARY 角部微光 + 细颗粒）
  bg-light.jpg  浅色渐变（BG 基调，极克制）——浅色内容页/汇报档封面用
  glow.png      1024×1024 ACCENT 径向辉光（透明底），垫在大数字/标题后面

build.js 用法:
  s.background = { path: "assets/bg-dark.jpg" };                    // 封面/深色页
  s.addImage({ path: "assets/glow.png", x, y, w, h, transparency: 40 });  // 先辉光后文字
"""
import json
import os
import subprocess
import sys

from PIL import Image, ImageDraw, ImageFilter

W, H = 1920, 1080
G = 1024


def load_theme(path):
    out = subprocess.check_output(
        ["node", "-e", f"console.log(JSON.stringify(require('./{os.path.basename(path)}')))"],
        cwd=os.path.dirname(os.path.abspath(path)) or ".")
    return json.loads(out)


def hex2rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def mix(c1, c2, t):
    return tuple(int(a + (b - a) * t) for a, b in zip(c1, c2))


def clamp(c):
    return tuple(max(0, min(255, v)) for v in c)


def vgrad(c_top, c_bottom, w=W, h=H):
    img = Image.new("RGB", (w, h))
    px = img.load()
    for y in range(h):
        c = mix(c_top, c_bottom, y / h)
        for x in range(w):
            px[x, y] = c
    return img


def radial_alpha(size, k=1.0):
    """中心亮→边缘 0 的 L 面（radial_gradient 是中心黑，反转并缩放）。"""
    r = Image.radial_gradient("L").resize((size, size))
    return r.point(lambda v: int((255 - v) * k))


def corner_tint(img, color, corner, alpha_max):
    tint = Image.new("RGB", img.size, color)
    a = radial_alpha(1600, alpha_max / 255.0)
    mask = Image.new("L", img.size, 0)
    pos = {"tr": (img.width - 1600, -540), "bl": (-540, img.height - 1600),
           "tl": (-540, -540), "br": (img.width - 1600, img.height - 1600)}[corner]
    mask.paste(a, pos)
    img.paste(tint, (0, 0), mask)


def grain(img, sigma=16, strength=0.045):
    n = Image.effect_noise(img.size, sigma).convert("RGB")
    return Image.blend(img, n, strength)


def vgrad(c_top, c_bottom, w=W, h=H, diagonal=True, vignette=0.22, streak=True):
    """对角渐变 + 暗角 + 对角光带（小图逐像素后放大，快且平滑）。"""
    sw, sh = 480, 270
    small = Image.new("RGB", (sw, sh))
    px = small.load()
    for y in range(sh):
        for x in range(sw):
            t = (x / sw + y / sh) / 2 if diagonal else y / sh
            px[x, y] = mix(c_top, c_bottom, t)
    img = small.resize((w, h), Image.BILINEAR)
    if streak:  # 对角柔光带
        band = Image.new("L", (w, h), 0)
        bd = ImageDraw.Draw(band)
        bd.polygon([(int(w * 0.55), 0), (int(w * 0.8), 0), (int(w * 0.25), h), (int(w * 0.05), h)],
                   fill=26)
        band = band.filter(ImageFilter.GaussianBlur(160))
        img = Image.composite(Image.new("RGB", (w, h), (255, 255, 255)), img,
                              band.point(lambda v: v // 3))
    if vignette:  # 暗角
        r = Image.radial_gradient("L").resize((w, h)).point(
            lambda v: int((v / 255) ** 1.8 * vignette * 255))
        img = Image.composite(Image.new("RGB", (w, h), (0, 0, 0)), img, r)
    return img


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    theme_path = sys.argv[1]
    out_dir = sys.argv[sys.argv.index("--out") + 1] if "--out" in sys.argv \
        else os.path.join(os.path.dirname(os.path.abspath(theme_path)), "assets")
    os.makedirs(out_dir, exist_ok=True)
    T = load_theme(theme_path)

    dark = hex2rgb(T.get("DARK", T["BG"]))
    accent = hex2rgb(T["ACCENT"])
    primary = hex2rgb(T["PRIMARY"])

    # bg-dark：对角渐变 DARK→更深 + 暗角 + 光带，角部 ACCENT/PRIMARY 微光，颗粒
    bg = vgrad(clamp(mix(dark, (0, 0, 0), 0.14)), clamp(mix(dark, (255, 255, 255), 0.06)),
               vignette=0.24)
    corner_tint(bg, accent, "tr", 26)
    corner_tint(bg, primary, "bl", 20)
    bg = grain(bg)
    bg.save(os.path.join(out_dir, "bg-dark.jpg"), quality=88)

    # bg-light：BG→略沉，PRIMARY 一角极淡
    lbg = vgrad(hex2rgb(T["BG"]), clamp(mix(hex2rgb(T["BG"]), (0, 0, 0), 0.05)))
    corner_tint(lbg, primary, "tr", 12)
    lbg = grain(lbg, sigma=10, strength=0.03)
    lbg.save(os.path.join(out_dir, "bg-light.jpg"), quality=88)

    # glow：ACCENT 径向辉光（透明底）
    glow = Image.new("RGBA", (G, G), (0, 0, 0, 0))
    a = radial_alpha(G, 1.0).point(lambda v: int(v * 0.42))
    layer = Image.new("RGBA", (G, G), accent + (255,))
    layer.putalpha(a)
    glow = Image.alpha_composite(glow, layer).filter(ImageFilter.GaussianBlur(30))
    glow.save(os.path.join(out_dir, "glow.png"))

    # —— SVG 矢量版（无限清晰、体积极小；PowerPoint 2016+/WPS 2019+ 原生渲染，
    #     pptxgenjs 自动内嵌 PNG 兜底给旧播放器）——
    d_hex, p_hex, a_hex = T.get("DARK", T["BG"]), T["PRIMARY"], T["ACCENT"]

    def rgbhex(rgb):
        return "".join(f"{c:02X}" for c in rgb[:3])

    def svg_bg():
        darker = rgbhex(mix(hex2rgb(d_hex), (0, 0, 0), 0.14))
        return f"""<svg xmlns="http://www.w3.org/2000/svg" width="1920" height="1080" viewBox="0 0 1920 1080">
  <defs>
    <linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#{d_hex}"/><stop offset="1" stop-color="#{darker}"/>
    </linearGradient>
    <radialGradient id="v" cx="0.5" cy="0.5" r="0.75">
      <stop offset="0.5" stop-color="#000" stop-opacity="0"/>
      <stop offset="1" stop-color="#000" stop-opacity="0.3"/>
    </radialGradient>
    <radialGradient id="tr" cx="1" cy="0" r="0.8">
      <stop offset="0" stop-color="#{a_hex}" stop-opacity="0.16"/>
      <stop offset="0.6" stop-color="#{a_hex}" stop-opacity="0"/>
    </radialGradient>
    <radialGradient id="bl" cx="0" cy="1" r="0.8">
      <stop offset="0" stop-color="#{p_hex}" stop-opacity="0.12"/>
      <stop offset="0.6" stop-color="#{p_hex}" stop-opacity="0"/>
    </radialGradient>
  </defs>
  <rect width="1920" height="1080" fill="url(#g)"/>
  <rect width="1920" height="1080" fill="url(#tr)"/>
  <rect width="1920" height="1080" fill="url(#bl)"/>
  <rect width="1920" height="1080" fill="url(#v)"/>
</svg>"""

    def svg_glow():
        return f"""<svg xmlns="http://www.w3.org/2000/svg" width="1024" height="1024" viewBox="0 0 1024 1024">
  <defs><radialGradient id="gl" cx="0.5" cy="0.5" r="0.5">
    <stop offset="0" stop-color="#{a_hex}" stop-opacity="0.5"/>
    <stop offset="1" stop-color="#{a_hex}" stop-opacity="0"/>
  </radialGradient></defs>
  <rect width="1024" height="1024" fill="url(#gl)"/>
</svg>"""

    open(os.path.join(out_dir, "bg-dark.svg"), "w").write(svg_bg())
    open(os.path.join(out_dir, "glow.svg"), "w").write(svg_glow())
    print(f"assets → {out_dir}/bg-dark.jpg|svg, bg-light.jpg, glow.png|svg")


if __name__ == "__main__":
    sys.exit(main())
