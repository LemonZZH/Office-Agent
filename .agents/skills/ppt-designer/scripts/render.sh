#!/usr/bin/env bash
# render.sh — PPTX → PDF → 逐页 PNG（供视觉验收）
# 用法: render.sh <deck.pptx> [输出目录]   默认输出到 pptx 同级 render/
#
# macOS 已知坑：此 LibreOffice 不枚举系统字体，若预览中文全部空白，
# 把任意系统 CJK 字体拷入其自带目录后重试：
#   cp /System/Library/Fonts/STHeiti*.ttc /Applications/LibreOffice.app/Contents/Resources/Fonts/truetype/
set -euo pipefail

PPTX="$1"
[ -f "$PPTX" ] || { echo "not found: $PPTX" >&2; exit 1; }
OUT="${2:-$(dirname "$PPTX")/render}"
mkdir -p "$OUT"
STEM="$(basename "${PPTX%.pptx}")"

SOFFICE="$(command -v soffice || echo /Applications/LibreOffice.app/Contents/MacOS/soffice)"
"$SOFFICE" --headless --convert-to pdf --outdir "$OUT" "$PPTX" >/dev/null
pdftoppm -png -r 150 "$OUT/$STEM.pdf" "$OUT/slide"
echo "rendered: $OUT/slide-*.png"
