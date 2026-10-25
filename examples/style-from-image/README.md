# 风格参考图 → theme（"照这个风格做"）

输入 `reference-poster.jpg`（深青底 + 珊瑚强调的虚构品牌海报）：

```bash
python3 ../../.agents/skills/ppt-designer/scripts/theme_from_image.py reference-poster.jpg --out theme-from-ref.js
THEME=./theme-from-ref.js OUT=Q3经营汇报-深海季风参考.pptx node ../quarterly-report/build.js
```

产出（slide-01 / slide-04 为渲染预览）：同一份汇报构建脚本零改动，自动换上参考图的
深海青背景 0A2D33 + 珊瑚主色 E85F4C + 米白正文——提取的色值见 theme-from-ref.js。
