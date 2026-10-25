---
name: ppt-designer
description: 生成与迭代 PPT 的主技能。只要用户提到做 PPT/幻灯片/演示文稿/汇报/路演/deck/slides，给一段话、一份文档、一张图片要求做成 PPT，或要对已有 deck 改内容、加页、换风格、换配色、改数据 —— 即使用户没说"PPT"三个字，也使用本技能。
---

# PPT Designer — 主流程

一句话：**输入可以是任意东西，产出是"源码工作区 + .pptx"，一切改进发生在源码层，一切偏好进入记忆。**

## 0. 每次任务开始（必做）

1. 读 `memory/profile.md`（用户偏好画像；不存在则跳过，收尾时创建）。偏好必须被应用，如用户偏好深色则浅色预设直接排除。
2. 新建 deck 时扫 `memory/decks.md`：有可复用的风格、母题、素材先复用。
3. 确定 deck 工作区 `decks/<deck-name>/`（kebab-case，见名知义）。接手已有 deck 时先读其 `brief.md`。

## 1. 摄入

按 `references/ingestion.md` 处理用户输入（一段话 / 文档 / 图片 / .pptx / 组合），产出写入 `brief.md`：
输入来源与摘要、场景与受众判断、**风格决策及理由**、关键假设清单。

信息不足时：主会话内可一次性问清（最多一轮）；子代理内不提问，按合理假设推进并在交付时说明。

## 2. 风格决策

按场景自动选择风格预设并在 brief.md 说明理由：

- **汇报**（`references/style-report.md`）：内部同步、周月季汇报、总结复盘 → 清晰高效
- **路演**（`references/style-pitch.md`）：融资、对外商务、比赛 → 叙事与数据说服
- **高端**（`references/style-premium.md`）：发布会、品牌、决审投标 → 逐案定制 + AI 配图

用户可用自然语言覆盖任何细节（"深色""活泼点""像 Apple 发布会"），以其为准。
配色/字体从 `references/design-tokens.md` 取预设或按其规则逐案生成。

**风格参考图流程**（用户说"照这个风格/参考这张海报"时）：
1. Read 视觉判读参考图：行业气质、字体感受（衬线/无衬线）、版式密度、候选 motif。
2. `python3 scripts/theme_from_image.py <参考图> --out decks/<name>/theme-from-ref.js` 提取配色 token（背景=边缘众数聚类、PRIMARY=高覆盖高饱和、ACCENT=异色相高饱和或明度变体、深浅自动判定）。
3. 以脚本输出为候选做人工判读微调：PRIMARY/ACCENT 语义是否符合参考图的视觉主角；字体按判读替换 FONT；把 motif 写进 theme。
4. theme 定稿后正常走大纲→构建（`THEME=` 或改名为 theme.js）。

## 3. 大纲 `outline.md`

逐页大纲是**内容的唯一事实源**，格式：

```markdown
# <deck 标题>  （风格: 深空蓝金 | 预计 10 页）
## p1 封面
- 关键信息: 一句话主标题 + 副标题
- 视觉: 全幅深色背景 + 超大标题，AI 配图
## p2 行业痛点
- 关键信息: 三个痛点，核心数据支撑
- 视觉: 三列图标行 + 大数字 87%
...
```

规则：每页恰好一个关键信息，写明视觉形式（图表/图示/大数字/配图/时间线）；数据要么有来源，要么标 `est.`；汇报档 8–12 页、路演档 12–14 页、高端档 10–14 页，用户指定优先。

## 4. 设计 token `theme.js`

```javascript
module.exports = {
  W: 13.33, H: 7.5, M: 0.5,
  BG: "0B1B33", PRIMARY: "3B82F6", ACCENT: "F59E0B",
  TEXT: "F8FAFC", MUTED: "94A3B8",
  SERIES: ["3B82F6", "1D4ED8", "93C5FD"],   // 图表系列色（PRIMARY 梯度）
  FONT_TITLE: "微软雅黑", FONT_BODY: "微软雅黑",
  MOTIF: "oversized-number",                 // 签名母题
};
```

禁止在 build.js 里散落硬编码颜色；加色只能走 theme.js——**所有**颜色（hairline、网格线、深底辅字、图表网格）都必须来自 theme 常量，这是"一键换肤"的硬前提。build.js 以 `require(process.env.THEME || "./theme.js")` 加载主题、`process.env.OUT` 决定输出名：换皮肤 = `THEME=./theme-dark.js OUT=新名字.pptx node build.js`，构建脚本零改动。

## 5. 构建 `build.js`

**先加载官方 `pptx` 技能**（presentations 插件）——pptxgenjs API、避坑清单（#hex、breakLine、bullet 缩进、阴影 offset、富文本 pPr）以它为准，本技能不重复。

本技能追加要求：`LAYOUT_WIDE`（13.33×7.5"）；所有坐标从 W/H/M 派生；每页一个视觉焦点；图表用原生 addChart（样式跟随 theme.js）；图片 `sizing: cover` 防拉伸；文案入片前对照 outline 的关键信息，不凭空扩写；每页加 speaker notes（演讲者备注写"这页讲什么+强调点"）。

**质感层（对抗"平铺色很廉价"）**：`python3 scripts/make_bg.py <theme.js>` 由当前配色本地生成 `assets/bg-dark.jpg`（深色纵向渐变+角部微光+颗粒）、`bg-light.jpg`（浅色克制渐变）、`glow.png`（ACCENT 径向辉光）。用法：封面/封底/深色页 `s.background = { path: "assets/bg-dark.jpg" }`；大数字/关键标题先垫 `s.addImage({ path: "assets/glow.png", ..., transparency: 35–60 })` 再写字；图表浮于渐变时 chartArea fill 加 `transparency: 100`。高端/路演档默认启用；汇报档至少封面封底用。theme 换肤后重跑 make_bg 再构建。

### 5.5 动效（构建后注入）

pptxgenjs 不支持动画——用 `scripts/animate.py` 做 OOXML 后处理：

1. build.js 中给需要精确控制的元素加 `objectName`（图表、KPI 大数字、卡片等）。
2. 写 `anim.json`：`transition`（全场默认切换）+ `transitions`（按页覆盖）+ `slides`（逐页动效）。`target: "*"` = 该页全部形状按位置排序级联；`target: objectName` = 精准指向。
3. `python3 scripts/animate.py deck.pptx anim.json`（原地注入；rebuild 后需重跑）。

效果词汇表（trigger：afterPrev / withPrev / click；通用参数：dur / delay / stagger / autoRev / repeat / accel / decel）：

| 家族 | 效果 | 用途 |
|---|---|---|
| 基础进场 | fade / appear / wipe(dir) / flyIn(dir) / zoomIn(from%) | 万金油；zoomIn 适合封面标题 |
| 滤镜进场 | wheel(1/2/4/8) / checkerboard / blinds / circle / diamond / dissolveIn / strips / wedge / randomBar | 视觉花样；正式商务慎用复古滤镜（棋盘/百叶窗） |
| 退场 | fadeOut / zoomOut / wipeOut / flyOut（滤镜加 `"out": true` 亦成退场） | 讲完即走，给结论留白 |
| 强调 | pulse / grow / shrink（scale%）/ spin(deg) / colorTo(RRGGBB) | 关键数字/图标跳出，不动可见性 |
| 路径 | move(dx, dy)（slide 宽高小数）+ autoRev | 表达流向、往返示意 |
| 图表构建 | 效果上加 `"chart": "category"|"series"` | 柱图逐根升起 / 折线逐条画出——数据页首选 |
| 文字构建 | `"byPara": true`（target 须为具体名称） | 要点逐条进场，可配合 click 控制演讲节奏 |
| 切换 | fade / push / wipe / zoom / cover / split / wheel / comb / dissolve / checkerboard / blinds / strips / wedge / circle / random… | 全场统一，按页覆盖 |

动效纪律：**动效为叙事服务，不是炫技**。汇报档克制（fade/wipe 级联 + 图表按类构建）；路演档数据页 wipe 强化"增长感"；高端档慢 fade（600–900ms、大 stagger）；封底页留空。同一形状可叠加（进场→强调→退场）。滤镜家族与强调类默认只用于演示/氛围页。参考完整示例：`decks/effects-showcase/`。

**WPS 兼容性**：图表挂进退场动画在 WPS 全屏放映不触发（图表被隐藏）——animate.py 默认跳过图表的进退场（图表随页面显示，日志有提示）；确需 PowerPoint 的按类/按系列构建时，anim.json 顶层加 `"animate_charts": true` 并接受 WPS 降级为静态显示。

## 6. QA 门（动效注入后必跑）

1. `python3 scripts/qa_check.py <deck>.pptx` — 溢出/越界/重叠/占位符。有问题先修再进下一关。
   - 预览环境注意（macOS）：中文渲染空白 → 把系统 CJK 字体拷入 LibreOffice 自带字体目录（见 render.sh 头注释）；**粗体中的半角 `+` 会回退成 `✦`** → 粗体文案用中文表述（"同比增长 23%"），半角 + 只出现在常规体中。
2. `bash scripts/render.sh <deck>.pptx` 渲染 PNG。
3. 派发 `presentations:visual-judge` 子代理验收全部页面，按其 verdict 修复后**复检**。汇报/路演 1 轮通过即可；高端档循环至 pass（上限 3 轮，仍不过则如实告知遗留问题）。visual-judge 不可用时（如子代理启动报 provider 错误），亲自逐页检查渲染图，按同等标准（溢出/重叠/对比度/空洞/AI 味装饰）给出结论。
4. 交付前：工作区内只留源文件、assets、pptx 与 render/；无 TODO 残留。

## 7. 交付

报告：产物路径、页数与结构、风格说明、关键假设、遗留问题。然后按 `references/memory.md` 更新记忆（decks.md 必更）。

## 8. 对话式迭代（收到修改意见时）

判断改动层级 → 只改对应层 → 重建 → 按 §6 复检（qa_check.py 必跑；仅布局级大改才重跑视觉验收）：

| 用户说什么 | 改哪 | 示例 |
|---|---|---|
| 内容级 | outline.md 对应页 → 同步 build.js | "第3页数据换成 Q4 的" |
| 风格级 | theme.js | "整体换深色""字太小" |
| 结构级 | outline.md + build.js | "加两页竞品分析""删掉团队页" |
| 素材级 | assets/ + build.js | "封面图换抽象一点的" |
| 需求级 | brief.md 重新决策，再走上面四层 | "改成给投资人看的版本" |

用户口头表扬/批评设计时，除执行修改外，按 memory.md 沉淀为规则。

## 9. 收尾（每次任务结束前）

按 `references/memory.md` 协议更新记忆文件，确保下次会话可无缝接手。
