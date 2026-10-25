# Office-Agent · PPT 生成 Agent

本项目是一个专门生成与迭代 PPT 的 ZCode Agent。用户用任意形式的输入（一段话 / 一份文档 / 一张图片）请求生成 PPT，之后在对话中持续改进。

## 接到 PPT 类请求时

1. **先加载技能** `.agents/skills/ppt-designer/SKILL.md`，严格按其主流程执行。技能的 references/ 与 scripts/ 是流程的一部分，按 SKILL.md 的指引按需读取。
2. 大多数 PPT 任务经 Agent 工具派发给 `ppt-designer` 子代理执行；简单修改（如改一页文字）可在主会话直接完成。

## 目录约定

- `decks/<deck-name>/` — 每份 PPT 一个工作区：`brief.md`（需求纪要+风格决策）、`outline.md`（逐页大纲，内容唯一事实源）、`theme.js`（设计 token）、`build.js`（pptxgenjs 构建脚本）、`assets/`（配图素材）、产物 `.pptx`。**改动 PPT = 改这些源文件后重建**，不直接编辑 .pptx 二进制。
- `memory/` — Agent 的长期记忆（profile.md 偏好画像 / feedback.md 反馈沉淀 / decks.md 作品索引）。**每次任务开始必读 profile.md，任务收尾按协议更新**；读写规则见 `.agents/skills/ppt-designer/references/memory.md`。含个人数据，已被 .gitignore 排除，不要提交。
- `examples/` — 验收样例，供贡献者参考"一份合格的 deck 工作区长什么样"。

## 硬性约束

- 构建一律 pptxgenjs + LAYOUT_WIDE（13.33×7.5"），配色/字体从 theme.js 常量引用，禁止散落硬编码；换肤 = `THEME=./theme-x.js OUT=xxx.pptx node build.js`。
- 构建后按 anim.json 注入动效（`scripts/animate.py`，支持切换/进场/退场），再跑 `qa_check.py`；布局级大改后再跑 render.sh + 视觉验收。
- 渲染依赖 LibreOffice（soffice）与 poppler（pdftoppm）；官方 presentations 插件的 pptx 技能负责库级规范（API、避坑、模板继承），构建前加载它。
- 涉及 dreamina 生图会消耗积分，先告知用户再执行。
- 不复制第三方技能/插件的文本内容进本仓库（开源合规，MIT）。
