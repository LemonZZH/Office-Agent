# 架构与设计决策

## 定位

Office-Agent 是一个运行在 ZCode 中的 PPT 生成 Agent。它不是模板库，而是一套**工作流 + 设计系统 + 记忆**的组合：

- **多模态摄入** — 一段话、一份文档（docx/pdf/md/xlsx）、一张图片（内容照片或风格参考）都可以是起点。
- **可再生的 deck 工作区** — 每份 PPT 始终保留三类源文件（大纲 / 设计 token / 构建脚本），对话式改进 = 改源文件后重建，可 diff、可验证、可接管。
- **记忆** — 跨 deck 沉淀用户偏好与设计反馈，越用越懂用户。

## 双状态模型

```
memory/（跨 deck，用户级）          decks/<name>/（单 deck 状态）
├── profile.md  偏好画像            ├── brief.md    需求纪要+风格决策
├── feedback.md 反馈→设计规则       ├── outline.md  逐页大纲（内容事实源）
└── decks.md    作品索引            ├── theme.js    设计 token
                                   ├── build.js    构建脚本
                                   ├── assets/     配图
                                   └── <Deck>.pptx 产物
```

为什么不在 .pptx 上直接改？二进制编辑每次都要逆向推断设计意图，改动不可复现；而"改源码再重建"让任何一次修改都可审查、可回滚，新会话读 `brief.md + outline.md + theme.js` 就能无缝接手。需要编辑用户给定的既有 .pptx 且要求保留原模板时，走官方 pptx 技能的模板继承流程，产物仍落回工作区。

## 主流程

```
读记忆 → 多模态摄入 → 风格自适应 → 大纲 → 设计token → pptxgenjs构建
   ↑                                                        ↓
记忆更新 ← 交付 ← 视觉验收 ← 渲染PNG ← qa_check.py ←────────┘
                        （对话式迭代回到大纲/token层，循环重建）
```

## 分层职责

| 层 | 位置 | 职责 |
|---|---|---|
| 派发层 | `.zcode/agents/ppt-designer.md` | 识别 PPT 类请求、约束子代理行为 |
| 流程层 | `.agents/skills/ppt-designer/SKILL.md` + references/ | 摄入/风格/大纲/QA/记忆协议/风格预设 |
| 机制层 | 官方 presentations 插件 pptx 技能 | pptxgenjs API、避坑清单、渲染、模板继承（运行时依赖，不入库） |
| 工具层 | `scripts/` | qa_check.py 溢出重叠检查、render.sh 渲染管线 |

本项目技能**不重复教学** pptxgenjs 用法——那是官方 pptx 技能的职责；本项目贡献的是其之上的场景化设计系统与状态管理。

## 风格体系

三种风格预设（汇报 / 路演 / 高端）是设计起点而非门槛：Agent 按场景自动选择并说明理由，用户可用自然语言覆盖任意细节。高端风格额外启用逐案定制（先出设计方向再动工）与 dreamina AI 配图（消耗积分，先告知）。

设计 token 统一遵循 BACKGROUND → PRIMARY → ACCENT 三角色模型，见 `references/design-tokens.md`。

## 记忆协议（摘要）

- 写入时机：明确偏好 → profile.md；表扬/批评 → feedback.md 提炼为规则；交付 → decks.md 索引。
- 合并而非覆盖；冲突时最新反馈优先。
- 完整协议见 `references/memory.md`。memory/ 与 decks/ 含个人数据，.gitignore 排除，不入库。

## 扩展路线

- Lark Slides / Google Slides 输出适配器
- 从 xlsx 数据文件直接驱动数据页
- 风格参考图驱动的模板生成（"照这个风格做"的深化）
- 多语言 deck
