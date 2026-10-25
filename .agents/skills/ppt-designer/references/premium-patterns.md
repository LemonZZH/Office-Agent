# 高端样式模式库（premium patterns）

来源：Apple Keynote 设计法则（LinkedIn/Inc./Forres 拆解）、iSlide 高级感教程、
知乎大字封面、Canva 高级技巧、编辑部衬线字体研究（Studio2am 等）。
一句话总纲：**高级感 = 极端克制 + 极端对比**。一个页面只允许一个主角。

## 1. 一页一大思想（Apple 法则核心）

- 每页只承载**一个**论点/数字/图像；其余全部删掉。
- 章节页/金句页：整页只有一个短语（60–96pt）+ 品牌签名，别的一概不放。
- " striking numbers"：把最有冲击力的数字放大到 90–140pt 独占半屏，标签 12–14pt 弱化在下方。

## 2. 字阶与字体对比（编辑部式排版）

| 角色 | 字号 | 字体 | 用法 |
|---|---|---|---|
| Display | 88–140pt | **衬线**（CJK: 宋体系；Latin: 高对比 serif） | 封面标题、章节序号、大数字——只此一处用衬线 |
| Title | 30–36pt | 无衬线 Bold | 内容页标题 |
| Body | 13–16pt | 无衬线 | 正文 |
| Caption | 12pt | 无衬线 | 图注/来源 |

- 对比产生高级感：Display 与 Body 至少差 4 个字阶；同页字号种类 ≤3。
- 大字排版不等于居中大字——用**大小对比、色块分割、局部留白**造层次（知乎/CSDN 共识）。
- 深底正文用白色但**非纯白**（#F2F5FA 级），浅底正文用黑但非纯黑。

## 3. 留白与密度

- 高端档页边距 ≥0.9"，单页内容组 ≤4 组；能不画盒子就不画（hairline + 对齐分组）。
- 卡片若必须用：**玻璃质感**——CARD 色 fill + transparency 25–35（透出渐变背景），1px 亮线描边；禁实心色块堆叠。

## 4. 背景与光

- 深色 = 有层次的深：**对角渐变 + 暗角 + 微颗粒**（make_bg.py 自动生成），禁纯色平铺。
- 单一强调色 ≤10% 面积（金/珊瑚），用在：大数字、一条 hairline、一个图标。
- 满版图片必须压 55–75% 深色渐变蒙版再上字（Canva 双色调技巧同理）。

## 5. 图片与实机截图

- 截图统一处理：细 1px 亮边框 + 图注弱化；两图并排保持同高对齐。
- 能用原生图表就不截图；截图只证明"真实存在"。

## 6. 动效叙事化（内容先行的编排）

- 编排三问（写 anim.json 前必答）：这页的**论点**是什么？观众**先看什么**？什么需要**后揭晓**？
- 数字页：数字 zoomIn 先行 → 标签/正文跟上 →（可选）pulse 一次。对比页：先左后右。流程页：按序号逐段。结论页：论据退场 → 论点 fade。
- 切换全场统一 fade/zoom；花哨滤镜只属于演示页。

## 来源

- [The Apple Keynote formula: 5 design principles](https://www.linkedin.com/posts/mayursheth_pitchdeckdesign-pitchdeck-storytelling-activity-7414555643463749632-SCwB)
- [8 Presentation Techniques from Apple's WWDC Keynote (Forbes)](https://www.forbes.com/sites/carminegallo/)
- [iSlide：让 PPT 设计具有视觉高级感](https://www.islide.cc/blog/202)
- [知乎：如何快速做出高端大气的 PPT 封面](https://zhuanlan.zhihu.com/p/401148081)
- [CSDN：从海量资源到莫兰迪高级感](https://bbs.csdn.net/weixin_30000735/article/details/100514893)
- [Canva：20 个让 PPT 更高级的设计技巧](https://www.canva.cn/learn/creative-decks-presentation)
- [Studio2am：8 Editorial Serif Fonts](https://studio2am.co/blogs/news/8-editorial-serif-fonts-with-enough-contrast-to-own-a-masthead)
- [SlideStack：30 Modern and Clean Fonts](https://slidestack.com/blog/30-modern-and-clean-fonts-every-presentation-designer-should-know)
