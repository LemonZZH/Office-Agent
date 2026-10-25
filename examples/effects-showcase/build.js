// build.js — 动效能力演示 · animate.py v2
const pptxgen = require("pptxgenjs");
const T = require("./theme.js");

const p = new pptxgen();
p.layout = "LAYOUT_WIDE";
p.author = "Office-Agent";
p.title = "动效能力演示";

const F = T.FONT, W = T.W, H = T.H, M = T.M;

function page(s, title, sub, num, note) {
  s.background = { color: T.BG };
  s.addText(title, { x: M, y: 0.42, w: 9, h: 0.6, fontSize: 30, bold: true,
    color: T.TEXT, fontFace: F, margin: 0 });
  if (sub) s.addText(sub, { x: M, y: 1.04, w: 11, h: 0.34, fontSize: 14,
    color: T.MUTED, fontFace: F, margin: 0 });
  s.addText(`${num} / 8`, { x: W - M - 1.2, y: 7.08, w: 1.2, h: 0.3, fontSize: 12,
    color: T.MUTED, fontFace: F, align: "right", margin: 0 });
  if (note) s.addNotes(note);
}

// —— p1 封面 ——
{
  const s = p.addSlide();
  s.background = { color: T.BG };
  s.addText("动效，也是叙事", { objectName: "cover-title", x: 0, y: 2.7, w: W, h: 1.3,
    fontSize: 60, bold: true, color: T.ON_DARK, fontFace: F, align: "center", margin: 0 });
  s.addText("切换 · 滤镜 · 强调 · 路径 · 图表构建 · 逐段文字 · 退场", { objectName: "cover-sub",
    x: 0, y: 4.15, w: W, h: 0.5, fontSize: 18, color: T.DIM, fontFace: F,
    align: "center", margin: 0 });
  s.addNotes("封面：标题缩放进场。切换动效：zoom。");
}

// —— p2 滤镜家族 ——
{
  const s = p.addSlide();
  page(s, "滤镜家族", "同一张卡片，八种进场滤镜", 2,
    "逐卡进场，每张卡片演示一种滤镜；切换动效：push。");
  const fxs = [
    ["fx-wheel", "wheel", "轮辐"], ["fx-checker", "checkerboard", "棋盘"],
    ["fx-blinds", "blinds", "百叶窗"], ["fx-circle", "circle", "圆形展开"],
    ["fx-strips", "strips", "条带"], ["fx-dissolve", "dissolveIn", "溶解"],
    ["fx-wedge", "wedge", "扇形"], ["fx-randombar", "randomBar", "随机条"],
  ];
  fxs.forEach((fx, i) => {
    const x = 0.5 + (i % 4) * 3.1, y = 1.75 + Math.floor(i / 4) * 2.55;
    s.addShape(p.shapes.ROUNDED_RECTANGLE, { objectName: fx[0], x, y, w: 2.9, h: 2.25,
      rectRadius: 0.08, fill: { color: T.CARD }, line: { color: T.LINE, width: 0.75 } });
    s.addText(fx[1], { x: x + 0.25, y: y + 0.45, w: 2.4, h: 0.5, fontSize: 19, bold: true,
      color: T.TEXT, fontFace: F, margin: 0 });
    s.addText(fx[2], { x: x + 0.25, y: y + 1.05, w: 2.4, h: 0.4, fontSize: 14,
      color: T.ACCENT, fontFace: F, margin: 0 });
  });
}

// —— p3 强调效果 ——
{
  const s = p.addSlide();
  page(s, "强调效果", "不改变可见性，让关键元素跳出来", 3,
    "KPI 脉冲、圆环旋转、文字变色、色条生长；切换动效：fade。");
  s.addText("128%", { objectName: "kpi-main", x: 0.9, y: 1.9, w: 3.6, h: 1.5,
    fontSize: 80, bold: true, color: T.ACCENT, fontFace: F, margin: 0 });
  s.addText("NRR 净收入留存", { x: 0.9, y: 3.45, w: 3.6, h: 0.4, fontSize: 15,
    color: T.MUTED, fontFace: F, margin: 0 });
  s.addShape(p.shapes.OVAL, { objectName: "ring-icon", x: 5.5, y: 2.1, w: 1.3, h: 1.3,
    fill: { color: T.BG, transparency: 100 }, line: { color: T.PRIMARY, width: 3 } });
  s.addText("spin · 旋转强调", { x: 4.95, y: 3.55, w: 2.4, h: 0.35, fontSize: 13,
    color: T.MUTED, fontFace: F, align: "center", margin: 0 });
  s.addText("看我变色", { objectName: "txt-color", x: 8.3, y: 2.15, w: 3.6, h: 0.8,
    fontSize: 30, bold: true, color: T.TEXT, fontFace: F, margin: 0 });
  s.addText("colorTo · 颜色强调", { x: 8.3, y: 3.0, w: 3.6, h: 0.35, fontSize: 13,
    color: T.MUTED, fontFace: F, margin: 0 });
  s.addShape(p.shapes.ROUNDED_RECTANGLE, { objectName: "bar-grow", x: 8.3, y: 4.3,
    w: 3.2, h: 0.55, rectRadius: 0.05, fill: { color: T.PRIMARY } });
  s.addText("grow · 生长强调", { x: 8.3, y: 4.95, w: 3.6, h: 0.35, fontSize: 13,
    color: T.MUTED, fontFace: F, margin: 0 });
}

// —— p4 路径动画 ——
{
  const s = p.addSlide();
  page(s, "路径动画", "move · 元素沿路径位移（autoRev 自动往返）", 4,
    "圆球往返、箭头平移；切换动效：cover。");
  s.addShape(p.shapes.OVAL, { objectName: "ball", x: 1.6, y: 2.6, w: 1.1, h: 1.1,
    fill: { color: T.ACCENT } });
  s.addText("move autoRev", { x: 1.1, y: 3.85, w: 2.2, h: 0.35, fontSize: 13,
    color: T.MUTED, fontFace: F, align: "center", margin: 0 });
  s.addText("→", { objectName: "arrow-move", x: 1.6, y: 5.0, w: 1.2, h: 0.8,
    fontSize: 44, bold: true, color: T.PRIMARY, fontFace: F, margin: 0 });
  s.addText("move dx=0.3", { x: 1.1, y: 5.85, w: 2.2, h: 0.35, fontSize: 13,
    color: T.MUTED, fontFace: F, align: "center", margin: 0 });
  s.addText("路径动画适合表达「流向」：\n数据从采集端流向看板、\n物料从仓库流向产线。",
    { x: 5.2, y: 2.5, w: 7.2, h: 2.2, fontSize: 18, color: T.TEXT, fontFace: F,
      margin: 0, lineSpacing: 34 });
}

// —— p5 图表构建 ——
{
  const s = p.addSlide();
  page(s, "图表构建", "柱状图按 category 逐根升起 · 折线图按 series 逐条画出", 5,
    "图表自身支持按类/按系列构建；切换动效：wipe。");
  s.addChart(p.charts.BAR, [{
    name: "季度营收（万元）", labels: ["Q1", "Q2", "Q3", "Q4E"], values: [3150, 3520, 4280, 5100],
  }], {
    objectName: "chart-cat", x: 0.7, y: 1.7, w: 6.0, h: 4.9, barDir: "col",
    chartColors: [T.SERIES[0], T.SERIES[0], T.ACCENT, T.SERIES[2]], varyColors: true,
    chartArea: { fill: { color: T.BG } },
    catAxisLabelColor: T.MUTED, catAxisLabelFontSize: 12, catAxisLabelFontFace: F,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: T.TEXT,
    dataLabelFontSize: 12, dataLabelFontFace: F, showLegend: false, showTitle: false,
  });
  s.addChart(p.charts.LINE, [
    { name: "付费企业", labels: ["M1", "M2", "M3", "M4", "M5", "M6"], values: [8, 14, 22, 35, 52, 80] },
    { name: "流失企业", labels: ["M1", "M2", "M3", "M4", "M5", "M6"], values: [1, 2, 2, 3, 4, 4] },
  ], {
    objectName: "chart-ser", x: 7.0, y: 1.7, w: 5.8, h: 4.9,
    chartColors: [T.SERIES[0], T.ACCENT], lineSize: 2.5, lineSmooth: false,
    lineDataSymbol: "circle", lineDataSymbolSize: 6,
    chartArea: { fill: { color: T.BG } },
    catAxisLabelColor: T.MUTED, catAxisLabelFontSize: 12, catAxisLabelFontFace: F,
    valAxisLabelColor: T.MUTED, valAxisLabelFontSize: 12, valAxisLabelFontFace: F,
    valGridLine: { color: T.GRID, size: 0.5 }, catGridLine: { style: "none" },
    showLegend: true, legendPos: "b", legendColor: T.MUTED, legendFontSize: 12,
    legendFontFace: F, showTitle: false,
  });
}

// —— p6 逐段文字 ——
{
  const s = p.addSlide();
  page(s, "逐段文字", "byPara · 要点逐条进场（点击节奏或自动级联）", 6,
    "本页为 click 触发：每点一下出一条；切换动效：split。");
  const bu = () => ({ code: "2022", indent: 12 });
  s.addText([
    { text: "数据要有来源，或标注 est.——不编造数字", options: { bullet: bu(), breakLine: true } },
    { text: "每页只讲一个关键信息，标题写结论", options: { bullet: bu(), breakLine: true } },
    { text: "动效为叙事服务：强调关键，而不是炫技", options: { bullet: bu(), breakLine: true } },
    { text: "交付前必过 QA：程序化检查 + 视觉验收", options: { bullet: bu() } },
  ], { objectName: "bullets", x: 1.2, y: 2.2, w: 10.5, h: 4.2, fontSize: 22,
    color: T.TEXT, fontFace: F, paraSpaceAfter: 22, margin: 0 });
}

// —— p7 退场编排 ——
{
  const s = p.addSlide();
  page(s, "退场编排", "fadeOut / flyOut · 讲完即走，页面留白给结论", 7,
    "三个词先出现再退场，终句 fade 接上；切换动效：checkerboard。");
  const words = [["chip1", "采集", 1.4], ["chip2", "分析", 5.4], ["chip3", "行动", 9.4]];
  words.forEach(wd => {
    s.addShape(p.shapes.ROUNDED_RECTANGLE, { objectName: wd[0], x: wd[2], y: 2.4,
      w: 2.5, h: 1.15, rectRadius: 0.1, fill: { color: T.CARD },
      line: { color: T.LINE, width: 0.75 } });
    s.addText(wd[1], { x: wd[2], y: 2.4, w: 2.5, h: 1.15, fontSize: 24, bold: true,
      color: T.TEXT, fontFace: F, align: "center", valign: "middle", margin: 0 });
  });
  s.addShape(p.shapes.OVAL, { objectName: "bye-icon", x: 6.35, y: 4.35, w: 0.65, h: 0.65,
    fill: { color: T.ACCENT } });
  s.addText("剩下的，交给结论。", { objectName: "final-msg", x: 0, y: 5.3, w: W, h: 0.9,
    fontSize: 34, bold: true, color: T.PRIMARY, fontFace: F, align: "center", margin: 0 });
}

// —— p8 封底 ——
{
  const s = p.addSlide();
  s.background = { color: T.BG };
  s.addText("Office-Agent", { objectName: "end-title", x: 0, y: 3.0, w: W, h: 1.0,
    fontSize: 40, bold: true, color: T.ON_DARK, fontFace: F, align: "center",
    charSpacing: 6, margin: 0 });
  s.addText("动效由 animate.py 注入 · anim.json 声明", { x: 0, y: 4.15, w: W, h: 0.4,
    fontSize: 14, color: T.DIM, fontFace: F, align: "center", margin: 0 });
  s.addNotes("切换动效：wheel。");
}

p.writeFile({ fileName: "动效能力演示.pptx" }).then(f => console.log("written:", f));
