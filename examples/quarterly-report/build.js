// build.js — Q3 经营汇报 · 汇报档·墨蓝商务
// 运行: node build.js  （先加载官方 pptx 技能规范；坐标全部从 theme.js 的 W/H/M 派生）
const pptxgen = require("pptxgenjs");
const T = require("./theme.js");

const p = new pptxgen();
p.layout = "LAYOUT_WIDE";
p.author = "Office-Agent";
p.title = "Q3 经营汇报 — 智能喂食器业务";

const F = T.FONT, W = T.W, H = T.H, M = T.M;
const shadow = () => ({ type: "outer", color: "000000", blur: 6, offset: 2, angle: 45, opacity: 0.12 });

// —— 通用骨架 ——
function contentSlide(title, sub, note) {
  const s = p.addSlide();
  s.background = { color: T.BG };
  s.addText(title, { x: M, y: 0.42, w: W - 2 * M, h: 0.6, fontSize: 30, bold: true,
    color: T.PRIMARY, fontFace: F, margin: 0 });
  if (sub) s.addText(sub, { x: M, y: 1.04, w: W - 2 * M, h: 0.34, fontSize: 14,
    color: T.MUTED, fontFace: F, margin: 0 });
  if (note) s.addNotes(note);
  return s;
}
function footer(s, num, source) {
  s.addText(source || "示例数据，仅供演示", { x: M, y: 7.08, w: 6, h: 0.3,
    fontSize: 12, color: T.MUTED, fontFace: F, margin: 0 });
  s.addText(`${num} / 11`, { x: W - M - 1.2, y: 7.08, w: 1.2, h: 0.3, fontSize: 12,
    color: T.MUTED, fontFace: F, align: "right", margin: 0 });
}
function bigNum(s, txt, x, y, w, size, color) {
  s.addText(txt, { x, y, w, h: (size * 1.45) / 72, fontSize: size, bold: true,
    color: color || T.ACCENT, fontFace: F, margin: 0 });
}
function hairline(s, x, y, w) {
  s.addShape(p.shapes.LINE, { x, y, w, h: 0, line: { color: "D7DEE8", width: 0.75 } });
}

// —— p1 封面 ——
{
  const s = p.addSlide();
  s.background = { color: T.DARK };
  s.addShape(p.shapes.OVAL, { x: 9.1, y: -2.2, w: 6.4, h: 6.4,
    fill: { color: T.DARK, transparency: 100 }, line: { color: "2E4A6B", width: 1 } });
  s.addShape(p.shapes.OVAL, { x: 10.4, y: 4.9, w: 3.6, h: 3.6,
    fill: { color: T.DARK, transparency: 100 }, line: { color: "24405E", width: 1 } });
  s.addText("FEEDMEOW · 智能喂食器业务", { x: M, y: 0.55, w: 6, h: 0.32, fontSize: 13,
    color: "8FA6C0", fontFace: F, charSpacing: 3, margin: 0 });
  s.addText("Q3 经营汇报", { x: M, y: 2.55, w: 9.5, h: 1.15, fontSize: 54, bold: true,
    color: "FFFFFF", fontFace: F, margin: 0 });
  s.addText("营收增长 · 成本归因 · Q4 行动", { x: M, y: 3.75, w: 9, h: 0.45,
    fontSize: 20, color: "A8C0DC", fontFace: F, margin: 0 });
  s.addText([
    { text: "汇报人：经营分析组", options: { breakLine: true } },
    { text: "2026 年 10 月 · 数据为示例，仅供演示" },
  ], { x: M, y: 6.45, w: 7, h: 0.7, fontSize: 13, color: "8FA6C0", fontFace: F,
    paraSpaceAfter: 4, margin: 0 });
  s.addNotes("开场 30 秒：先说结论——增长达标、成本超支、Q4 两件事。");
}

// —— p2 本期结论 ——
{
  const s = contentSlide("本期结论", "先说结论，细节见后文对应章节",
    "三条结论对应三个章节：p4 营收、p5 成本、p8 行动。");
  const rows = [
    ["01", "营收 4,280 万，同比增长 23%，超季度目标 5pct", "F2 新品上量与渠道扩张双驱动", T.PRIMARY],
    ["02", "单台履约成本上升 9%，增量 78% 来自物流与包材", "外部涨价 + 包材损耗率偏高", T.ACCENT],
    ["03", "Q4 主抓两件事：降本 1,500 万、扩张两条渠道", "详见 p8 改进措施与 p9 目标", T.PRIMARY],
  ];
  rows.forEach((r, i) => {
    const y = 1.75 + i * 1.72;
    s.addText(r[0], { x: M, y: y + 0.05, w: 1.3, h: 1.0, fontSize: 44, bold: true,
      color: r[3], fontFace: F, margin: 0 });
    s.addText(r[1], { x: 2.0, y: y, w: 10.3, h: 0.55, fontSize: 21, bold: true,
      color: T.TEXT, fontFace: F, margin: 0 });
    s.addText(r[2], { x: 2.0, y: y + 0.62, w: 10.3, h: 0.4, fontSize: 14,
      color: T.MUTED, fontFace: F, margin: 0 });
    if (i < 2) hairline(s, M, y + 1.42, W - 2 * M);
  });
  footer(s, 2);
}

// —— p3 关键进展 ——
{
  const s = contentSlide("关键进展", "Q3 四件大事",
    "时间线从左到右讲，每件事一句话，重点停在 F2 上量。");
  const y0 = 3.1;
  s.addShape(p.shapes.LINE, { x: 1.1, y: y0, w: 11.1, h: 0, line: { color: "C9D4E2", width: 1.5 } });
  const nodes = [
    ["7 月", "F2 新品上量", "首月出货 3.2 万台", 1.1],
    ["8 月", "进驻头部渠道", "京东自营 + 天猫旗舰店", 4.1],
    ["9 月", "复购率 38%", "环比 +6pct，订阅制贡献过半", 7.1],
    ["9 月末", "海外试点启动", "东南亚两国小批量试销", 10.1],
  ];
  nodes.forEach((n, i) => {
    const x = n[3];
    s.addShape(p.shapes.OVAL, { x: x + 0.06, y: y0 - 0.11, w: 0.22, h: 0.22,
      fill: { color: i === 3 ? T.ACCENT : T.PRIMARY } });
    s.addText(n[0], { x, y: y0 - 0.75, w: 2.1, h: 0.35, fontSize: 14, bold: true,
      color: T.PRIMARY, fontFace: F, margin: 0 });
    s.addText(n[1], { x, y: y0 + 0.35, w: 2.5, h: 0.4, fontSize: 17, bold: true,
      color: T.TEXT, fontFace: F, margin: 0 });
    s.addText(n[2], { x, y: y0 + 0.8, w: 2.5, h: 0.7, fontSize: 13,
      color: T.MUTED, fontFace: F, margin: 0 });
  });
  footer(s, 3);
}

// —— p4 营收表现 ——
{
  const s = contentSlide("营收表现", "增长来自量价齐升，F2 贡献 42% 出货",
    "先报大数字，再读图：Q3 4,280 万超目标；Q4E 为目标值非预测。");
  bigNum(s, "23%", M, 1.7, 3.4, 64, T.ACCENT);
  s.addText("同比增长 · 超季度目标 5pct", { x: M, y: 3.1, w: 3.6, h: 0.4, fontSize: 15,
    color: T.MUTED, fontFace: F, margin: 0 });
  s.addText([
    { text: "Q3 营收 4,280 万", options: { breakLine: true, bold: true, fontSize: 18, color: T.TEXT } },
    { text: "环比 +21.6%", options: { fontSize: 14, color: T.MUTED } },
  ], { x: M, y: 4.6, w: 3.6, h: 1.0, fontFace: F, paraSpaceAfter: 6, margin: 0 });
  s.addChart(p.charts.BAR, [{
    name: "季度营收（万元）",
    labels: ["Q1", "Q2", "Q3", "Q4 目标"],
    values: [3150, 3520, 4280, 5100],
  }], {
    x: 4.6, y: 1.7, w: 8.1, h: 5.0, barDir: "col", barGapWidthPct: 60,
    chartColors: [T.SERIES[0], T.SERIES[0], T.ACCENT, T.SERIES[2]], varyColors: true,
    chartArea: { fill: { color: T.BG } },
    catAxisLabelColor: T.MUTED, catAxisLabelFontSize: 13, catAxisLabelFontFace: F,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    valAxisLineShow: false, serAxisLineShow: false,
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: T.TEXT,
    dataLabelFontSize: 13, dataLabelFontFace: F,
    showLegend: false, showTitle: false,
  });
  footer(s, 4);
}

// —— p5 成本问题 ——
{
  const s = contentSlide("成本问题", "单台履约成本 213 元，同比 +9%",
    "环形图讲结构：物流+包材合计 56%；右侧回答增量从哪来。");
  s.addChart(p.charts.DOUGHNUT, [{
    name: "单台成本结构",
    labels: ["物流 34%", "包材 22%", "制造 31%", "其他 13%"],
    values: [34, 22, 31, 13],
  }], {
    x: 0.8, y: 1.65, w: 5.6, h: 4.9, holeSize: 62,
    chartColors: [T.ACCENT, "C2540A", T.SERIES[0], T.SERIES[2]],
    chartArea: { fill: { color: T.BG } },
    showPercent: false, showValue: false, showTitle: false,
    showLegend: true, legendPos: "b", legendColor: T.MUTED,
    legendFontSize: 13, legendFontFace: F,
    dataBorder: { pt: 1.5, color: T.BG },
  });
  s.addText("物流与包材", { x: 2.35, y: 3.55, w: 2.5, h: 0.4, fontSize: 15, bold: true,
    color: T.TEXT, fontFace: F, align: "center", margin: 0 });
  s.addText("占成本 56%", { x: 2.35, y: 3.95, w: 2.5, h: 0.35, fontSize: 13,
    color: T.MUTED, fontFace: F, align: "center", margin: 0 });
  bigNum(s, "78%", 7.1, 1.8, 3.2, 56, T.ACCENT);
  s.addText("成本增量中物流与包材的占比", { x: 7.1, y: 3.05, w: 5.4, h: 0.4,
    fontSize: 15, color: T.MUTED, fontFace: F, margin: 0 });
  const causes = [
    ["物流费上涨", "快运单价 Q3 环比 +12%，大件抛比吃亏"],
    ["包材损耗率高", "F2 包装兼容旧模具，破损率 3.1%"],
  ];
  causes.forEach((c, i) => {
    const y = 3.75 + i * 1.25;
    s.addText(c[0], { x: 7.1, y, w: 5.4, h: 0.42, fontSize: 17, bold: true,
      color: T.TEXT, fontFace: F, margin: 0 });
    s.addText(c[1], { x: 7.1, y: y + 0.45, w: 5.4, h: 0.4, fontSize: 13.5,
      color: T.MUTED, fontFace: F, margin: 0 });
  });
  footer(s, 5);
}

// —— p6 产能与交付 ——
{
  const s = contentSlide("产能与交付", "良率爬坡符合预期，交付周期回到常态",
    "两个数字先说：97.2% 和 5.2 天；折线说明 9 月已稳定。");
  bigNum(s, "97.2%", M, 1.75, 4.2, 54, T.PRIMARY);
  s.addText("F2 良率（9 月末）", { x: M, y: 2.98, w: 4.2, h: 0.38, fontSize: 15,
    color: T.MUTED, fontFace: F, margin: 0 });
  bigNum(s, "4.8 天", M, 3.9, 4.2, 54, T.PRIMARY);
  s.addText("平均交付周期（较 Q2 -1.2 天）", { x: M, y: 5.13, w: 4.2, h: 0.38, fontSize: 15,
    color: T.MUTED, fontFace: F, margin: 0 });
  s.addChart(p.charts.LINE, [{
    name: "周度良率（%）",
    labels: ["W1", "W2", "W3", "W4", "W5", "W6"],
    values: [93.8, 94.6, 95.5, 96.3, 96.9, 97.2],
  }], {
    x: 5.2, y: 1.7, w: 7.5, h: 5.0,
    chartColors: [T.PRIMARY], lineSize: 2.5, lineSmooth: false,
    lineDataSymbol: "circle", lineDataSymbolSize: 6,
    chartArea: { fill: { color: T.BG } },
    catAxisLabelColor: T.MUTED, catAxisLabelFontSize: 12, catAxisLabelFontFace: F,
    valAxisLabelColor: T.MUTED, valAxisLabelFontSize: 12, valAxisLabelFontFace: F,
    valAxisMinVal: 92, valAxisMaxVal: 98,
    valGridLine: { color: "E2E8F0", size: 0.5 }, catGridLine: { style: "none" },
    showValue: false, showLegend: false, showTitle: false,
  });
  s.addText("良率爬坡 · 周度", { x: 5.2, y: 6.62, w: 7.5, h: 0.3, fontSize: 12,
    color: T.MUTED, fontFace: F, align: "center", margin: 0 });
  footer(s, 6);
}

// —— p7 问题与归因 ——
{
  const s = contentSlide("问题与归因", "三大问题，都有明确的动作指向",
    "逐行读：问题→影响→归因；p8 的四项措施与前三行一一对应。");
  const head = ["问题", "影响", "归因"];
  const xs = [M, 4.6, 9.0], ws = [3.9, 4.2, 3.8];
  head.forEach((h, i) => s.addText(h, { x: xs[i], y: 1.6, w: ws[i], h: 0.36,
    fontSize: 14, bold: true, color: T.PRIMARY, fontFace: F, margin: 0 }));
  const rows = [
    ["快运物流费上涨", "履约成本 +34 元/台", "大件网络议价能力弱，单一供应商"],
    ["包材损耗率 3.1%", "报废 + 售后双损失", "F2 沿用旧模具，缓冲结构不匹配"],
    ["渠道返利挤压毛利", "头部渠道毛利率 -4pct", "返利阶梯设置滞后于规模增长"],
  ];
  rows.forEach((r, i) => {
    const y = 2.15 + i * 1.55;
    s.addText(r[0], { x: xs[0], y, w: ws[0], h: 0.5, fontSize: 17, bold: true,
      color: T.ACCENT, fontFace: F, margin: 0 });
    s.addText(r[1], { x: xs[1], y, w: ws[1], h: 0.5, fontSize: 15,
      color: T.TEXT, fontFace: F, margin: 0 });
    s.addText(r[2], { x: xs[2], y, w: ws[2], h: 0.9, fontSize: 14,
      color: T.MUTED, fontFace: F, margin: 0 });
    if (i < 2) hairline(s, M, y + 1.15, W - 2 * M);
  });
  footer(s, 7);
}

// —— p8 改进措施 ——
{
  const s = contentSlide("改进措施", "四项降本行动，目标合计 -1,500 万/年",
    "四张卡片各 20 秒：动作+目标；负责人当场确认，不另开会。");
  const cards = [
    ["物流招标重谈", "引入两家快运比价，目标 -18%", "供应链部 · 11 月中"],
    ["包材 V2 减重", "新模具 + 减重 15%，损耗 ≤1.5%", "包装组 · 12 月上"],
    ["自建华东仓", "区域仓发替代跨省直发，运费 -22%", "履约组 · Q4 末"],
    ["返利政策调整", "阶梯与规模挂钩，毛利 +2pct", "销售运营 · 11 月下"],
  ];
  const cw = 5.9, ch = 2.25, gx = 0.55, gy = 0.45, x0 = M, y0 = 1.75;
  cards.forEach((c, i) => {
    const x = x0 + (i % 2) * (cw + gx), y = y0 + Math.floor(i / 2) * (ch + gy);
    s.addShape(p.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, rectRadius: 0.08,
      fill: { color: T.CARD }, line: { color: "E3E9F1", width: 0.75 }, shadow: shadow() });
    s.addText(c[0], { x: x + 0.35, y: y + 0.28, w: cw - 0.7, h: 0.45, fontSize: 19,
      bold: true, color: T.PRIMARY, fontFace: F, margin: 0 });
    s.addText(c[1], { x: x + 0.35, y: y + 0.85, w: cw - 0.7, h: 0.75, fontSize: 14.5,
      color: T.TEXT, fontFace: F, margin: 0 });
    s.addText(c[2], { x: x + 0.35, y: y + 1.68, w: cw - 0.7, h: 0.35, fontSize: 13,
      color: T.ACCENT, bold: true, fontFace: F, margin: 0 });
  });
  footer(s, 8);
}

// —— p9 Q4 计划与目标 ——
{
  const s = contentSlide("Q4 计划与目标", "一个营收目标，两条增长线",
    "目标 5,100 万的构成：存量自然增长 + 两渠道 + 双十一；里程碑看下半页时间线。");
  const goals = [["5,100 万", "Q4 营收目标", T.ACCENT], ["2 条", "新增头部渠道", T.PRIMARY],
    ["-1,500 万", "年化降本目标", T.PRIMARY]];
  goals.forEach((g, i) => {
    const x = M + i * 4.28;
    bigNum(s, g[0], x, 1.7, 4.0, 44, g[2]);
    s.addText(g[1], { x, y: 2.7, w: 4.0, h: 0.38, fontSize: 15, color: T.MUTED,
      fontFace: F, margin: 0 });
  });
  const y0 = 4.35;
  s.addShape(p.shapes.LINE, { x: 1.3, y: y0, w: 10.7, h: 0, line: { color: "C9D4E2", width: 1.5 } });
  const miles = [["11 月", "物流新约生效"], ["11 月中", "渠道 A 上线"], ["12 月", "华东仓投产; 双十一复盘"], ["1 月", "年度目标锁定"]];
  miles.forEach((m2, i) => {
    const x = 1.3 + i * 2.68;
    s.addShape(p.shapes.OVAL, { x: x - 0.05, y: y0 - 0.1, w: 0.2, h: 0.2,
      fill: { color: i === 2 ? T.ACCENT : T.PRIMARY } });
    s.addText(m2[0], { x: x - 0.3, y: y0 + 0.3, w: 2.4, h: 0.35, fontSize: 14, bold: true,
      color: T.PRIMARY, fontFace: F, margin: 0 });
    s.addText(m2[1], { x: x - 0.3, y: y0 + 0.68, w: 2.45, h: 0.6, fontSize: 13,
      color: T.MUTED, fontFace: F, margin: 0 });
  });
  footer(s, 9);
}

// —— p10 双十一备战 ——
{
  const s = contentSlide("双十一备战", "11.11 大促倒计时 3 周",
    "备货是重点：8 万台对应日均 6,000+ 单峰值；履约和客服排班同步跟上。");
  bigNum(s, "8 万台", M, 1.9, 4.6, 54, T.ACCENT);
  s.addText("F2 备货目标（含 15% 安全库存）", { x: M, y: 3.1, w: 4.6, h: 0.4, fontSize: 15,
    color: T.MUTED, fontFace: F, margin: 0 });
  const items = [
    ["备货", "产线 11 月初拉满，11.1 前入库 6.5 万台"],
    ["渠道", "两渠道大促资源位与坑位费已确认，主推 F2 套装"],
    ["履约", "华东仓 11.1 投产赶上大促，临时人力 +40%"],
  ];
  items.forEach((it, i) => {
    const y = 1.9 + i * 1.45;
    s.addText(it[0], { x: 5.6, y, w: 1.4, h: 0.4, fontSize: 17, bold: true,
      color: T.PRIMARY, fontFace: F, margin: 0 });
    s.addText(it[1], { x: 7.1, y, w: 5.4, h: 0.85, fontSize: 14.5,
      color: T.TEXT, fontFace: F, margin: 0 });
    if (i < 2) hairline(s, 5.6, y + 1.05, W - M - 5.6);
  });
  footer(s, 10);
}

// —— p10 封底 ——
{
  const s = p.addSlide();
  s.background = { color: T.DARK };
  s.addShape(p.shapes.OVAL, { x: -2.6, y: 3.8, w: 6.2, h: 6.2,
    fill: { color: T.DARK, transparency: 100 }, line: { color: "24405E", width: 1 } });
  s.addText("谢谢", { x: 0, y: 2.7, w: W, h: 1.1, fontSize: 48, bold: true,
    color: "FFFFFF", fontFace: F, align: "center", margin: 0 });
  s.addText("明细数据见经营看板 · 示例数据仅供演示", { x: 0, y: 3.95, w: W, h: 0.4,
    fontSize: 15, color: "8FA6C0", fontFace: F, align: "center", margin: 0 });
  s.addNotes("收尾：重复一遍三条结论，留 10 分钟问答。");
}

p.writeFile({ fileName: "Q3经营汇报-智能喂食器业务.pptx" }).then(f => console.log("written:", f));
