// build.js — 灵犀会议机 · 天使轮路演 · 路演档·深空蓝金（12 页，全深色）
const pptxgen = require("pptxgenjs");
const T = require("./theme.js");

const p = new pptxgen();
p.layout = "LAYOUT_WIDE";
p.author = "Office-Agent";
p.title = "灵犀会议机 · 天使轮商业路演";

const F = T.FONT, W = T.W, H = T.H, M = T.M;

function contentSlide(title, sub, note) {
  const s = p.addSlide();
  s.background = { color: T.BG };
  s.addText(title, { x: M, y: 0.42, w: W - 2 * M, h: 0.6, fontSize: 30, bold: true,
    color: T.TEXT, fontFace: F, margin: 0 });
  if (sub) s.addText(sub, { x: M, y: 1.04, w: W - 2 * M, h: 0.34, fontSize: 14,
    color: T.MUTED, fontFace: F, margin: 0 });
  if (note) s.addNotes(note);
  return s;
}
function footer(s, num, source) {
  s.addText(source || "示例数据，仅供演示", { x: M, y: 7.08, w: 6, h: 0.3,
    fontSize: 12, color: T.MUTED, fontFace: F, margin: 0 });
  s.addText(`${num} / 12`, { x: W - M - 1.2, y: 7.08, w: 1.2, h: 0.3, fontSize: 12,
    color: T.MUTED, fontFace: F, align: "right", margin: 0 });
}
function bigNum(s, txt, x, y, w, size, color) {
  s.addText(txt, { x, y, w, h: (size * 1.45) / 72, fontSize: size, bold: true,
    color: color || T.ACCENT, fontFace: F, margin: 0 });
}
function hairline(s, x, y, w) {
  s.addShape(p.shapes.LINE, { x, y, w, h: 0, line: { color: T.LINE, width: 0.75 } });
}

// —— p1 封面 ——
{
  const s = p.addSlide();
  s.background = { color: T.BG };
  s.addShape(p.shapes.OVAL, { x: 8.9, y: -2.4, w: 6.8, h: 6.8,
    fill: { color: T.BG, transparency: 100 }, line: { color: T.LINE, width: 1 } });
  s.addShape(p.shapes.OVAL, { x: 10.6, y: 4.7, w: 3.9, h: 3.9,
    fill: { color: T.BG, transparency: 100 }, line: { color: "1A3457", width: 1 } });
  s.addText("LINGXI · 天使轮融资路演", { x: M, y: 0.55, w: 7, h: 0.32, fontSize: 13,
    color: T.MUTED, fontFace: F, charSpacing: 3, margin: 0 });
  s.addText("灵犀会议机", { x: M, y: 2.45, w: 10, h: 1.1, fontSize: 54, bold: true,
    color: T.TEXT, fontFace: F, margin: 0 });
  s.addText("把 90 分钟会议，变成 3 页行动清单", { x: M, y: 3.7, w: 10, h: 0.5,
    fontSize: 22, color: T.ACCENT, fontFace: F, margin: 0 });
  s.addText("2026 年 10 月 · 数据为示例，仅供演示", { x: M, y: 6.5, w: 7, h: 0.35,
    fontSize: 13, color: T.MUTED, fontFace: F, margin: 0 });
  s.addNotes("开场 15 秒说定位句，然后直接进痛点。");
}

// —— p2 痛点 ——
{
  const s = contentSlide("会议结束，执行才开始", "但执行需要的信息，恰恰在最需要的时候丢了",
    "三个数字从左到右递进：时间投入、记录缺失、执行漏洞。");
  const stats = [
    ["18.6 小时", "白领平均每周会议时长", T.TEXT],
    ["30%", "纪要覆盖率（有记录的会议占比）", T.TEXT],
    ["42%", "行动项遗漏率", T.ACCENT],
  ];
  stats.forEach((it, i) => {
    const x = M + i * 4.28;
    bigNum(s, it[0], x, 2.1, 4.0, 46, it[2]);
    s.addText(it[1], { x, y: 3.35, w: 3.9, h: 0.7, fontSize: 14, color: T.MUTED,
      fontFace: F, margin: 0 });
    if (i < 2) hairline(s, x + 4.05, 2.3, 0.005);
  });
  s.addText("会议的价值在会后 24 小时内流失 60% —— 记录工具解决不了跟进问题。",
    { x: M, y: 5.15, w: W - 2 * M, h: 0.5, fontSize: 18, bold: true, color: T.ACCENT,
      fontFace: F, margin: 0 });
  footer(s, 2, "来源：团队 2026 年 8 月对 60 家企业调研（示例数据）");
}

// —— p3 解决方案 ——
{
  const s = contentSlide("灵犀会议机", "拾音硬件 + 执行闭环 SaaS",
    "流程图从左到右：录→写→派→追；强调闭环二字，竞争页会回收。");
  const steps = [
    ["拾音转写", "6 麦阵列 + 本地转写芯片，准确率优先", T.PRIMARY],
    ["智能总结", "3 页纪要模板：结论/讨论/风险", T.PRIMARY],
    ["行动项分发", "自动识别责任人，一键派发到企业微信/钉钉", T.ACCENT],
    ["回执追踪", "未完成自动升级提醒，闭环率看板", T.ACCENT],
  ];
  steps.forEach((st, i) => {
    const x = M + i * 3.22;
    s.addShape(p.shapes.ROUNDED_RECTANGLE, { x, y: 2.3, w: 2.85, h: 2.5, rectRadius: 0.08,
      fill: { color: T.CARD }, line: { color: T.LINE, width: 0.75 } });
    s.addText(`0${i + 1}`, { x: x + 0.25, y: 2.55, w: 2.3, h: 0.5, fontSize: 22, bold: true,
      color: st[2], fontFace: F, margin: 0 });
    s.addText(st[0], { x: x + 0.25, y: 3.15, w: 2.4, h: 0.45, fontSize: 17, bold: true,
      color: T.TEXT, fontFace: F, margin: 0 });
    s.addText(st[1], { x: x + 0.25, y: 3.65, w: 2.4, h: 1.0, fontSize: 12.5,
      color: T.MUTED, fontFace: F, margin: 0 });
    if (i < 3) s.addText("→", { x: x + 2.82, y: 3.25, w: 0.45, h: 0.5, fontSize: 20,
      color: T.MUTED, fontFace: F, align: "center", margin: 0 });
  });
  s.addText("硬件保证拾音质量，SaaS 完成从「记录」到「执行闭环」的最后一公里。",
    { x: M, y: 5.35, w: W - 2 * M, h: 0.45, fontSize: 16, color: T.TEXT, fontFace: F, margin: 0 });
  footer(s, 3);
}

// —— p4 市场规模（迭代 R1：数字块 → 同心圆图示） ——
{
  const s = contentSlide("市场规模", "从智能会议终端切入的效率软件市场",
    "同心圆讲分层：外圈 TAM、中圈 SAM、核心 SOM；左侧图例逐层解释。");
  const cx = 9.0, cy = 4.15;
  s.addShape(p.shapes.OVAL, { x: cx - 2.5, y: cy - 2.5, w: 5.0, h: 5.0,
    fill: { color: T.CARD }, line: { color: T.LINE, width: 1 } });
  s.addShape(p.shapes.OVAL, { x: cx - 1.55, y: cy - 1.55, w: 3.1, h: 3.1,
    fill: { color: T.PRIMARY, transparency: 45 }, line: { color: T.PRIMARY, width: 1 } });
  s.addShape(p.shapes.OVAL, { x: cx - 0.75, y: cy - 0.75, w: 1.5, h: 1.5,
    fill: { color: T.ACCENT } });
  s.addText("TAM · 400 亿", { x: cx - 1.6, y: cy - 2.32, w: 3.2, h: 0.35, fontSize: 14,
    bold: true, color: T.MUTED, fontFace: F, align: "center", margin: 0 });
  s.addText("SAM · 60 亿", { x: cx - 1.6, y: cy - 1.38, w: 3.2, h: 0.35, fontSize: 14,
    bold: true, color: T.TEXT, fontFace: F, align: "center", margin: 0 });
  s.addText("SOM\n3 亿", { x: cx - 0.75, y: cy - 0.55, w: 1.5, h: 1.1, fontSize: 15,
    bold: true, color: "0B1B33", fontFace: F, align: "center", valign: "middle", margin: 0 });
  const legend = [
    ["400 亿", "TAM · 企业会议设备与效率软件", T.MUTED],
    ["60 亿", "SAM · 智能会议终端与配套 SaaS", T.PRIMARY],
    ["3 亿", "SOM · 三年可获取市场", T.ACCENT],
  ];
  legend.forEach((l, i) => {
    const y = 2.2 + i * 1.5;
    s.addShape(p.shapes.OVAL, { x: M, y: y + 0.18, w: 0.22, h: 0.22, fill: { color: l[2] } });
    bigNum(s, l[0], M + 0.45, y, 2.6, 34, l[2]);
    s.addText(l[1], { x: 3.7, y: y + 0.32, w: 3.4, h: 0.8, fontSize: 14,
      color: T.TEXT, fontFace: F, margin: 0 });
  });
  footer(s, 4, "估算（est.），基于公开行业报告口径 · 示例数据");
}

// —— p5 商业模式 ——
{
  const s = contentSlide("商业模式", "硬件获客，SaaS 赚钱",
    "一句话逻辑：硬件平进平出让客户零门槛，利润与留存都在订阅。");
  const cards = [
    ["2,999 元", "会议机（一次性）", "按成本价销售，获客入口", T.PRIMARY],
    ["1.2 万/年", "灵犀 SaaS 订阅", "按企业授权，利润与留存载体", T.ACCENT],
  ];
  cards.forEach((c, i) => {
    const x = M + i * 6.42;
    s.addShape(p.shapes.ROUNDED_RECTANGLE, { x, y: 2.0, w: 5.9, h: 2.6, rectRadius: 0.08,
      fill: { color: T.CARD }, line: { color: T.LINE, width: 0.75 } });
    bigNum(s, c[0], x + 0.4, 2.35, 5.0, 40, c[3]);
    s.addText(c[1], { x: x + 0.4, y: 3.42, w: 5.0, h: 0.4, fontSize: 16, bold: true,
      color: T.TEXT, fontFace: F, margin: 0 });
    s.addText(c[2], { x: x + 0.4, y: 3.85, w: 5.0, h: 0.4, fontSize: 13,
      color: T.MUTED, fontFace: F, margin: 0 });
  });
  s.addText("试点客单价 1.2 万/年 · NRR 118% —— 订阅驱动，客户越用越离不开。",
    { x: M, y: 5.15, w: W - 2 * M, h: 0.45, fontSize: 16, color: T.TEXT, fontFace: F, margin: 0 });
  footer(s, 5);
}

// —— p6 早期数据 ——
{
  const s = contentSlide("早期验证", "12 家试点企业，8 家付费",
    "漏斗从上往下读：试点→付费 67%；右侧三个质量指标。");
  s.addText("试点企业", { x: M, y: 2.0, w: 2.2, h: 0.4, fontSize: 15, color: T.MUTED,
    fontFace: F, margin: 0 });
  s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: M, y: 2.45, w: 11.8, h: 0.85, rectRadius: 0.06,
    fill: { color: T.CARD }, line: { color: T.LINE, width: 0.75 } });
  s.addText("12 家", { x: 0.8, y: 2.45, w: 11.3, h: 0.85, fontSize: 20, bold: true,
    color: T.TEXT, fontFace: F, valign: "middle", margin: 0 });
  s.addText("付费转化", { x: M, y: 3.55, w: 2.2, h: 0.4, fontSize: 15, color: T.MUTED,
    fontFace: F, margin: 0 });
  s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: M, y: 4.0, w: 7.9, h: 0.85, rectRadius: 0.06,
    fill: { color: T.PRIMARY, transparency: 55 }, line: { color: T.PRIMARY, width: 1 } });
  s.addText("8 家 · 67%", { x: 0.8, y: 4.0, w: 7.4, h: 0.85, fontSize: 20, bold: true,
    color: T.TEXT, fontFace: F, valign: "middle", margin: 0 });
  const qs = [["118%", "NRR 净收入留存"], ["76%", "周活跃使用率"], ["1.2 万", "平均客单价（万元/年）"]];
  qs.forEach((q, i) => {
    const y = 5.3;
    const x = M + i * 4.28;
    bigNum(s, q[0], x, y, 4.0, 30, T.ACCENT);
    s.addText(q[1], { x, y: y + 0.62, w: 4.0, h: 0.35, fontSize: 13, color: T.MUTED,
      fontFace: F, margin: 0 });
  });
  footer(s, 6, "试点期数据（示例）");
}

// —— p7 竞争格局 ——
{
  const s = contentSlide("竞争格局", "我们不和转写工具比谁记得全，和没人跟进比谁让事落地",
    "表格三行两列对比；最后一行是我们，用 ACCENT 强调闭环优势。");
  const head = ["维度", "通用转写软件", "智能会议平板", "灵犀会议机"];
  const rows = [
    ["拾音质量", "依赖设备麦克风", "会议室定向阵列", "6 麦阵列 · 本地转写"],
    ["行动项闭环", "无", "无", "分发 · 回执追踪"],
    ["部署成本", "低", "单台 2 万起", "硬件 2,999 · 订阅"],
  ];
  const tableData = [
    head.map((h, i) => ({ text: h, options: { fill: { color: "1E3A5F" }, color: "FFFFFF",
      bold: true, fontSize: 14, fontFace: F, align: i === 0 ? "left" : "left" } })),
    ...rows.map((r, ri) => r.map((c, ci) => ({
      text: c,
      options: {
        fill: { color: ri === 1 && ci === 3 ? T.ACCENT : (ri % 2 ? "0F2747" : T.CARD) },
        color: ci === 3 && ri === 1 ? "0B1B33" : (ci === 3 ? T.ACCENT : (ci === 0 ? T.MUTED : T.TEXT)),
        bold: ci === 3 || ci === 0,
        fontSize: 13, fontFace: F,
      },
    }))),
  ];
  s.addTable(tableData, { x: M, y: 1.9, w: W - 2 * M, colW: [2.2, 3.4, 3.4, 3.83],
    rowH: 0.75, border: { pt: 0.75, color: T.LINE }, valign: "middle" });
  s.addText("差异化的本质：别人停在「记录」，灵犀走到「跟进」。",
    { x: M, y: 5.6, w: W - 2 * M, h: 0.45, fontSize: 16, bold: true, color: T.ACCENT,
      fontFace: F, margin: 0 });
  footer(s, 7);
}

// —— p8 壁垒 ——
{
  const s = contentSlide("为什么是我们", "三道壁垒",
    "三张卡片各 20 秒：声学是硬壁垒，数据飞轮随时间变宽。");
  const cards = [
    ["声学硬件壁垒", "拾音阵列与转写芯片的联合调优，需要 3 年以上声学团队积累，纯软件公司难以复制"],
    ["闭环数据飞轮", "行动项完成数据反哺总结模型，用得越久总结越准——单点工具没有这个数据"],
    ["行业 Know-how", "60 家企业调研沉淀的纪要模板与行动项识别规则，随行业扩展持续加厚"],
  ];
  cards.forEach((c, i) => {
    const x = M + i * 4.28;
    s.addShape(p.shapes.ROUNDED_RECTANGLE, { x, y: 2.0, w: 3.9, h: 3.4, rectRadius: 0.08,
      fill: { color: T.CARD }, line: { color: T.LINE, width: 0.75 } });
    s.addText(`0${i + 1}`, { x: x + 0.3, y: 2.3, w: 3.3, h: 0.55, fontSize: 26, bold: true,
      color: T.ACCENT, fontFace: F, margin: 0 });
    s.addText(c[0], { x: x + 0.3, y: 3.0, w: 3.3, h: 0.45, fontSize: 17, bold: true,
      color: T.TEXT, fontFace: F, margin: 0 });
    s.addText(c[1], { x: x + 0.3, y: 3.55, w: 3.3, h: 1.6, fontSize: 13,
      color: T.MUTED, fontFace: F, margin: 0 });
  });
  footer(s, 8);
}

// —— p9 团队 ——
{
  const s = contentSlide("团队", "声学 × 企业服务 × AI，正好拼齐这件事",
    "一人 20 秒，强调背景与岗位的强关联。");
  const team = [
    ["林", "林一舟 · CEO", "前 SaaS 公司销售 VP，带队 0→8,000 家企业客户"],
    ["沈", "沈知远 · CTO", "前音频大厂声学负责人，主导三代拾音阵列量产"],
    ["顾", "顾言 · AI 负责人", "前大厂算法专家，会议摘要方向 5 年积累"],
  ];
  team.forEach((t, i) => {
    const x = M + i * 4.28;
    s.addShape(p.shapes.OVAL, { x: x + 1.35, y: 1.95, w: 1.2, h: 1.2,
      fill: { color: i === 1 ? T.ACCENT : T.PRIMARY } });
    s.addText(t[0], { x: x + 1.35, y: 1.95, w: 1.2, h: 1.2, fontSize: 30, bold: true,
      color: "FFFFFF", fontFace: F, align: "center", valign: "middle", margin: 0 });
    s.addText(t[1], { x, y: 3.4, w: 3.9, h: 0.45, fontSize: 17, bold: true,
      color: T.TEXT, fontFace: F, align: "center", margin: 0 });
    s.addText(t[2], { x: x + 0.25, y: 3.9, w: 3.4, h: 1.1, fontSize: 13,
      color: T.MUTED, fontFace: F, align: "center", margin: 0 });
  });
  footer(s, 9, "团队履历为示例");
}

// —— p10 里程碑 ——
{
  const s = contentSlide("18 个月里程碑", "从产品打磨到规模化复制",
    "时间线四段，最后一个节点回到收入目标。");
  const y0 = 3.4;
  s.addShape(p.shapes.LINE, { x: 1.2, y: y0, w: 10.9, h: 0, line: { color: T.LINE, width: 1.5 } });
  const miles = [
    ["M0–M6", "产品打磨", "量产与首个 20 家标杆客户", 1.2],
    ["M6–M10", "渠道验证", "两条代理渠道跑通模型", 4.0],
    ["M10–M14", "规模投放", "销售团队扩至 15 人", 6.8],
    ["M14–M18", "200 家付费", "营收 2,400 万", 9.6],
  ];
  miles.forEach((m2, i) => {
    const x = m2[3];
    s.addShape(p.shapes.OVAL, { x: x - 0.09, y: y0 - 0.1, w: 0.22, h: 0.22,
      fill: { color: i === 3 ? T.ACCENT : T.PRIMARY } });
    s.addText(m2[0], { x: x - 0.4, y: y0 - 0.75, w: 2.4, h: 0.35, fontSize: 14, bold: true,
      color: T.MUTED, fontFace: F, margin: 0 });
    s.addText(m2[1], { x: x - 0.4, y: y0 + 0.35, w: 2.6, h: 0.45, fontSize: 18, bold: true,
      color: T.TEXT, fontFace: F, margin: 0 });
    s.addText(m2[2], { x: x - 0.4, y: y0 + 0.85, w: 2.6, h: 0.6, fontSize: 13,
      color: T.MUTED, fontFace: F, margin: 0 });
  });
  footer(s, 10);
}

// —— p11 融资 Ask ——
{
  const s = contentSlide("融资计划", "天使轮 · 800 万元",
    "大数字开口：800 万、出让 10%；环形图讲用途，收在里程碑。");
  bigNum(s, "800 万", M, 2.0, 4.6, 56, T.ACCENT);
  s.addText("出让 10%（示例条款）", { x: M, y: 3.2, w: 4.6, h: 0.4, fontSize: 16,
    color: T.MUTED, fontFace: F, margin: 0 });
  s.addText("18 个月里程碑：付费企业 200 家，营收 2,400 万。",
    { x: M, y: 4.3, w: 5.4, h: 0.8, fontSize: 15, color: T.TEXT, fontFace: F, margin: 0 });
  s.addChart(p.charts.DOUGHNUT, [{
    name: "资金用途",
    labels: ["研发 50%", "市场 30%", "运营 20%"],
    values: [50, 30, 20],
  }], {
    x: 6.4, y: 1.6, w: 6.2, h: 5.1, holeSize: 60,
    chartColors: [T.PRIMARY, T.ACCENT, "64748B"],
    chartArea: { fill: { color: T.BG } },
    showPercent: false, showValue: false, showTitle: false,
    showLegend: true, legendPos: "b", legendColor: T.MUTED,
    legendFontSize: 13, legendFontFace: F,
    dataBorder: { pt: 1.5, color: T.BG },
  });
  footer(s, 11);
}

// —— p12 封底 ——
{
  const s = p.addSlide();
  s.background = { color: T.BG };
  s.addShape(p.shapes.OVAL, { x: -2.8, y: 3.6, w: 6.6, h: 6.6,
    fill: { color: T.BG, transparency: 100 }, line: { color: T.LINE, width: 1 } });
  s.addText("让每个会议，都有人跟进。", { x: 0, y: 2.75, w: W, h: 1.0, fontSize: 40,
    bold: true, color: T.TEXT, fontFace: F, align: "center", margin: 0 });
  s.addText("灵犀会议机 · contact@lingxi.example.com · 示例数据仅供演示",
    { x: 0, y: 4.0, w: W, h: 0.4, fontSize: 15, color: T.MUTED, fontFace: F,
      align: "center", margin: 0 });
  s.addNotes("收在定位句，留问答。");
}

p.writeFile({ fileName: "灵犀会议机-天使轮路演.pptx" }).then(f => console.log("written:", f));
