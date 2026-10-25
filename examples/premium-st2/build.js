// build.js — 星尘 ST2 发布会 · 高端档·深空香幕（10 页，大留白+单焦点）
const pptxgen = require("pptxgenjs");
const T = require("./theme.js");

const p = new pptxgen();
p.layout = "LAYOUT_WIDE";
p.author = "Office-Agent";
p.title = "星尘 ST2 · 新品发布会";

const F = T.FONT, W = T.W, H = T.H, M = T.M;

// 高端档页脚极简：仅页码，来源信息内嵌到内容行
function pageTag(s, num) {
  s.addText(`${num}`, { x: W - M - 0.8, y: H - 0.62, w: 0.8, h: 0.3, fontSize: 12,
    color: T.MUTED, fontFace: F, align: "right", margin: 0 });
}
function moon(s, cx, cy, d, color, transp) {
  s.addShape(p.shapes.OVAL, { x: cx - d / 2, y: cy - d / 2, w: d, h: d,
    fill: { color: T.BG, transparency: 100 },
    line: { color, width: 1.25, transparency: transp ?? 0 } });
}

// —— p1 封面 ——
{
  const s = p.addSlide();
  s.background = { color: T.BG };
  moon(s, W - 3.2, 2.1, 2.2, T.PRIMARY, 40);
  s.addText("星尘 ST2", { x: 0, y: 2.6, w: W, h: 1.5, fontSize: 88, bold: true,
    color: T.PRIMARY, fontFace: F, align: "center", charSpacing: 8, margin: 0 });
  s.addText("智 能 香 氛 机", { x: 0, y: 4.25, w: W, h: 0.6, fontSize: 24,
    color: T.TEXT, fontFace: F, align: "center", charSpacing: 10, margin: 0 });
  s.addText("新品发布会 · 2026 年 10 月", { x: 0, y: 6.55, w: W, h: 0.35, fontSize: 13,
    color: T.MUTED, fontFace: F, align: "center", margin: 0 });
  s.addNotes("开场不说话，先放 30 秒品牌视频「夜晚的房间」。");
}

// —— p2 情绪页 · 开场 ——
{
  const s = p.addSlide();
  s.background = { color: T.BG };
  moon(s, W / 2, 2.6, 1.4, T.PRIMARY, 55);
  s.addText("夜晚的房间", { x: 0, y: 4.35, w: W, h: 1.0, fontSize: 40, bold: true,
    color: T.TEXT, fontFace: F, align: "center", charSpacing: 6, margin: 0 });
  s.addText("品牌视频 · 30 秒", { x: 0, y: 5.5, w: W, h: 0.35, fontSize: 14,
    color: T.MUTED, fontFace: F, align: "center", margin: 0 });
  pageTag(s, 2);
}

// —— p3 理念页 ——
{
  const s = p.addSlide();
  s.background = { color: T.BG };
  s.addText("让家，有自己的天气。", { x: 0, y: 3.1, w: W, h: 1.2, fontSize: 48, bold: true,
    color: T.PRIMARY, fontFace: F, align: "center", charSpacing: 4, margin: 0 });
  pageTag(s, 3);
  s.addNotes("停顿三秒再翻页，让这句话落地。");
}

// —— p4 三重感知 ——
{
  const s = p.addSlide();
  s.background = { color: T.BG };
  const rows = [
    ["气味", "24 种香型，随时间与心情流转"],
    ["光线", "2700K 暖光，跟随日落渐暗"],
    ["声音", "28dB 静音，比图书馆更轻"],
  ];
  rows.forEach((r, i) => {
    const y = 1.85 + i * 1.65;
    s.addText(r[0], { x: M, y, w: 2.6, h: 1.0, fontSize: 40, bold: true,
      color: T.PRIMARY, fontFace: F, margin: 0 });
    s.addText(r[1], { x: 4.1, y: y + 0.28, w: 8.2, h: 0.6, fontSize: 18,
      color: T.MUTED, fontFace: F, margin: 0 });
    if (i < 2) s.addShape(p.shapes.LINE, { x: M, y: y + 1.32, w: W - 2 * M, h: 0,
      line: { color: T.LINE, width: 0.75 } });
  });
  pageTag(s, 4);
  s.addNotes("三重感知是产品主结构，每行 15 秒。");
}

// —— p5 数字页 ——
{
  const s = p.addSlide();
  s.background = { color: T.BG };
  const stats = [
    ["28dB", "运行噪音 · 比图书馆更安静"],
    ["60 天", "单次香氛续航 · 一季只加一次"],
    ["24 种", "香型 · 覆盖 6 大香系"],
  ];
  stats.forEach((st, i) => {
    const x = M + i * 3.95;
    s.addText(st[0], { x, y: 2.7, w: 3.6, h: 1.3, fontSize: 64, bold: true,
      color: i === 1 ? T.ACCENT : T.PRIMARY, fontFace: F, margin: 0 });
    s.addText(st[1], { x, y: 4.15, w: 3.6, h: 0.6, fontSize: 14, color: T.MUTED,
      fontFace: F, margin: 0 });
  });
  s.addShape(p.shapes.LINE, { x: M, y: 2.35, w: W - 2 * M, h: 0,
    line: { color: T.LINE, width: 0.75 } });
  pageTag(s, 5);
}

// —— p6 气味深读 ——
{
  const s = p.addSlide();
  s.background = { color: T.BG };
  s.addText("24", { x: M, y: 1.9, w: 3.6, h: 2.4, fontSize: 120, bold: true,
    color: T.PRIMARY, fontFace: F, margin: 0 });
  s.addText("种香型，6 大香系", { x: M, y: 4.5, w: 3.8, h: 0.5, fontSize: 18,
    color: T.TEXT, fontFace: F, margin: 0 });
  const families = [["雪松", "木质 / 沉静"], ["柑橘", "清新 / 提神"], ["墨苔", "湿润 / 深睡"],
    ["花香", "柔暖 / 晚餐后"], ["茶意", "回甘 / 阅读"], ["琥珀", "包裹 / 冬夜"]];
  families.forEach((f2, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 5.4 + col * 3.6, y = 1.95 + row * 1.45;
    s.addText(f2[0], { x, y, w: 3.2, h: 0.5, fontSize: 22, bold: true,
      color: T.TEXT, fontFace: F, margin: 0 });
    s.addText(f2[1], { x, y: y + 0.55, w: 3.2, h: 0.4, fontSize: 13,
      color: T.MUTED, fontFace: F, margin: 0 });
    if (row < 2) s.addShape(p.shapes.LINE, { x, y: y + 1.12, w: 3.0, h: 0,
      line: { color: T.LINE, width: 0.75 } });
  });
  pageTag(s, 6);
}

// —— p7 价格页 ——
{
  const s = p.addSlide();
  s.background = { color: T.BG };
  moon(s, W / 2, 1.95, 0.9, T.ACCENT, 30);
  s.addText("首发价", { x: 0, y: 2.75, w: W, h: 0.5, fontSize: 20, color: T.MUTED,
    fontFace: F, align: "center", charSpacing: 8, margin: 0 });
  s.addText("¥ 799", { x: 0, y: 3.3, w: W, h: 1.9, fontSize: 110, bold: true,
    color: T.ACCENT, fontFace: F, align: "center", margin: 0 });
  s.addText([
    { text: "日常价 999 元", options: { strike: true, breakLine: true } },
    { text: "首发限时 · 天猫直播间", options: {} },
  ], { x: 0, y: 5.6, w: W, h: 0.9, fontSize: 16, color: T.MUTED, fontFace: F,
    align: "center", paraSpaceAfter: 6, margin: 0 });
  pageTag(s, 7);
  s.addNotes("报价后停顿，不解释，等反应。");
}

// —— p8 预售信息 ——
{
  const s = p.addSlide();
  s.background = { color: T.BG };
  const rows = [
    ["渠道", "天猫旗舰店 · 首发直播间"],
    ["时间", "时间未定 · 以官网公告为准"],
    ["权益", "首发赠 2 支装香芯 · 延保一年"],
  ];
  rows.forEach((r, i) => {
    const y = 2.3 + i * 1.35;
    s.addText(r[0], { x: M, y, w: 1.6, h: 0.5, fontSize: 18, bold: true,
      color: T.PRIMARY, fontFace: F, margin: 0 });
    s.addText(r[1], { x: 2.9, y: y + 0.03, w: 9.2, h: 0.5, fontSize: 18,
      color: i === 1 ? T.MUTED : T.TEXT, fontFace: F, margin: 0 });
    if (i < 2) s.addShape(p.shapes.LINE, { x: M, y: y + 0.95, w: W - 2 * M, h: 0,
      line: { color: T.LINE, width: 0.75 } });
  });
  pageTag(s, 8);
}

// —— p9 收束情绪页 ——
{
  const s = p.addSlide();
  s.background = { color: T.BG };
  moon(s, W / 2, 2.35, 1.1, T.PRIMARY, 40);
  s.addText("让家，有自己的天气。", { x: 0, y: 3.6, w: W, h: 1.1, fontSize: 44, bold: true,
    color: T.PRIMARY, fontFace: F, align: "center", charSpacing: 4, margin: 0 });
  pageTag(s, 9);
}

// —— p10 封底 ——
{
  const s = p.addSlide();
  s.background = { color: T.BG };
  s.addText("星尘 STARDUST", { x: 0, y: 3.15, w: W, h: 0.9, fontSize: 36, bold: true,
    color: T.PRIMARY, fontFace: F, align: "center", charSpacing: 10, margin: 0 });
  s.addText("谢谢 · 现场体验区在展厅东侧", { x: 0, y: 4.25, w: W, h: 0.4, fontSize: 14,
    color: T.MUTED, fontFace: F, align: "center", margin: 0 });
}

p.writeFile({ fileName: "星尘ST2-发布会.pptx" }).then(f => console.log("written:", f));
