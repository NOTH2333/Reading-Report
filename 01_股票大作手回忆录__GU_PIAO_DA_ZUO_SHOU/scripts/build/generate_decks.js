const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const {
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
} = require("../helpers/pptxgenjs/layout");

const ROOT = path.resolve(__dirname, "..", "..");
const data = JSON.parse(
  fs.readFileSync(
    path.join(ROOT, "data", "metadata", "book_content.json"),
    "utf8"
  )
);

const OUT_DIR = path.join(ROOT, "output", "ppt");
fs.mkdirSync(OUT_DIR, { recursive: true });
const W = 13.333;
const H = 7.5;
const FONT = "Microsoft YaHei";
const FONT_ALT = "Georgia";
const SHAPE = new pptxgen().ShapeType;

const COLORS = {
  forest: "123524",
  forest2: "1D4D38",
  cream: "F5F0E1",
  paper: "EFE6CE",
  ink: "142013",
  slate: "5E6A64",
  copper: "B46A3C",
  red: "9A3D2E",
  gold: "C99D3A",
};

function stageColor(stage) {
  const map = {
    起步试错期: COLORS.copper,
    方法成形期: COLORS.gold,
    扩张与回撤期: COLORS.red,
    专业化总结期: COLORS.forest2,
    收官警示期: "5B4D8A",
  };
  return map[stage] || COLORS.copper;
}

function safeName(name) {
  return name.replace(/[\\/:*?"<>|]/g, "_");
}

function bulletRuns(items) {
  return items.map((text, idx) => ({
    text,
    options: { bullet: true, breakLine: idx < items.length - 1 },
  }));
}

function addBg(slide, bgPath) {
  slide.addImage({ path: bgPath, x: 0, y: 0, w: W, h: H });
}

function addHeader(slide, chapterLabel, slideTitle, stage, pageLabel) {
  slide.addShape(SHAPE.rect, {
    x: 0,
    y: 0,
    w: W,
    h: 0.55,
    fill: { color: stageColor(stage) },
    line: { color: stageColor(stage) },
  });
  slide.addText(chapterLabel, {
    x: 0.34,
    y: 0.1,
    w: 3.8,
    h: 0.24,
    fontFace: FONT,
    fontSize: 19,
    bold: true,
    color: COLORS.cream,
    margin: 0,
  });
  slide.addText(slideTitle, {
    x: 4.2,
    y: 0.1,
    w: 6.5,
    h: 0.24,
    fontFace: FONT,
    fontSize: 18,
    bold: true,
    color: COLORS.cream,
    margin: 0,
    align: "center",
  });
  slide.addText(pageLabel, {
    x: 11.1,
    y: 0.1,
    w: 1.85,
    h: 0.24,
    fontFace: FONT,
    fontSize: 11,
    color: COLORS.cream,
    margin: 0,
    align: "right",
  });
}

function addFooter(slide, text) {
  slide.addText(text, {
    x: 0.42,
    y: 7.08,
    w: 12.4,
    h: 0.16,
    fontFace: FONT,
    fontSize: 9,
    color: COLORS.slate,
    margin: 0,
    align: "right",
  });
}

function addRoundedBox(slide, x, y, w, h, fill, outline = "D7C8A8", radius = 0.08) {
  slide.addShape(SHAPE.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: radius,
    fill: { color: fill },
    line: { color: outline, width: 1 },
  });
}

function addCardTitle(slide, text, x, y, w) {
  slide.addText(text, {
    x,
    y,
    w,
    h: 0.28,
    fontFace: FONT,
    fontSize: 17,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
}

function applyWarnings(slide, pptx) {
  warnIfSlideHasOverlaps(slide, pptx, { muteContainment: true, ignoreLines: false });
  warnIfSlideElementsOutOfBounds(slide, pptx);
}

function newDeck(title) {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "OpenAI Codex";
  pptx.company = "OpenAI";
  pptx.subject = "《股票大作手回忆录》章节拆解";
  pptx.title = title;
  pptx.lang = "zh-CN";
  pptx.theme = { headFontFace: FONT, bodyFontFace: FONT, lang: "zh-CN" };
  return pptx;
}

function chapterLabel(chapter) {
  return `第${String(chapter.chapter_no).padStart(2, "0")}章 · ${chapter.stage}`;
}

function pageLabel(item) {
  return `注疏版 p.${item.start_page}-${item.end_page}`;
}

function addBulletText(slide, items, x, y, w, h, opts = {}) {
  slide.addText(bulletRuns(items), {
    x,
    y,
    w,
    h,
    fontFace: FONT,
    fontSize: opts.fontSize || 18,
    color: opts.color || COLORS.ink,
    margin: opts.margin || 0.08,
    valign: "top",
    breakLine: true,
    paraSpaceAfterPt: opts.paraSpaceAfterPt || 11,
  });
}

function addImageContain(slide, imgPath, x, y, w, h) {
  if (!imgPath || !fs.existsSync(imgPath)) return;
  slide.addImage({ path: imgPath, x, y, w, h, sizing: { type: "contain", w, h } });
}

function buildChapterDeck(chapter) {
  const pptx = newDeck(`${chapter.chapter_no}. ${chapter.title}`);
  const theme = data.theme;
  const stage = chapter.stage;
  const accent = stageColor(stage);

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.cover);
    slide.addText("《股票大作手回忆录》", {
      x: 0.9, y: 0.95, w: 3.4, h: 0.25,
      fontFace: FONT, fontSize: 15, color: "D6E8D6", margin: 0,
    });
    slide.addText(chapter.title, {
      x: 0.88, y: 1.48, w: 5.15, h: 1.95,
      fontFace: FONT, fontSize: 25, bold: true, color: COLORS.cream,
      margin: 0, valign: "mid",
    });
    slide.addText(chapter.one_line, {
      x: 0.9, y: 3.58, w: 4.4, h: 1.3,
      fontFace: FONT, fontSize: 18, color: "D6E8D6", margin: 0,
    });
    slide.addText(`阶段：${stage}`, {
      x: 0.9, y: 5.25, w: 2.7, h: 0.22,
      fontFace: FONT, fontSize: 13, bold: true, color: COLORS.cream, margin: 0,
    });
    slide.addText(`页码：${chapter.start_page}-${chapter.end_page}`, {
      x: 0.9, y: 5.62, w: 2.9, h: 0.22,
      fontFace: FONT, fontSize: 12, color: COLORS.cream, margin: 0,
    });
    addImageContain(slide, chapter.concept_card, 7.2, 1.85, 4.95, 4.1);
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, chapterLabel(chapter), "这一章一句话讲明白", stage, pageLabel(chapter));
    slide.addText(chapter.one_line, {
      x: 0.65, y: 0.98, w: 12.0, h: 0.65,
      fontFace: FONT, fontSize: 26, bold: true, color: COLORS.forest,
      align: "center", margin: 0,
    });
    const xs = [0.7, 4.55, 8.4];
    chapter.core_points.forEach((point, idx) => {
      addRoundedBox(slide, xs[idx], 2.0, 3.35, 3.05, "FBF7EC");
      slide.addShape(SHAPE.rect, {
        x: xs[idx], y: 2.0, w: 3.35, h: 0.18,
        fill: { color: accent }, line: { color: accent },
      });
      slide.addText(`核心点 ${idx + 1}`, {
        x: xs[idx] + 0.18, y: 2.26, w: 1.6, h: 0.2,
        fontFace: FONT, fontSize: 16, bold: true, color: COLORS.ink, margin: 0,
      });
      slide.addText(point, {
        x: xs[idx] + 0.18, y: 2.74, w: 2.98, h: 1.18,
        fontFace: FONT, fontSize: 20, bold: true, color: COLORS.forest, margin: 0,
      });
      slide.addText(chapter.actions[idx] || "", {
        x: xs[idx] + 0.18, y: 4.05, w: 2.94, h: 0.76,
        fontFace: FONT, fontSize: 15, color: COLORS.slate, margin: 0,
      });
    });
    addFooter(slide, `规则：${chapter.rule}`);
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, chapterLabel(chapter), "它在整本书里的位置", stage, pageLabel(chapter));
    addImageContain(slide, data.meta.roadmap_image, 0.48, 1.02, 9.2, 5.96);
    addRoundedBox(slide, 9.9, 1.4, 2.95, 4.85, "FBF7EC");
    slide.addText(`当前章节\n第${chapter.chapter_no}章`, {
      x: 10.22, y: 1.72, w: 2.2, h: 0.95,
      fontFace: FONT, fontSize: 24, bold: true, color: COLORS.forest,
      align: "center", margin: 0,
    });
    slide.addText(chapter.stage, {
      x: 10.18, y: 2.95, w: 2.3, h: 0.28,
      fontFace: FONT, fontSize: 18, bold: true, color: accent,
      align: "center", margin: 0,
    });
    addBulletText(slide, [
      `主题：${chapter.title}`,
      `一句话：${chapter.one_line}`,
      `关键原则：${chapter.principle}`,
    ], 10.14, 3.5, 2.35, 2.05, { fontSize: 14 });
    addFooter(slide, "把这一章放回全书路径里，零基础读者更容易理解它为什么重要。");
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, chapterLabel(chapter), "时代和市场背景", stage, pageLabel(chapter));
    addRoundedBox(slide, 0.55, 1.12, 6.2, 5.75, "FBF7EC");
    addCardTitle(slide, "这章发生时，市场处在什么环境里？", 0.82, 1.38, 5.2);
    addBulletText(slide, chapter.market_context, 0.8, 1.9, 5.45, 4.4, { fontSize: 19 });
    addImageContain(slide, chapter.visual_image, 7.05, 1.25, 5.7, 5.45);
    addFooter(slide, "重点不是记年份，而是看懂环境如何改变交易的胜率与打法。");
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, chapterLabel(chapter), "事件时间线", stage, pageLabel(chapter));
    slide.addShape(SHAPE.line, {
      x: 1.0, y: 3.55, w: 11.1, h: 0,
      line: { color: accent, width: 2.5 },
    });
    const xs = [0.9, 3.85, 6.8, 9.75];
    chapter.storyline.forEach((item, idx) => {
      slide.addShape(SHAPE.ellipse, {
        x: xs[idx] + 0.62, y: 3.22, w: 0.42, h: 0.42,
        fill: { color: accent }, line: { color: accent },
      });
      addRoundedBox(slide, xs[idx], idx % 2 === 0 ? 1.35 : 4.0, 2.55, 1.62, "FBF7EC");
      slide.addText(`0${idx + 1}`.slice(-2), {
        x: xs[idx] + 0.18, y: (idx % 2 === 0 ? 1.52 : 4.16), w: 0.5, h: 0.2,
        fontFace: FONT_ALT, fontSize: 18, bold: true, color: accent, margin: 0,
      });
      slide.addText(item, {
        x: xs[idx] + 0.22, y: (idx % 2 === 0 ? 1.84 : 4.48), w: 2.05, h: 0.82,
        fontFace: FONT, fontSize: 15, color: COLORS.ink, margin: 0,
      });
    });
    addFooter(slide, "先按顺序看清发生了什么，再谈方法和教训。");
    applyWarnings(slide, pptx);
  }
  chapter.key_trades.forEach((trade, idx) => {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, chapterLabel(chapter), `关键交易 ${idx + 1}`, stage, pageLabel(chapter));
    addRoundedBox(slide, 0.72, 1.06, 12.0, 5.95, "FBF7EC");
    slide.addShape(SHAPE.rect, {
      x: 0.72, y: 1.06, w: 12.0, h: 0.28,
      fill: { color: accent }, line: { color: accent },
    });
    slide.addText(trade.title, {
      x: 1.05, y: 1.6, w: 5.35, h: 0.52,
      fontFace: FONT, fontSize: 24, bold: true, color: COLORS.forest, margin: 0,
    });
    slide.addText(trade.detail, {
      x: 1.05, y: 2.15, w: 5.7, h: 0.75,
      fontFace: FONT, fontSize: 18, color: COLORS.slate, margin: 0,
    });
    addBulletText(slide, [
      chapter.storyline[idx] || chapter.storyline[0],
      chapter.market_context[Math.min(idx, chapter.market_context.length - 1)],
      chapter.actions[idx] || chapter.actions[0],
    ], 1.0, 3.0, 5.55, 2.55, { fontSize: 18 });
    addImageContain(slide, idx === 2 ? chapter.concept_card : chapter.visual_image, 7.35, 1.88, 4.65, 4.15);
    addFooter(slide, "把关键动作拆开看，能帮助新手理解‘为什么这一步重要’。");
    applyWarnings(slide, pptx);
  });

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, chapterLabel(chapter), "这一章最容易犯的错", stage, pageLabel(chapter));
    const ys = [1.55, 3.1, 4.65];
    chapter.mistakes.forEach((item, idx) => {
      addRoundedBox(slide, 0.9, ys[idx], 11.55, 1.08, "FAEEE9", "E4C0B7");
      slide.addShape(SHAPE.ellipse, {
        x: 1.15, y: ys[idx] + 0.24, w: 0.38, h: 0.38,
        fill: { color: COLORS.red }, line: { color: COLORS.red },
      });
      slide.addText(item, {
        x: 1.78, y: ys[idx] + 0.22, w: 10.15, h: 0.45,
        fontFace: FONT, fontSize: 20, color: COLORS.ink, margin: 0,
      });
    });
    slide.addText(`误区：${chapter.misconception}`, {
      x: 0.95, y: 6.1, w: 11.8, h: 0.44,
      fontFace: FONT, fontSize: 15, italic: true, color: COLORS.slate, margin: 0,
    });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, chapterLabel(chapter), "他做对了什么", stage, pageLabel(chapter));
    const xs = [0.8, 4.48, 8.16];
    chapter.strengths.forEach((item, idx) => {
      addRoundedBox(slide, xs[idx], 1.85, 3.05, 3.9, "ECF4EC", "B9D0BF");
      slide.addShape(SHAPE.rect, {
        x: xs[idx], y: 1.85, w: 3.05, h: 0.18,
        fill: { color: accent }, line: { color: accent },
      });
      slide.addText(`优势 ${idx + 1}`, {
        x: xs[idx] + 0.18, y: 2.18, w: 1.2, h: 0.2,
        fontFace: FONT, fontSize: 16, bold: true, color: COLORS.forest, margin: 0,
      });
      slide.addText(item, {
        x: xs[idx] + 0.18, y: 2.65, w: 2.62, h: 2.38,
        fontFace: FONT, fontSize: 17, color: COLORS.ink, margin: 0,
      });
    });
    slide.addText(`关键原则：${chapter.principle}`, {
      x: 0.9, y: 6.15, w: 11.8, h: 0.35,
      fontFace: FONT, fontSize: 17, bold: true, color: accent, margin: 0, align: "center",
    });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, chapterLabel(chapter), "术语拆解：给零基础读者", stage, pageLabel(chapter));
    const xs = [0.72, 4.47, 8.22];
    chapter.terms.forEach((term, idx) => {
      addRoundedBox(slide, xs[idx], 1.55, 3.2, 4.5, "FBF7EC");
      slide.addText(term.name, {
        x: xs[idx] + 0.2, y: 1.92, w: 2.8, h: 0.3,
        fontFace: FONT, fontSize: 20, bold: true, color: accent, margin: 0, align: "center",
      });
      slide.addShape(SHAPE.line, {
        x: xs[idx] + 0.35, y: 2.42, w: 2.5, h: 0,
        line: { color: "D7C8A8", width: 1.5 },
      });
      slide.addText(term.plain, {
        x: xs[idx] + 0.22, y: 2.78, w: 2.75, h: 1.55,
        fontFace: FONT, fontSize: 17, color: COLORS.ink, margin: 0,
        align: "center", valign: "mid",
      });
    });
    slide.addText("把概念翻成白话，才算真正读懂了这章。", {
      x: 0.9, y: 6.25, w: 11.4, h: 0.25,
      fontFace: FONT, fontSize: 15, italic: true, color: COLORS.slate,
      margin: 0, align: "center",
    });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, chapterLabel(chapter), "图表和配图怎么读", stage, pageLabel(chapter));
    addImageContain(slide, chapter.visual_image, 0.72, 1.22, 6.2, 5.42);
    addRoundedBox(slide, 7.2, 1.22, 5.4, 5.42, "FBF7EC");
    slide.addText("图里最该看什么？", {
      x: 7.52, y: 1.55, w: 3.6, h: 0.28,
      fontFace: FONT, fontSize: 20, bold: true, color: COLORS.forest, margin: 0,
    });
    addBulletText(slide, [
      chapter.chart_takeaway,
      `如果把它翻成新手能懂的话：${chapter.one_line}`,
      `相关页码：${chapter.start_page}-${chapter.end_page}${chapter.chart_image ? "（含书内图表/插图）" : "（使用自制记忆图）"}`,
    ], 7.45, 2.05, 4.7, 3.6, { fontSize: 17 });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, chapterLabel(chapter), "心理课：人最容易在哪一步自毁", stage, pageLabel(chapter));
    slide.addText(chapter.psychology, {
      x: 0.88, y: 1.28, w: 11.6, h: 1.2,
      fontFace: FONT, fontSize: 23, bold: true, color: COLORS.forest,
      margin: 0, align: "center",
    });
    addRoundedBox(slide, 0.95, 2.8, 5.6, 2.55, "FBF7EC");
    addCardTitle(slide, "最常见误区", 1.2, 3.05, 3.0);
    slide.addText(chapter.misconception, {
      x: 1.2, y: 3.45, w: 4.85, h: 1.35,
      fontFace: FONT, fontSize: 18, color: COLORS.ink, margin: 0,
    });
    addRoundedBox(slide, 6.8, 2.8, 5.6, 2.55, "FBF7EC");
    addCardTitle(slide, "更容易记住的类比", 7.05, 3.05, 3.0);
    slide.addText(chapter.analogy, {
      x: 7.05, y: 3.45, w: 4.9, h: 1.35,
      fontFace: FONT, fontSize: 18, color: COLORS.ink, margin: 0,
    });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, chapterLabel(chapter), "风控课：这章提醒你防什么", stage, pageLabel(chapter));
    addRoundedBox(slide, 0.8, 1.3, 12.0, 1.1, "FAEEE9", "E4C0B7");
    slide.addText(chapter.risk, {
      x: 1.0, y: 1.62, w: 11.45, h: 0.4,
      fontFace: FONT, fontSize: 19, bold: true, color: COLORS.red, margin: 0, align: "center",
    });
    addRoundedBox(slide, 0.8, 2.8, 5.8, 3.1, "FBF7EC");
    addCardTitle(slide, "可以立刻执行的做法", 1.05, 3.05, 3.2);
    addBulletText(slide, chapter.actions, 1.05, 3.48, 4.95, 2.15, { fontSize: 17 });
    addRoundedBox(slide, 6.75, 2.8, 6.0, 3.1, "FBF7EC");
    addCardTitle(slide, "一句规则先记住", 7.0, 3.05, 3.2);
    slide.addText(chapter.rule, {
      x: 7.08, y: 3.55, w: 5.2, h: 1.0,
      fontFace: FONT, fontSize: 23, bold: true, color: COLORS.forest,
      margin: 0, align: "center", valign: "mid",
    });
    slide.addText(`为什么：${chapter.principle}`, {
      x: 7.15, y: 4.75, w: 5.05, h: 0.85,
      fontFace: FONT, fontSize: 16, color: COLORS.slate, margin: 0, align: "center",
    });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, chapterLabel(chapter), "如果你是今天的新手", stage, pageLabel(chapter));
    slide.addText(chapter.modern, {
      x: 0.85, y: 1.2, w: 11.6, h: 0.9,
      fontFace: FONT, fontSize: 23, bold: true, color: COLORS.forest,
      margin: 0, align: "center",
    });
    const xs = [1.0, 4.3, 7.6, 10.9];
    chapter.actions.forEach((item, idx) => {
      addRoundedBox(slide, xs[idx] - 0.45, 2.75, 2.35, 2.55, "FBF7EC");
      slide.addShape(SHAPE.ellipse, {
        x: xs[idx] + 0.35, y: 2.95, w: 0.6, h: 0.6,
        fill: { color: accent }, line: { color: accent },
      });
      slide.addText(String(idx + 1), {
        x: xs[idx] + 0.52, y: 3.12, w: 0.22, h: 0.16,
        fontFace: FONT_ALT, fontSize: 14, bold: true, color: COLORS.cream, margin: 0,
      });
      slide.addText(item, {
        x: xs[idx] - 0.22, y: 3.78, w: 1.9, h: 1.1,
        fontFace: FONT, fontSize: 15, color: COLORS.ink, align: "center", margin: 0,
      });
    });
    slide.addText(`类比记忆：${chapter.analogy}`, {
      x: 1.0, y: 5.9, w: 11.25, h: 0.48,
      fontFace: FONT, fontSize: 16, color: COLORS.slate, margin: 0, align: "center",
    });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, chapterLabel(chapter), "记忆图：把这一章记成一张图", stage, pageLabel(chapter));
    addImageContain(slide, chapter.concept_card, 0.72, 1.18, 7.55, 5.85);
    addRoundedBox(slide, 8.62, 1.45, 3.95, 5.0, "FBF7EC");
    slide.addText("为什么这张图值得记？", {
      x: 8.88, y: 1.78, w: 3.2, h: 0.28,
      fontFace: FONT, fontSize: 19, bold: true, color: COLORS.forest, margin: 0,
    });
    addBulletText(slide, [
      chapter.one_line,
      chapter.core_points[0],
      chapter.core_points[1],
      chapter.core_points[2],
    ], 8.84, 2.28, 3.0, 2.6, { fontSize: 16 });
    slide.addText(`规则：${chapter.rule}`, {
      x: 8.92, y: 5.3, w: 2.95, h: 0.75,
      fontFace: FONT, fontSize: 16, bold: true, color: accent, margin: 0,
    });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, chapterLabel(chapter), "关键短句", stage, pageLabel(chapter));
    slide.addText(`“${chapter.excerpt}”`, {
      x: 0.95, y: 1.6, w: 11.4, h: 0.9,
      fontFace: FONT, fontSize: 28, bold: true, color: COLORS.forest,
      align: "center", margin: 0,
    });
    slide.addText(`书中对应页码：p.${chapter.quote_page}`, {
      x: 0.95, y: 2.65, w: 11.3, h: 0.22,
      fontFace: FONT, fontSize: 12, color: COLORS.slate, align: "center", margin: 0,
    });
    addRoundedBox(slide, 1.05, 3.25, 11.1, 2.45, "FBF7EC");
    addBulletText(slide, [
      `这句短话为什么重要：${chapter.principle}`,
      `它背后的心理提醒：${chapter.psychology}`,
      `如果你只能记住一句，就记住：${chapter.rule}`,
    ], 1.35, 3.6, 10.3, 1.7, { fontSize: 17 });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.closing);
    slide.addText(`${chapterLabel(chapter)} · 行动清单`, {
      x: 0.82, y: 0.72, w: 5.5, h: 0.3,
      fontFace: FONT, fontSize: 24, bold: true, color: COLORS.cream, margin: 0,
    });
    slide.addText(chapter.rule, {
      x: 0.82, y: 1.28, w: 5.25, h: 1.1,
      fontFace: FONT, fontSize: 28, bold: true, color: "D6E8D6", margin: 0,
    });
    addImageContain(slide, chapter.concept_card, 7.3, 0.95, 5.2, 4.2);
    addRoundedBox(slide, 0.78, 3.05, 5.9, 3.45, "133826", "335945");
    slide.addText("现在就能做的 4 件事", {
      x: 1.05, y: 3.35, w: 2.8, h: 0.22,
      fontFace: FONT, fontSize: 18, bold: true, color: COLORS.cream, margin: 0,
    });
    addBulletText(slide, chapter.actions, 1.02, 3.8, 5.05, 2.15, {
      fontSize: 17, color: COLORS.cream, margin: 0.05,
    });
    slide.addText(`最后提醒：${chapter.misconception}`, {
      x: 7.05, y: 5.45, w: 5.4, h: 0.85,
      fontFace: FONT, fontSize: 16, color: COLORS.cream, margin: 0,
    });
    applyWarnings(slide, pptx);
  }

  return pptx;
}

function buildAppendixDeck() {
  const appendix = data.appendix;
  const theme = data.theme;
  const stage = "专业化总结期";
  const accent = stageColor(stage);
  const pptx = newDeck("附录合辑");
  const enterRules = [
    "顺上涨股做多，顺下跌股做空。",
    "只有趋势明确时才交易，不要天天找单子。",
    "入场节奏要和关键价位、关键时间点匹配。",
    "等市场先证明你的观点，再迅速出手。",
    "一开仓就顺利的交易，往往更值得继续持有。",
  ];
  const holdRules = [
    "有利润就让它跑，直到趋势结束。",
    "原来支持你赚钱的趋势不在了，就退出。",
    "优先做板块里最强或最弱的领头羊。",
    "股价创新高可考虑顺势做多，创新低可考虑顺势做空。",
    "离场同样要有规则，别把浮盈当终局。",
  ];
  const banRules = [
    "绝不摊平亏损头寸。",
    "绝不靠追加保证金来拖延认错。",
    "别因为价格高就不敢买，也别因为价格低就不敢卖。",
    "股票跌坏了就放手，别把交易做成长期被套。",
    "市场永远比个人看法更有最终裁决权。",
  ];
  const leastResistanceBullets = [
    "市场会不断试探涨跌，就像水在微坡地面上边聚边试。",
    "成交价上涨，说明买方竞争更激烈；成交价下跌，说明卖方压力更大。",
    "当一边压力持续增强、另一边持续减弱，趋势就形成了。",
    "交易者不是去命令市场，而是去识别哪一边的阻力最小。",
  ];
  {
    const slide = pptx.addSlide();
    addBg(slide, theme.cover);
    slide.addText("附录合辑", {
      x: 0.9, y: 1.0, w: 3.2, h: 0.4,
      fontFace: FONT, fontSize: 30, bold: true, color: COLORS.cream, margin: 0,
    });
    slide.addText("年表、交易规则与最小阻力路线", {
      x: 0.92, y: 1.6, w: 5.0, h: 0.95,
      fontFace: FONT, fontSize: 25, bold: true, color: COLORS.cream, margin: 0,
    });
    slide.addText(appendix.one_line, {
      x: 0.95, y: 3.0, w: 4.9, h: 1.1,
      fontFace: FONT, fontSize: 18, color: "D6E8D6", margin: 0,
    });
    addImageContain(slide, appendix.concept_card, 7.15, 1.85, 4.9, 4.05);
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, "附录合辑", "附录为什么值得单独讲", stage, "注疏版 p.325-340");
    slide.addText(appendix.one_line, {
      x: 0.8, y: 1.15, w: 11.8, h: 0.6,
      fontFace: FONT, fontSize: 26, bold: true, color: COLORS.forest,
      align: "center", margin: 0,
    });
    const xs = [0.9, 4.62, 8.34];
    appendix.core_points.forEach((point, idx) => {
      addRoundedBox(slide, xs[idx], 2.1, 3.05, 3.4, "FBF7EC");
      slide.addText(point, {
        x: xs[idx] + 0.18, y: 2.55, w: 2.7, h: 1.2,
        fontFace: FONT, fontSize: 20, bold: true, color: COLORS.forest,
        margin: 0, align: "center",
      });
      slide.addText(idx === 0 ? "帮助你看见利弗莫尔的完整起伏周期。" : idx === 1 ? "把故事浓缩成可执行条款。" : "把趋势概念解释成人人看得懂的图像。", {
        x: xs[idx] + 0.18, y: 3.95, w: 2.65, h: 0.8,
        fontFace: FONT, fontSize: 15, color: COLORS.slate, margin: 0, align: "center",
      });
    });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, "附录一", "利弗莫尔生涯年表", stage, "注疏版 p.325-331");
    addImageContain(slide, appendix.chart_images[0] || appendix.concept_card, 0.7, 1.2, 6.3, 5.4);
    addRoundedBox(slide, 7.25, 1.2, 5.35, 5.4, "FBF7EC");
    addBulletText(slide, [
      "年表让读者看到：他的生涯不是线性上升，而是多次大起大落。",
      "真正值得学的不是神话战绩，而是每次失败后规则如何升级。",
      "把人物起伏放回时间线，能更好理解‘顺势、时机、纪律’为何反复出现。",
    ], 7.55, 1.55, 4.55, 3.4, { fontSize: 18 });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, "附录一", "把他的交易生涯拆成 5 个阶段", stage, "注疏版 p.325-331");
    addBulletText(slide, [
      "少年试错：靠观察和纸带阅读建立最早优势。",
      "方法升级：从抓小波动转到押大趋势。",
      "扩张与回撤：大成功伴随更大的诱惑与更贵的错误。",
      "成熟期：独立判断、持有纪律、风险意识更完整。",
      "收官期：把传奇经历压缩成警告与规则。",
    ], 0.95, 1.35, 11.0, 4.9, { fontSize: 21 });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, "附录二", "交易规则总览", stage, "注疏版 p.332-333");
    slide.addText("附录二最大的价值：把整本书讲过的经验，压缩成一个可以贴在屏幕旁边的清单。", {
      x: 0.9, y: 1.2, w: 11.6, h: 0.65,
      fontFace: FONT, fontSize: 23, bold: true, color: COLORS.forest,
      align: "center", margin: 0,
    });
    const groups = ["方向与时机", "持有与退出", "绝不能碰的禁区"];
    groups.forEach((g, idx) => {
      addRoundedBox(slide, 1.0 + idx * 4.05, 2.4, 3.35, 2.8, "FBF7EC");
      slide.addText(g, {
        x: 1.2 + idx * 4.05, y: 2.85, w: 2.95, h: 0.35,
        fontFace: FONT, fontSize: 22, bold: true, color: accent, align: "center", margin: 0,
      });
      slide.addText(idx === 0 ? "顺趋势、等确认、抓关键点。" : idx === 1 ? "有利润就让它跑，趋势结束就出。" : "不摊平、不补保证金、不跟价格讲面子。", {
        x: 1.25 + idx * 4.05, y: 3.5, w: 2.8, h: 1.05,
        fontFace: FONT, fontSize: 16, color: COLORS.ink, align: "center", margin: 0,
      });
    });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, "附录二", "规则组 1：方向与时机", stage, "注疏版 p.332-333");
    addBulletText(slide, enterRules, 0.95, 1.3, 11.1, 5.8, { fontSize: 19 });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, "附录二", "规则组 2：持有与退出", stage, "注疏版 p.332-333");
    addBulletText(slide, holdRules, 0.95, 1.3, 11.1, 5.8, { fontSize: 19 });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, "附录二", "规则组 3：禁区", stage, "注疏版 p.332-333");
    addBulletText(slide, banRules, 0.95, 1.3, 11.1, 5.8, { fontSize: 19 });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, "附录二", "这些规则为什么难做到", stage, "注疏版 p.332-333");
    addBulletText(slide, [
      "因为人天生想提前知道答案，不想等市场确认。",
      "因为人天生舍不得止损，也舍不得在盈利中耐心持有。",
      "因为人一旦被套，就会想通过摊平和补钱证明自己没错。",
      "因为规则朴素得像常识，人们反而低估了它们的杀伤力与价值。",
    ], 0.95, 1.45, 11.1, 5.4, { fontSize: 20 });
    applyWarnings(slide, pptx);
  }
  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, "附录三", "最小阻力路线是什么", stage, "注疏版 p.334-340");
    addImageContain(slide, appendix.chart_images[1] || appendix.concept_card, 0.7, 1.2, 5.9, 5.35);
    addBulletText(slide, leastResistanceBullets, 7.0, 1.45, 5.2, 4.8, { fontSize: 18 });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, "附录三", "趋势三种基本形态", stage, "注疏版 p.338-340");
    const blocks = [
      ["上升趋势", "更高的高点 + 更高的低点。"],
      ["下降趋势", "更低的低点 + 更低的高点。"],
      ["横向趋势", "没有明显新高，也没有明显新低。"],
    ];
    blocks.forEach((b, idx) => {
      addRoundedBox(slide, 0.9 + idx * 4.08, 2.05, 3.15, 3.25, "FBF7EC");
      slide.addText(b[0], {
        x: 1.12 + idx * 4.08, y: 2.45, w: 2.75, h: 0.34,
        fontFace: FONT, fontSize: 22, bold: true, color: accent, align: "center", margin: 0,
      });
      slide.addText(b[1], {
        x: 1.12 + idx * 4.08, y: 3.1, w: 2.75, h: 1.0,
        fontFace: FONT, fontSize: 17, color: COLORS.ink, align: "center", margin: 0,
      });
    });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, "附录三", "今天怎么把最小阻力路线用起来", stage, "注疏版 p.334-340");
    addBulletText(slide, [
      "先看价格在走更高高点/更高低点，还是更低低点/更低高点。",
      "看不清趋势时，默认仓位更轻，甚至不交易。",
      "趋势一旦明朗，再用规则决定进场、持有和退出。",
      "不要先爱上自己的观点，再去逼市场配合它。",
    ], 0.95, 1.45, 11.1, 5.2, { fontSize: 21 });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.content);
    addHeader(slide, "附录合辑", "一张图复习附录", stage, "注疏版 p.325-340");
    addImageContain(slide, appendix.concept_card, 0.8, 1.25, 7.0, 5.8);
    addRoundedBox(slide, 8.1, 1.55, 4.25, 4.8, "FBF7EC");
    addBulletText(slide, [
      "年表告诉你：高手也会反复摔倒。",
      "规则告诉你：真正能救命的原则都很朴素。",
      "最小阻力路线告诉你：别和市场争辩，先看它往哪边更容易走。",
    ], 8.35, 1.95, 3.45, 3.8, { fontSize: 17 });
    applyWarnings(slide, pptx);
  }

  {
    const slide = pptx.addSlide();
    addBg(slide, theme.closing);
    slide.addText("附录最终清单", {
      x: 0.85, y: 0.8, w: 3.5, h: 0.3,
      fontFace: FONT, fontSize: 26, bold: true, color: COLORS.cream, margin: 0,
    });
    slide.addText("读完传奇之后，真正要带走的是规则。", {
      x: 0.85, y: 1.38, w: 5.7, h: 0.5,
      fontFace: FONT, fontSize: 28, bold: true, color: "D6E8D6", margin: 0,
    });
    addRoundedBox(slide, 0.88, 2.35, 5.8, 3.8, "133826", "335945");
    addBulletText(slide, [
      "等市场证明，而不是先替市场下结论。",
      "有利润就想办法守住，趋势结束就走。",
      "绝不摊平亏损，绝不靠补钱证明自己。",
      "看不清最小阻力路线时，宁可轻仓或空仓。",
    ], 1.05, 2.75, 5.1, 2.8, { fontSize: 18, color: COLORS.cream, margin: 0.05 });
    addImageContain(slide, data.meta.roadmap_image, 7.2, 1.1, 5.2, 5.25);
    applyWarnings(slide, pptx);
  }

  return pptx;
}

async function writeDeck(pptx, fileName) {
  const outPath = path.join(OUT_DIR, fileName);
  await pptx.writeFile({ fileName: outPath });
  console.log(`wrote ${outPath}`);
}

async function main() {
  for (const chapter of data.chapters) {
    const deck = buildChapterDeck(chapter);
    await writeDeck(deck, `${String(chapter.chapter_no).padStart(2, "0")}_${safeName(chapter.title)}.pptx`);
  }
  const appendixDeck = buildAppendixDeck();
  await writeDeck(appendixDeck, "25_附录合辑_年表_规则_最小阻力路线.pptx");
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});


