"use strict";

const fs = require("fs");
const path = require("path");
const PptxGenJS = require("pptxgenjs");
const { imageSizingContain } = require("./pptxgenjs_helpers/image");
const {
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
} = require("./pptxgenjs_helpers/layout");
const { safeOuterShadow } = require("./pptxgenjs_helpers/util");
const SHAPE_TYPE = new PptxGenJS().ShapeType;

const ROOT = path.resolve(__dirname, "..");
const OUTPUT_ROOT = path.join(ROOT, "《量价分析》教学化交付包");
const PPT_ROOT = path.join(OUTPUT_ROOT, "01_PPT");
const DATASET_PATH = path.join(
  OUTPUT_ROOT,
  "05_中间数据",
  "章节教学内容.json"
);

const dataset = JSON.parse(fs.readFileSync(DATASET_PATH, "utf8"));
const palette = dataset.book_meta.design_theme;
const FONT = "Microsoft YaHei";
const TITLE_FONT = "Microsoft YaHei";

function hex(value) {
  return value.replace("#", "");
}

function makePpt(chapter) {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "OpenAI Codex";
  pptx.company = "OpenAI";
  pptx.subject = `《量价分析》${chapter.title} 教学化讲解`;
  pptx.title = `第${String(chapter.number).padStart(2, "0")}章 ${chapter.title}`;
  pptx.lang = "zh-CN";
  pptx.theme = {
    headFontFace: TITLE_FONT,
    bodyFontFace: FONT,
    lang: "zh-CN",
  };
  return pptx;
}

function baseSlide(slide, chapter, page, total) {
  slide.background = { color: hex(palette.bg) };
  slide.addShape(SHAPE_TYPE.line, {
    x: 0.45,
    y: 0.72,
    w: 12.2,
    h: 0,
    line: { color: hex(palette.primary), width: 1.6, beginArrowType: "none", endArrowType: "none" },
  });
  slide.addText(`第${String(chapter.number).padStart(2, "0")}章  ${chapter.title}`, {
    x: 0.45,
    y: 0.17,
    w: 8.4,
    h: 0.45,
    fontFace: TITLE_FONT,
    fontSize: 20,
    bold: true,
    color: hex(palette.primary),
  });
  slide.addText(`原书页码：${chapter.start_page}-${chapter.end_page}  |  第 ${page}/${total} 页`, {
    x: 9.2,
    y: 0.23,
    w: 3.6,
    h: 0.35,
    fontFace: FONT,
    fontSize: 9.5,
    align: "right",
    color: hex(palette.muted),
  });
  slide.addText("《量价分析》教学化交付", {
    x: 0.5,
    y: 7.03,
    w: 4.2,
    h: 0.22,
    fontFace: FONT,
    fontSize: 8.2,
    color: hex(palette.muted),
  });
}

function finalizeSlide(slide, pptx) {
  warnIfSlideHasOverlaps(slide, pptx, { ignoreLines: true, ignoreDecorativeShapes: true });
  warnIfSlideElementsOutOfBounds(slide, pptx);
}

function addSectionTitle(slide, title, subtitle) {
  slide.addText(title, {
    x: 0.7,
    y: 0.95,
    w: 6.5,
    h: 0.45,
    fontFace: TITLE_FONT,
    fontSize: 25,
    bold: true,
    color: hex(palette.primary),
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.72,
      y: 1.42,
      w: 11.6,
      h: 0.4,
      fontFace: FONT,
      fontSize: 11,
      color: hex(palette.muted),
    });
  }
}

function addBulletList(slide, items, box, opts = {}) {
  const runs = [];
  items.forEach((item) => {
    runs.push({
      text: item,
      options: {
        bullet: { indent: 14 },
        hanging: 3,
        breakLine: true,
      },
    });
  });
  slide.addText(runs, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    fontFace: FONT,
    fontSize: opts.fontSize || 16,
    color: hex(opts.color || palette.text),
    valign: "top",
    margin: 0.08,
    breakLine: false,
    paraSpaceAfterPt: opts.paraSpaceAfterPt || 10,
  });
}

function addCard(slide, x, y, w, h, title, body, theme = "primary") {
  const stroke = theme === "accent" ? palette.accent : theme === "secondary" ? palette.secondary : palette.primary;
  const fill = theme === "accent" ? "FFF3E4" : theme === "secondary" ? "E8F4F3" : "EAF1F6";
  slide.addShape(SHAPE_TYPE.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    line: { color: hex(stroke), width: 1.4 },
    fill: { color: fill },
    shadow: safeOuterShadow("64748B", 0.12, 45, 1.5, 1),
  });
  slide.addText(title, {
    x: x + 0.16,
    y: y + 0.12,
    w: w - 0.32,
    h: 0.3,
    fontFace: TITLE_FONT,
    fontSize: 16,
    bold: true,
    color: hex(stroke),
  });
  slide.addText(body, {
    x: x + 0.16,
    y: y + 0.48,
    w: w - 0.32,
    h: h - 0.58,
    fontFace: FONT,
    fontSize: 11.5,
    color: hex(palette.text),
    valign: "top",
    margin: 0.04,
  });
}

function chapterImagePath(chapterRoot, figure) {
  return path.join(chapterRoot, "images", path.basename(figure.image_path));
}

function writeStubScript(chapter) {
  const chapterRoot = path.join(PPT_ROOT, chapter.folder);
  const scriptPath = path.join(chapterRoot, "生成本章PPT.js");
  const content = `"use strict";\nconst { buildSingle } = require("../../scripts/build_chapter_ppts");\nbuildSingle(${chapter.number}).catch((error) => {\n  console.error(error);\n  process.exit(1);\n});\n`;
  fs.writeFileSync(scriptPath, content, "utf8");
}

function addFigureSlide(slide, chapter, figure, chapterRoot, label) {
  addSectionTitle(slide, `${label} 原书图讲解`, figure.caption);
  const imagePath = chapterImagePath(chapterRoot, figure);
  if (fs.existsSync(imagePath)) {
    slide.addImage({
      path: imagePath,
      ...imageSizingContain(imagePath, 0.75, 1.85, 6.6, 4.9),
    });
  }
  slide.addShape(SHAPE_TYPE.roundRect, {
    x: 7.75,
    y: 1.95,
    w: 4.75,
    h: 4.55,
    rectRadius: 0.08,
    line: { color: hex(palette.secondary), width: 1.2 },
    fill: { color: "FFFFFF" },
  });
  slide.addText("怎么看这张图", {
    x: 7.98,
    y: 2.14,
    w: 2.3,
    h: 0.3,
    fontFace: TITLE_FONT,
    fontSize: 17,
    bold: true,
    color: hex(palette.secondary),
  });
  addBulletList(
    slide,
    [
      figure.explanation,
      `先回忆本章一句话：${chapter.one_liner}`,
      "读图顺序建议：先看背景，再看量，再看价格结果，最后看后续是否延续。",
    ],
    { x: 7.95, y: 2.55, w: 4.2, h: 3.45 },
    { fontSize: 15, color: palette.text, paraSpaceAfterPt: 11 }
  );
  slide.addText(`图示来源：${figure.page ? `原书第 ${figure.page} 页` : "教学示意图"}`, {
    x: 7.95,
    y: 6.12,
    w: 3.6,
    h: 0.24,
    fontFace: FONT,
    fontSize: 9.5,
    color: hex(palette.muted),
  });
}

function addClosingSlide(slide, chapter) {
  addSectionTitle(slide, "本章最后带走的三句话", "把这一章压缩成最容易记住的三条。");
  chapter.memory_formula.forEach((item, index) => {
    addCard(
      slide,
      0.9 + index * 4.05,
      2.0,
      3.55,
      2.05,
      `第 ${index + 1} 句`,
      item,
      index === 1 ? "secondary" : index === 2 ? "accent" : "primary"
    );
  });
  slide.addShape(SHAPE_TYPE.roundRect, {
    x: 1.1,
    y: 4.55,
    w: 11.1,
    h: 1.45,
    rectRadius: 0.08,
    line: { color: hex(palette.primary), width: 1.3 },
    fill: { color: "FFFFFF" },
  });
  slide.addText(chapter.summary_capability, {
    x: 1.35,
    y: 4.9,
    w: 10.6,
    h: 0.7,
    fontFace: TITLE_FONT,
    fontSize: 22,
    bold: true,
    align: "center",
    color: hex(palette.primary),
    valign: "mid",
  });
}

async function buildSingle(chapterNumber) {
  const chapter = dataset.chapters.find((item) => item.number === chapterNumber);
  if (!chapter) {
    throw new Error(`未找到章节：${chapterNumber}`);
  }

  const chapterRoot = path.join(PPT_ROOT, chapter.folder);
  fs.mkdirSync(chapterRoot, { recursive: true });
  writeStubScript(chapter);

  const pptx = makePpt(chapter);
  const figureDeck = chapter.selected_figures.slice(0, 3);
  const totalSlides = 13;

  let slide = pptx.addSlide();
  baseSlide(slide, chapter, 1, totalSlides);
  slide.addShape(SHAPE_TYPE.roundRect, {
    x: 0.8,
    y: 1.2,
    w: 11.7,
    h: 5.2,
    rectRadius: 0.08,
    line: { color: hex(palette.primary), transparency: 100 },
    fill: { color: "FFFFFF" },
    shadow: safeOuterShadow("64748B", 0.1, 45, 1.4, 1),
  });
  slide.addText(`第${String(chapter.number).padStart(2, "0")}章`, {
    x: 1.15,
    y: 1.65,
    w: 1.5,
    h: 0.4,
    fontFace: FONT,
    fontSize: 15,
    bold: true,
    color: hex(palette.secondary),
  });
  slide.addText(chapter.title, {
    x: 1.12,
    y: 2.05,
    w: 7.35,
    h: 0.8,
    fontFace: TITLE_FONT,
    fontSize: 28,
    bold: true,
    color: hex(palette.primary),
  });
  slide.addText(chapter.one_liner, {
    x: 1.15,
    y: 3.0,
    w: 7.2,
    h: 0.65,
    fontFace: FONT,
    fontSize: 18,
    color: hex(palette.text),
  });
  slide.addText(`给零基础读者的导语：${chapter.kid_hook}`, {
    x: 1.15,
    y: 3.88,
    w: 6.9,
    h: 1.5,
    fontFace: FONT,
    fontSize: 13.5,
    color: hex(palette.muted),
    valign: "top",
    margin: 0.02,
  });
  addCard(slide, 9.2, 1.9, 2.35, 1.6, "本章先看什么", "先别急着记技巧，先把本章最想回答的问题装进脑子。", "accent");
  addCard(slide, 9.2, 3.75, 2.35, 1.6, "你会得到什么", chapter.summary_capability, "secondary");
  finalizeSlide(slide, pptx);

  slide = pptx.addSlide();
  baseSlide(slide, chapter, 2, totalSlides);
  addSectionTitle(slide, "本章回答什么问题", chapter.big_question);
  addCard(slide, 0.9, 1.9, 3.7, 1.75, "问题", chapter.big_question, "primary");
  addCard(slide, 4.85, 1.9, 3.7, 1.75, "一句话答案", chapter.one_liner, "secondary");
  addCard(slide, 8.8, 1.9, 3.7, 1.75, "为什么重要", chapter.chapter_report.core_thesis, "accent");
  slide.addText("本章路线图", { x: 0.95, y: 4.1, w: 2.1, h: 0.3, fontFace: TITLE_FONT, fontSize: 18, bold: true, color: hex(palette.primary) });
  addBulletList(slide, chapter.logic_chain, { x: 1.0, y: 4.45, w: 11.2, h: 1.95 }, { fontSize: 15.2 });
  finalizeSlide(slide, pptx);

  slide = pptx.addSlide();
  baseSlide(slide, chapter, 3, totalSlides);
  addSectionTitle(slide, "本章核心概念图", "先把本章最关键的四个词装成卡片。");
  chapter.concepts.slice(0, 4).forEach((item, index) => {
    const x = 0.9 + (index % 2) * 6.0;
    const y = 1.95 + Math.floor(index / 2) * 2.2;
    addCard(slide, x, y, 5.4, 1.8, item.term, item.simple, index % 2 === 0 ? "primary" : "secondary");
  });
  finalizeSlide(slide, pptx);

  slide = pptx.addSlide();
  baseSlide(slide, chapter, 4, totalSlides);
  addSectionTitle(slide, "术语翻译页", "把专业词换成孩子也能懂的话。");
  chapter.concepts.slice(0, 4).forEach((item, index) => {
    slide.addShape(SHAPE_TYPE.roundRect, {
      x: 0.85,
      y: 2.05 + index * 1.08,
      w: 2.4,
      h: 0.7,
      rectRadius: 0.06,
      line: { color: hex(palette.primary), width: 1.1 },
      fill: { color: index % 2 === 0 ? "EAF1F6" : "E8F4F3" },
    });
    slide.addText(item.term, {
      x: 1.0,
      y: 2.24 + index * 1.08,
      w: 2.0,
      h: 0.25,
      fontFace: TITLE_FONT,
      fontSize: 17,
      bold: true,
      color: hex(palette.primary),
      align: "center",
    });
    slide.addText(item.simple, {
      x: 3.55,
      y: 2.12 + index * 1.08,
      w: 8.15,
      h: 0.5,
      fontFace: FONT,
      fontSize: 15,
      color: hex(palette.text),
    });
  });
  finalizeSlide(slide, pptx);

  slide = pptx.addSlide();
  baseSlide(slide, chapter, 5, totalSlides);
  addSectionTitle(slide, "关键逻辑拆解页", "把本章最重要的推理顺序排成一条线。");
  chapter.logic_chain.forEach((item, index) => {
    const y = 1.9 + index * 1.22;
    slide.addShape(SHAPE_TYPE.roundRect, {
      x: 0.95,
      y,
      w: 1.05,
      h: 0.62,
      rectRadius: 0.06,
      line: { color: hex(palette.secondary), width: 1.2 },
      fill: { color: "E8F4F3" },
    });
    slide.addText(`步骤 ${index + 1}`, {
      x: 1.06,
      y: y + 0.17,
      w: 0.8,
      h: 0.2,
      fontFace: FONT,
      fontSize: 13,
      bold: true,
      align: "center",
      color: hex(palette.secondary),
    });
    slide.addShape(SHAPE_TYPE.line, {
      x: 2.15,
      y: y + 0.31,
      w: 0.55,
      h: 0,
      line: { color: hex(palette.accent), width: 2.2, beginArrowType: "none", endArrowType: "triangle" },
    });
    slide.addText(item, {
      x: 2.85,
      y: y + 0.03,
      w: 9.2,
      h: 0.5,
      fontFace: FONT,
      fontSize: 15.2,
      color: hex(palette.text),
    });
  });
  finalizeSlide(slide, pptx);

  const figureLabels = ["A", "B", "C"];
  figureDeck.forEach((figure, idx) => {
    slide = pptx.addSlide();
    baseSlide(slide, chapter, 6 + idx, totalSlides);
    addFigureSlide(slide, chapter, figure, chapterRoot, figureLabels[idx]);
    finalizeSlide(slide, pptx);
  });

  for (let fillerIndex = figureDeck.length; fillerIndex < 3; fillerIndex += 1) {
    slide = pptx.addSlide();
    baseSlide(slide, chapter, 6 + fillerIndex, totalSlides);
    addSectionTitle(slide, `${figureLabels[fillerIndex]} 图解延伸`, "当原书图不足时，用文字把图像背后的判断顺序补齐。");
    addCard(slide, 0.95, 2.0, 3.7, 1.8, "先看背景", chapter.kid_hook, "primary");
    addCard(slide, 4.85, 2.0, 3.7, 1.8, "再看量价", chapter.chapter_report.logic, "secondary");
    addCard(slide, 8.75, 2.0, 3.7, 1.8, "最后看收获", chapter.chapter_report.deep_gain, "accent");
    finalizeSlide(slide, pptx);
  }

  slide = pptx.addSlide();
  baseSlide(slide, chapter, 9, totalSlides);
  addSectionTitle(slide, "常见误解或反例页", "先知道自己最容易在哪里看错。");
  addBulletList(slide, chapter.misreads, { x: 0.95, y: 2.0, w: 5.85, h: 3.7 }, { fontSize: 17 });
  addCard(slide, 7.15, 2.1, 5.0, 2.1, "纠错提醒", "看到信号时不要只问“像不像”，而要问“它出现在哪里，量有没有配合，后面有没有继续”。", "accent");
  addCard(slide, 7.15, 4.45, 5.0, 1.35, "成熟读图习惯", "先看异常，再等确认；先看背景，再想操作。", "secondary");
  finalizeSlide(slide, pptx);

  slide = pptx.addSlide();
  baseSlide(slide, chapter, 10, totalSlides);
  addSectionTitle(slide, "实战观察步骤页", "把书里的概念变成一个能重复执行的动作流程。");
  chapter.practice_steps.forEach((item, index) => {
    addCard(slide, 0.9 + (index % 2) * 6.0, 1.95 + Math.floor(index / 2) * 2.1, 5.35, 1.65, `动作 ${index + 1}`, item, index % 2 === 0 ? "secondary" : "primary");
  });
  finalizeSlide(slide, pptx);

  slide = pptx.addSlide();
  baseSlide(slide, chapter, 11, totalSlides);
  addSectionTitle(slide, "记忆口诀页", "用短句把整章压缩成脑中的提示词。");
  chapter.memory_formula.forEach((item, index) => {
    addCard(slide, 1.1, 1.9 + index * 1.55, 11.0, 1.15, `口诀 ${index + 1}`, item, index === 1 ? "accent" : "primary");
  });
  finalizeSlide(slide, pptx);

  slide = pptx.addSlide();
  baseSlide(slide, chapter, 12, totalSlides);
  addSectionTitle(slide, "本章收获与复盘页", "看完这一章，你应该真正得到什么。");
  addBulletList(slide, chapter.takeaways, { x: 0.95, y: 1.95, w: 6.1, h: 3.7 }, { fontSize: 17 });
  addCard(slide, 7.35, 2.0, 4.8, 2.5, "你现在应具备的判断力", chapter.summary_capability, "secondary");
  addCard(slide, 7.35, 4.8, 4.8, 1.2, "与后续章节的衔接", chapter.chapter_report.bridge, "accent");
  finalizeSlide(slide, pptx);

  slide = pptx.addSlide();
  baseSlide(slide, chapter, 13, totalSlides);
  addClosingSlide(slide, chapter);
  finalizeSlide(slide, pptx);

  const outputFile = path.join(chapterRoot, `${chapter.folder}.pptx`);
  await pptx.writeFile({ fileName: outputFile });
  return outputFile;
}

async function buildAll() {
  const outputs = [];
  for (const chapter of dataset.chapters) {
    outputs.push(await buildSingle(chapter.number));
  }
  return outputs;
}

if (require.main === module) {
  const arg = process.argv[2];
  const task = arg ? buildSingle(Number(arg)) : buildAll();
  task.catch((error) => {
    console.error(error);
    process.exit(1);
  });
}

module.exports = {
  buildSingle,
  buildAll,
};
