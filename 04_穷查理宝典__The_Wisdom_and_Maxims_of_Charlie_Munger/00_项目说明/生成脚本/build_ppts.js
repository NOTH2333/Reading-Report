"use strict";

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const pptx = {
  ShapeType: {
    rect: "rect",
    roundRect: "roundRect",
    ellipse: "ellipse",
  },
};
const { imageSizingContain } = require("./pptxgenjs_helpers/image");
const {
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
} = require("./pptxgenjs_helpers/layout");

const ROOT = path.resolve(__dirname, "..", "..");
const CONTENT_PATH = path.join(ROOT, "05_中间素材", "chapter_content.json");
const PPT_DIR = path.join(ROOT, "03_PPT");
const PPT_IMG_DIR = path.join(ROOT, "04_配图", "PPT");
const REPORT_IMG_DIR = path.join(ROOT, "04_配图", "阅读报告");
const SUMMARY_IMG_DIR = path.join(ROOT, "04_配图", "关键知识总结");
const SOURCE_IMG_DIR = path.join(ROOT, "04_配图", "原书页图");

const content = JSON.parse(fs.readFileSync(CONTENT_PATH, "utf8"));

function themeFor(chapter) {
  return {
    headFontFace: "Microsoft YaHei",
    bodyFontFace: "Microsoft YaHei",
    lang: "zh-CN",
  };
}

function baseDeck(chapter) {
  const ppt = new pptxgen();
  ppt.layout = "LAYOUT_WIDE";
  ppt.author = "OpenAI Codex";
  ppt.company = "OpenAI";
  ppt.subject = "穷查理宝典儿童可读版";
  ppt.title = `${chapter.full_title} 儿童可读版`;
  ppt.theme = themeFor(chapter);
  ppt.lang = "zh-CN";
  ppt.defineLayout({ name: "LAYOUT_WIDE", width: 13.333, height: 7.5 });
  return ppt;
}

function chapterFileName(chapter) {
  return `第${chapter.number}章_${chapter.title}_儿童可读版`;
}

function hex(color) {
  return color.replace("#", "");
}

function addFooter(slide, chapter) {
  slide.addText("《穷查理宝典》儿童可读版学习资料包", {
    x: 0.5,
    y: 7.05,
    w: 6.2,
    h: 0.22,
    fontFace: "Microsoft YaHei",
    fontSize: 9,
    color: "666666",
    margin: 0,
  });
  slide.addText(`第${chapter.number}章`, {
    x: 12.0,
    y: 7.03,
    w: 0.8,
    h: 0.22,
    fontFace: "Microsoft YaHei",
    fontSize: 10,
    bold: true,
    align: "right",
    color: hex(chapter.palette.primary),
    margin: 0,
  });
}

function addContentHeader(slide, chapter, title, eyebrow = "Charlie Munger") {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.58,
    line: { color: hex(chapter.palette.primary), transparency: 100 },
    fill: { color: hex(chapter.palette.primary) },
  });
  slide.addText(eyebrow, {
    x: 0.65,
    y: 0.14,
    w: 2.8,
    h: 0.2,
    fontFace: "Microsoft YaHei",
    fontSize: 12,
    color: "FFFFFF",
    bold: true,
    margin: 0,
  });
  slide.addText(title, {
    x: 0.65,
    y: 0.82,
    w: 8.8,
    h: 0.45,
    fontFace: "Microsoft YaHei",
    fontSize: 26,
    bold: true,
    color: hex(chapter.palette.dark),
    margin: 0,
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 10.9,
    y: 0.83,
    w: 1.65,
    h: 0.36,
    rectRadius: 0.06,
    line: { color: hex(chapter.palette.secondary), transparency: 100 },
    fill: { color: hex(chapter.palette.bg) },
  });
  slide.addText(`第${chapter.number}章`, {
    x: 11.12,
    y: 0.92,
    w: 1.2,
    h: 0.18,
    fontFace: "Microsoft YaHei",
    fontSize: 14,
    bold: true,
    color: hex(chapter.palette.primary),
    margin: 0,
    align: "center",
  });
  addFooter(slide, chapter);
}

function finalizeSlide(ppt, slide, options = {}) {
  if (!options.skipOverlap) {
    warnIfSlideHasOverlaps(slide, ppt, { muteContainment: true });
  }
  warnIfSlideElementsOutOfBounds(slide, ppt);
}

function addCoverSlide(ppt, chapter) {
  const slide = ppt.addSlide();
  slide.background = { color: hex(chapter.palette.dark) };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.45,
    y: 0.42,
    w: 12.43,
    h: 6.65,
    line: { color: hex(chapter.palette.secondary), width: 2.2 },
    fill: { color: hex(chapter.palette.dark), transparency: 5 },
  });
  slide.addText("《穷查理宝典》", {
    x: 0.82,
    y: 0.88,
    w: 3.4,
    h: 0.35,
    fontFace: "Microsoft YaHei",
    fontSize: 20,
    color: "FFFFFF",
    bold: true,
    margin: 0,
  });
  slide.addText(chapter.full_title, {
    x: 0.82,
    y: 1.45,
    w: 5.2,
    h: 0.9,
    fontFace: "Microsoft YaHei",
    fontSize: 28,
    bold: true,
    color: "FFFFFF",
    margin: 0,
  });
  slide.addText("儿童可读版学习 PPT", {
    x: 0.84,
    y: 2.5,
    w: 4.2,
    h: 0.32,
    fontFace: "Microsoft YaHei",
    fontSize: 16,
    color: hex(chapter.palette.secondary),
    margin: 0,
  });
  slide.addText(chapter.core_message, {
    x: 0.84,
    y: 3.15,
    w: 4.8,
    h: 1.3,
    fontFace: "Microsoft YaHei",
    fontSize: 20,
    color: "F6F1EA",
    breakLine: false,
    margin: 0,
    valign: "mid",
  });
  const poster = path.join(PPT_IMG_DIR, `${chapter.id}_poster.png`);
  if (fs.existsSync(poster)) {
    slide.addImage({
      path: poster,
      ...imageSizingContain(poster, 6.15, 0.92, 6.1, 5.65),
    });
  }
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.84,
    y: 5.6,
    w: 4.4,
    h: 0.92,
    rectRadius: 0.08,
    line: { color: hex(chapter.palette.primary), transparency: 100 },
    fill: { color: hex(chapter.palette.primary) },
  });
  slide.addText(`核心问题：${chapter.children_question}`, {
    x: 1.08,
    y: 5.84,
    w: 3.95,
    h: 0.42,
    fontFace: "Microsoft YaHei",
    fontSize: 15,
    color: "FFFFFF",
    bold: true,
    margin: 0,
  });
  addFooter(slide, chapter);
  finalizeSlide(ppt, slide, { skipOverlap: true });
}

function addOneSentenceSlide(ppt, chapter) {
  const slide = ppt.addSlide();
  slide.background = { color: hex(chapter.palette.bg) };
  addContentHeader(slide, chapter, "这一章到底在讲什么");
  slide.addText(chapter.core_message, {
    x: 0.75,
    y: 1.55,
    w: 5.35,
    h: 1.3,
    fontFace: "Microsoft YaHei",
    fontSize: 22,
    bold: true,
    color: hex(chapter.palette.dark),
    margin: 0,
    valign: "mid",
  });
  slide.addText(`这章在全书中的位置：${chapter.position_in_book}`, {
    x: 0.78,
    y: 3.0,
    w: 5.0,
    h: 1.0,
    fontFace: "Microsoft YaHei",
    fontSize: 16,
    color: "4A5460",
    margin: 0,
    breakLine: false,
  });
  slide.addText(`孩子视角的问题：${chapter.children_question}`, {
    x: 0.78,
    y: 4.35,
    w: 5.0,
    h: 0.9,
    fontFace: "Microsoft YaHei",
    fontSize: 17,
    color: hex(chapter.palette.primary),
    bold: true,
    margin: 0,
  });
  const img = path.join(REPORT_IMG_DIR, `${chapter.id}_core_map.png`);
  if (fs.existsSync(img)) {
    slide.addImage({
      path: img,
      ...imageSizingContain(img, 6.42, 1.48, 6.05, 5.2),
    });
  }
  finalizeSlide(ppt, slide);
}

function addBackgroundSlide(ppt, chapter) {
  const slide = ppt.addSlide();
  slide.background = { color: "FFFFFF" };
  addContentHeader(slide, chapter, "背景故事与阅读入口");
  const opener = path.join(SOURCE_IMG_DIR, `${chapter.id}_opener_page.png`);
  if (fs.existsSync(opener)) {
    slide.addImage({
      path: opener,
      ...imageSizingContain(opener, 0.78, 1.45, 3.4, 4.95),
    });
  }
  slide.addText("背景故事", {
    x: 4.55,
    y: 1.5,
    w: 2.0,
    h: 0.28,
    fontFace: "Microsoft YaHei",
    fontSize: 18,
    bold: true,
    color: hex(chapter.palette.primary),
    margin: 0,
  });
  slide.addText(chapter.background_story, {
    x: 4.55,
    y: 1.86,
    w: 4.1,
    h: 2.25,
    fontFace: "Microsoft YaHei",
    fontSize: 16,
    color: "44515E",
    margin: 0,
    breakLine: false,
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 8.95,
    y: 1.52,
    w: 3.6,
    h: 4.95,
    rectRadius: 0.06,
    line: { color: hex(chapter.palette.secondary), width: 1.5 },
    fill: { color: hex(chapter.palette.bg) },
  });
  slide.addText("读这章最该看见的收获", {
    x: 9.18,
    y: 1.78,
    w: 2.8,
    h: 0.28,
    fontFace: "Microsoft YaHei",
    fontSize: 17,
    bold: true,
    color: hex(chapter.palette.dark),
    margin: 0,
  });
  let y = 2.22;
  chapter.chapter_gains.forEach((item) => {
    slide.addText("•", {
      x: 9.2,
      y,
      w: 0.18,
      h: 0.2,
      fontFace: "Microsoft YaHei",
      fontSize: 16,
      bold: true,
      color: hex(chapter.palette.primary),
      margin: 0,
    });
    slide.addText(item, {
      x: 9.45,
      y: y - 0.02,
      w: 2.76,
      h: 0.58,
      fontFace: "Microsoft YaHei",
      fontSize: 15,
      color: "42505C",
      margin: 0,
    });
    y += 1.05;
  });
  finalizeSlide(ppt, slide);
}

function addConceptSlide(ppt, chapter, title, items) {
  const slide = ppt.addSlide();
  slide.background = { color: hex(chapter.palette.bg) };
  addContentHeader(slide, chapter, title);
  const positions = [
    [0.8, 1.65],
    [6.9, 1.65],
    [0.8, 4.4],
  ];
  items.forEach((item, index) => {
    const [x, y] = positions[index];
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w: 5.55,
      h: 2.15,
      rectRadius: 0.06,
      line: { color: hex(chapter.palette.secondary), width: 1.4 },
      fill: { color: "FFFFFF" },
    });
    slide.addText(item.name, {
      x: x + 0.22,
      y: y + 0.18,
      w: 3.9,
      h: 0.28,
      fontFace: "Microsoft YaHei",
      fontSize: 18,
      bold: true,
      color: hex(chapter.palette.primary),
      margin: 0,
    });
    slide.addText(item.explain, {
      x: x + 0.22,
      y: y + 0.62,
      w: 4.5,
      h: 1.05,
      fontFace: "Microsoft YaHei",
      fontSize: 15,
      color: "46515E",
      margin: 0,
    });
    slide.addText(String(index + 1), {
      x: x + 4.7,
      y: y + 0.24,
      w: 0.32,
      h: 0.2,
      fontFace: "Microsoft YaHei",
      fontSize: 13,
      bold: true,
      color: hex(chapter.palette.primary),
      margin: 0,
      align: "center",
    });
  });
  const memory = path.join(SUMMARY_IMG_DIR, `${chapter.id}_memory_card.png`);
  if (fs.existsSync(memory)) {
    slide.addImage({
      path: memory,
      ...imageSizingContain(memory, 6.95, 4.35, 5.55, 2.2),
    });
  }
  finalizeSlide(ppt, slide);
}

function addStoriesSlide(ppt, chapter) {
  const slide = ppt.addSlide();
  slide.background = { color: "FFFFFF" };
  addContentHeader(slide, chapter, "代表性故事与案例");
  let x = 0.8;
  chapter.stories.forEach((story) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y: 1.7,
      w: 3.9,
      h: 4.9,
      rectRadius: 0.06,
      line: { color: hex(chapter.palette.secondary), width: 1.6 },
      fill: { color: hex(chapter.palette.bg) },
    });
    slide.addText(story.title, {
      x: x + 0.22,
      y: 1.98,
      w: 3.3,
      h: 0.48,
      fontFace: "Microsoft YaHei",
      fontSize: 18,
      bold: true,
      color: hex(chapter.palette.dark),
      margin: 0,
    });
    slide.addText(story.summary, {
      x: x + 0.22,
      y: 2.75,
      w: 3.35,
      h: 1.55,
      fontFace: "Microsoft YaHei",
      fontSize: 14,
      color: "46515E",
      margin: 0,
    });
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x + 0.22,
      y: 4.72,
      w: 3.38,
      h: 1.12,
      rectRadius: 0.05,
      line: { color: hex(chapter.palette.primary), transparency: 100 },
      fill: { color: hex(chapter.palette.primary) },
    });
    slide.addText(`启示：${story.lesson}`, {
      x: x + 0.38,
      y: 4.97,
      w: 3.05,
      h: 0.64,
      fontFace: "Microsoft YaHei",
      fontSize: 13,
      bold: true,
      color: "FFFFFF",
      margin: 0,
    });
    x += 4.15;
  });
  finalizeSlide(ppt, slide);
}

function addMistakeActionSlide(ppt, chapter) {
  const slide = ppt.addSlide();
  slide.background = { color: hex(chapter.palette.bg) };
  addContentHeader(slide, chapter, "常见误解与孩子能立刻做的事");
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.82,
    y: 1.62,
    w: 5.85,
    h: 4.95,
    rectRadius: 0.06,
    line: { color: hex(chapter.palette.secondary), width: 1.5 },
    fill: { color: "FFFFFF" },
  });
  slide.addText("容易误读的地方", {
    x: 1.05,
    y: 1.9,
    w: 2.2,
    h: 0.28,
    fontFace: "Microsoft YaHei",
    fontSize: 18,
    bold: true,
    color: hex(chapter.palette.primary),
    margin: 0,
  });
  let leftY = 2.35;
  chapter.misunderstandings.forEach((item) => {
    slide.addText("•", {
      x: 1.08,
      y: leftY,
      w: 0.15,
      h: 0.18,
      fontFace: "Microsoft YaHei",
      fontSize: 15,
      bold: true,
      color: hex(chapter.palette.primary),
      margin: 0,
    });
    slide.addText(item, {
      x: 1.34,
      y: leftY - 0.02,
      w: 4.95,
      h: 0.72,
      fontFace: "Microsoft YaHei",
      fontSize: 14,
      color: "46515E",
      margin: 0,
    });
    leftY += 1.12;
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.95,
    y: 1.62,
    w: 5.58,
    h: 4.95,
    rectRadius: 0.06,
    line: { color: hex(chapter.palette.secondary), width: 1.5 },
    fill: { color: "FFFFFF" },
  });
  slide.addText("孩子能立刻练的动作", {
    x: 7.18,
    y: 1.9,
    w: 2.6,
    h: 0.28,
    fontFace: "Microsoft YaHei",
    fontSize: 18,
    bold: true,
    color: hex(chapter.palette.primary),
    margin: 0,
  });
  let rightY = 2.35;
  chapter.child_actions.forEach((item, idx) => {
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 7.2,
      y: rightY - 0.02,
      w: 0.36,
      h: 0.36,
      line: { color: hex(chapter.palette.primary), transparency: 100 },
      fill: { color: hex(chapter.palette.primary) },
    });
    slide.addText(String(idx + 1), {
      x: 7.31,
      y: rightY + 0.08,
      w: 0.12,
      h: 0.1,
      fontFace: "Microsoft YaHei",
      fontSize: 9,
      bold: true,
      color: "FFFFFF",
      margin: 0,
      align: "center",
    });
    slide.addText(item, {
      x: 7.72,
      y: rightY - 0.04,
      w: 4.45,
      h: 0.82,
      fontFace: "Microsoft YaHei",
      fontSize: 14,
      color: "46515E",
      margin: 0,
    });
    rightY += 1.18;
  });
  finalizeSlide(ppt, slide);
}

function addRememberSlide(ppt, chapter) {
  const slide = ppt.addSlide();
  slide.background = { color: hex(chapter.palette.bg) };
  addContentHeader(slide, chapter, "这一章必须记住的 5 个点");
  let y = 1.62;
  chapter.must_remember.forEach((item, index) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 1.0,
      y,
      w: 11.2,
      h: 0.84,
      rectRadius: 0.05,
      line: { color: hex(chapter.palette.secondary), width: 1.2 },
      fill: { color: index % 2 === 0 ? "FFFFFF" : hex(chapter.palette.bg) },
    });
    slide.addText(String(index + 1), {
      x: 1.24,
      y: y + 0.18,
      w: 0.35,
      h: 0.22,
      fontFace: "Microsoft YaHei",
      fontSize: 18,
      bold: true,
      color: hex(chapter.palette.primary),
      margin: 0,
      align: "center",
    });
    slide.addText(item, {
      x: 1.72,
      y: y + 0.18,
      w: 9.95,
      h: 0.25,
      fontFace: "Microsoft YaHei",
      fontSize: 17,
      color: hex(chapter.palette.dark),
      bold: true,
      margin: 0,
    });
    y += 0.94;
  });
  finalizeSlide(ppt, slide);
}

function addGainsSlide(ppt, chapter) {
  const slide = ppt.addSlide();
  slide.background = { color: "FFFFFF" };
  addContentHeader(slide, chapter, "读完这一章，你应该带走什么");
  const colors = [chapter.palette.primary, chapter.palette.accent, chapter.palette.dark];
  let x = 0.82;
  chapter.chapter_gains.forEach((gain, index) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y: 2.0,
      w: 3.95,
      h: 3.8,
      rectRadius: 0.06,
      line: { color: hex(colors[index]), width: 1.4 },
      fill: { color: "FBFBFB" },
    });
    slide.addText(String(index + 1), {
      x: x + 0.22,
      y: 2.28,
      w: 0.6,
      h: 0.42,
      fontFace: "Microsoft YaHei",
      fontSize: 30,
      bold: true,
      color: hex(colors[index]),
      margin: 0,
    });
    slide.addText(gain, {
      x: x + 0.3,
      y: 3.0,
      w: 3.2,
      h: 1.85,
      fontFace: "Microsoft YaHei",
      fontSize: 18,
      bold: true,
      color: "394451",
      margin: 0,
      valign: "mid",
      align: "left",
    });
    x += 4.18;
  });
  slide.addText(`一句话复盘：${chapter.one_line_review}`, {
    x: 0.9,
    y: 6.2,
    w: 11.9,
    h: 0.42,
    fontFace: "Microsoft YaHei",
    fontSize: 18,
    bold: true,
    color: hex(chapter.palette.primary),
    margin: 0,
    align: "center",
  });
  finalizeSlide(ppt, slide);
}

function addRecapSlide(ppt, chapter) {
  const slide = ppt.addSlide();
  slide.background = { color: hex(chapter.palette.dark) };
  slide.addText(`第${chapter.number}章总复盘`, {
    x: 0.82,
    y: 0.72,
    w: 3.2,
    h: 0.35,
    fontFace: "Microsoft YaHei",
    fontSize: 22,
    bold: true,
    color: "FFFFFF",
    margin: 0,
  });
  slide.addText(chapter.one_line_review, {
    x: 0.82,
    y: 1.38,
    w: 4.8,
    h: 1.1,
    fontFace: "Microsoft YaHei",
    fontSize: 24,
    bold: true,
    color: hex(chapter.palette.secondary),
    margin: 0,
    valign: "mid",
  });
  slide.addText(`关键词：${chapter.must_remember.join(" / ")}`, {
    x: 0.85,
    y: 3.0,
    w: 4.7,
    h: 2.15,
    fontFace: "Microsoft YaHei",
    fontSize: 16,
    color: "F1F3F5",
    margin: 0,
  });
  const img = path.join(SUMMARY_IMG_DIR, `${chapter.id}_memory_card.png`);
  if (fs.existsSync(img)) {
    slide.addImage({
      path: img,
      ...imageSizingContain(img, 6.1, 1.0, 6.2, 5.45),
    });
  }
  addFooter(slide, chapter);
  finalizeSlide(ppt, slide);
}

function addLectureMapSlide(ppt, chapter) {
  const slide = ppt.addSlide();
  slide.background = { color: "FFFFFF" };
  addContentHeader(slide, chapter, "第四章总览：先看十一讲的地图");
  const img = path.join(PPT_IMG_DIR, "ch04_lecture_map.png");
  if (fs.existsSync(img)) {
    slide.addImage({
      path: img,
      ...imageSizingContain(img, 0.75, 1.42, 11.85, 5.35),
    });
  }
  finalizeSlide(ppt, slide);
}

function addLectureSlide(ppt, chapter, lecture, index) {
  const slide = ppt.addSlide();
  slide.background = { color: hex(chapter.palette.bg) };
  addContentHeader(slide, chapter, `第${index + 1}讲`, "Eleven Talks");
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.82,
    y: 1.48,
    w: 4.0,
    h: 5.3,
    rectRadius: 0.06,
    line: { color: hex(chapter.palette.primary), width: 1.8 },
    fill: { color: hex(chapter.palette.primary) },
  });
  slide.addText(lecture.title, {
    x: 1.08,
    y: 1.82,
    w: 3.45,
    h: 1.5,
    fontFace: "Microsoft YaHei",
    fontSize: 24,
    bold: true,
    color: "FFFFFF",
    margin: 0,
    valign: "mid",
  });
  slide.addText("最重要的意思", {
    x: 5.2,
    y: 1.72,
    w: 2.1,
    h: 0.25,
    fontFace: "Microsoft YaHei",
    fontSize: 18,
    bold: true,
    color: hex(chapter.palette.primary),
    margin: 0,
  });
  slide.addText(lecture.core, {
    x: 5.2,
    y: 2.1,
    w: 6.95,
    h: 1.1,
    fontFace: "Microsoft YaHei",
    fontSize: 20,
    bold: true,
    color: hex(chapter.palette.dark),
    margin: 0,
    valign: "mid",
  });
  slide.addText("孩子版翻译", {
    x: 5.2,
    y: 3.58,
    w: 1.8,
    h: 0.25,
    fontFace: "Microsoft YaHei",
    fontSize: 18,
    bold: true,
    color: hex(chapter.palette.primary),
    margin: 0,
  });
  slide.addText(`你可以把它理解成：${lecture.core}`, {
    x: 5.2,
    y: 3.95,
    w: 6.95,
    h: 0.95,
    fontFace: "Microsoft YaHei",
    fontSize: 16,
    color: "44505C",
    margin: 0,
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 5.2,
    y: 5.15,
    w: 6.9,
    h: 1.15,
    rectRadius: 0.05,
    line: { color: hex(chapter.palette.secondary), transparency: 100 },
    fill: { color: "FFFFFF" },
  });
  slide.addText(`和整本书的关系：${chapter.must_remember[index % chapter.must_remember.length]}`, {
    x: 5.45,
    y: 5.48,
    w: 6.45,
    h: 0.5,
    fontFace: "Microsoft YaHei",
    fontSize: 15,
    color: "475360",
    margin: 0,
  });
  finalizeSlide(ppt, slide);
}

function addArticleSlide(ppt, chapter) {
  const slide = ppt.addSlide();
  slide.background = { color: hex(chapter.palette.bg) };
  addContentHeader(slide, chapter, "本章关注的现实议题");
  let y = 1.65;
  chapter.article_examples.forEach((item, index) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 1.0,
      y,
      w: 11.25,
      h: 0.8,
      rectRadius: 0.04,
      line: { color: hex(chapter.palette.secondary), transparency: 100 },
      fill: { color: index % 2 === 0 ? "FFFFFF" : hex(chapter.palette.bg) },
    });
    slide.addText(`${index + 1}. ${item}`, {
      x: 1.25,
      y: y + 0.22,
      w: 9.6,
      h: 0.28,
      fontFace: "Microsoft YaHei",
      fontSize: 18,
      bold: true,
      color: hex(chapter.palette.dark),
      margin: 0,
    });
    y += 0.93;
  });
  finalizeSlide(ppt, slide);
}

function buildRegularDeck(chapter) {
  const ppt = baseDeck(chapter);
  addCoverSlide(ppt, chapter);
  addOneSentenceSlide(ppt, chapter);
  addBackgroundSlide(ppt, chapter);
  addConceptSlide(ppt, chapter, "关键概念 1/2", chapter.key_concepts.slice(0, 3));
  addConceptSlide(ppt, chapter, "关键概念 2/2", chapter.key_concepts.slice(3, 6));
  addStoriesSlide(ppt, chapter);
  addRememberSlide(ppt, chapter);
  if (chapter.id === "ch05") {
    addArticleSlide(ppt, chapter);
  }
  addMistakeActionSlide(ppt, chapter);
  addGainsSlide(ppt, chapter);
  addRecapSlide(ppt, chapter);
  return ppt;
}

function buildChapter4Deck(chapter) {
  const ppt = baseDeck(chapter);
  addCoverSlide(ppt, chapter);
  addOneSentenceSlide(ppt, chapter);
  addBackgroundSlide(ppt, chapter);
  addLectureMapSlide(ppt, chapter);
  chapter.lecture_index.forEach((lecture, index) => addLectureSlide(ppt, chapter, lecture, index));
  addConceptSlide(ppt, chapter, "第四章反复出现的五个抓手", chapter.key_concepts.slice(0, 3));
  addConceptSlide(ppt, chapter, "第四章反复出现的五个抓手（续）", chapter.key_concepts.slice(3, 5));
  addMistakeActionSlide(ppt, chapter);
  addRecapSlide(ppt, chapter);
  return ppt;
}

function buildDeck(chapterId) {
  const chapter = content.chapters.find((item) => item.id === chapterId);
  if (!chapter) {
    throw new Error(`Unknown chapter id: ${chapterId}`);
  }
  const ppt = chapterId === "ch04" ? buildChapter4Deck(chapter) : buildRegularDeck(chapter);
  return { ppt, chapter };
}

function chapterWrapperSource(chapter) {
  return [
    "\"use strict\";",
    "",
    "const path = require(\"path\");",
    "const { buildDeck, writeDeck } = require(path.resolve(__dirname, \"..\", \"00_项目说明\", \"生成脚本\", \"build_ppts.js\"));",
    "",
    `writeDeck("${chapter.id}").catch((error) => {`,
    "  console.error(error);",
    "  process.exitCode = 1;",
    "});",
    "",
  ].join("\n");
}

async function writeDeck(chapterId) {
  const { ppt, chapter } = buildDeck(chapterId);
  const outName = `${chapterFileName(chapter)}.pptx`;
  const outPath = path.join(PPT_DIR, outName);
  await ppt.writeFile({ fileName: outPath });
  const wrapperPath = path.join(PPT_DIR, `${chapterFileName(chapter)}.js`);
  fs.writeFileSync(wrapperPath, chapterWrapperSource(chapter), "utf8");
  return outPath;
}

async function buildAll() {
  for (const chapter of content.chapters) {
    await writeDeck(chapter.id);
  }
}

if (require.main === module) {
  buildAll().catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
}

module.exports = {
  buildDeck,
  writeDeck,
  buildAll,
};
