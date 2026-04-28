"use strict";

const fs = require("fs");
const path = require("path");

function loadPptxGen(projectRoot) {
  const candidates = [
    path.join(__dirname, "node_modules", "pptxgenjs"),
    path.join(__dirname, "..", "书籍精读交付", "node_modules", "pptxgenjs"),
    path.join(projectRoot, "node_modules", "pptxgenjs"),
    path.join(projectRoot, "..", "node_modules", "pptxgenjs"),
    path.join(projectRoot, "..", "..", "node_modules", "pptxgenjs"),
  ];
  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return require(candidate);
    }
  }
  return require("pptxgenjs");
}

function loadHelpers(projectRoot) {
  const helperBase = path.join(projectRoot, "_build", "scripts", "pptxgenjs_helpers");
  return {
    imageSizingContain: require(path.join(helperBase, "image.js")).imageSizingContain,
    warnIfSlideHasOverlaps: require(path.join(helperBase, "layout.js")).warnIfSlideHasOverlaps,
    warnIfSlideElementsOutOfBounds: require(path.join(helperBase, "layout.js")).warnIfSlideElementsOutOfBounds,
  };
}

function hex(color) {
  return String(color || "").replace("#", "");
}

function safeSegment(value) {
  return String(value)
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/[：，、“”"'《》·]/g, "_")
    .replace(/\s+/g, "")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function sanitizeBulletText(value) {
  return String(value || "")
    .trim()
    .replace(/[。；，、：;,:]+$/u, "");
}

function baseDeck(PptxGenJS, title, subject) {
  const ppt = new PptxGenJS();
  ppt.layout = "LAYOUT_WIDE";
  ppt.author = "OpenAI Codex";
  ppt.company = "OpenAI";
  ppt.subject = subject;
  ppt.title = title;
  ppt.lang = "zh-CN";
  ppt.defineLayout({ name: "LAYOUT_WIDE", width: 13.333, height: 7.5 });
  ppt.theme = { headFontFace: "Microsoft YaHei", bodyFontFace: "Microsoft YaHei", lang: "zh-CN" };
  return ppt;
}

function addHeader(slide, title, tag, colors) {
  slide.addShape("rect", {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.54,
    fill: { color: hex(colors.primary) },
    line: { color: hex(colors.primary), transparency: 100 },
  });
  slide.addText(tag, {
    x: 0.6,
    y: 0.13,
    w: 3.4,
    h: 0.2,
    fontFace: "Microsoft YaHei",
    fontSize: 12,
    bold: true,
    color: "FFFFFF",
    margin: 0,
  });
  slide.addText(title, {
    x: 0.7,
    y: 0.78,
    w: 10.6,
    h: 0.42,
    fontFace: "Microsoft YaHei",
    fontSize: 24,
    bold: true,
    color: hex(colors.dark),
    margin: 0,
  });
}

function addFooter(slide, leftText, rightText, color) {
  slide.addText(leftText, {
    x: 0.55,
    y: 7.02,
    w: 8.0,
    h: 0.18,
    fontFace: "Microsoft YaHei",
    fontSize: 9,
    color: "64748B",
    margin: 0,
  });
  slide.addText(String(rightText), {
    x: 11.75,
    y: 7.0,
    w: 0.95,
    h: 0.18,
    fontFace: "Microsoft YaHei",
    fontSize: 10,
    bold: true,
    align: "right",
    color: hex(color),
    margin: 0,
  });
}

function finalizeSlide(slide, ppt, helpers, skipOverlap = false) {
  if (!skipOverlap) {
    helpers.warnIfSlideHasOverlaps(slide, ppt, { muteContainment: true });
  }
  helpers.warnIfSlideElementsOutOfBounds(slide, ppt);
}

function bulletRuns(items) {
  return items.map((item, index) => ({
    text: sanitizeBulletText(item),
    options: { bullet: true, breakLine: index !== items.length - 1 },
  }));
}

function addBullets(slide, items, x, y, w, h, fontSize, color) {
  slide.addText(bulletRuns(items), {
    x,
    y,
    w,
    h,
    fontFace: "Microsoft YaHei",
    fontSize,
    color,
    margin: 0,
    paraSpaceAfterPt: 8,
    valign: "top",
  });
}

function addConceptCard(slide, concept, x, y, w, h, colors) {
  slide.addShape("roundRect", {
    x, y, w, h,
    rectRadius: 0.05,
    fill: { color: hex(colors.bg) },
    line: { color: hex(colors.secondary), width: 1.4 },
  });
  slide.addText(concept.name, {
    x: x + 0.18,
    y: y + 0.16,
    w: w - 0.36,
    h: 0.24,
    fontFace: "Microsoft YaHei",
    fontSize: 20,
    bold: true,
    color: hex(colors.primary),
    margin: 0,
  });
  slide.addText(concept.explain, {
    x: x + 0.18,
    y: y + 0.54,
    w: w - 0.36,
    h: h - 0.7,
    fontFace: "Microsoft YaHei",
    fontSize: 16,
    color: "334155",
    margin: 0,
    breakLine: false,
  });
}

function buildChapterDeck(projectRoot, book, chapter, helpers, PptxGenJS) {
  const ppt = baseDeck(PptxGenJS, `${book.title} ${chapter.full_title}`, `${book.title} 零基础教学版`);
  const colors = chapter.palette;
  const reportImgDir = path.join(projectRoot, "40_配图", "阅读报告");
  const summaryImgDir = path.join(projectRoot, "40_配图", "关键知识总结");
  const pptImgDir = path.join(projectRoot, "40_配图", "PPT");
  const sourceImgDir = path.join(projectRoot, "40_配图", "原书页图");
  const poster = path.join(pptImgDir, `${chapter.id}_poster.png`);
  const map = path.join(reportImgDir, `${chapter.id}_core_map.png`);
  const card = path.join(summaryImgDir, `${chapter.id}_memory_card.png`);
  const opener = path.join(sourceImgDir, `${chapter.id}_opener_page.png`);
  const fileBase = `${chapter.sequence_no}_第${chapter.sequence_no}章_${safeSegment(chapter.title)}_零基础教学版`;

  let slide = ppt.addSlide();
  slide.background = { color: hex(colors.dark) };
  slide.addShape("roundRect", {
    x: 0.45, y: 0.4, w: 12.45, h: 6.7,
    rectRadius: 0.06,
    fill: { color: hex(colors.bg), transparency: 8 },
    line: { color: hex(colors.secondary), width: 2 },
  });
  slide.addText(book.title, {
    x: 0.82, y: 0.78, w: 4.0, h: 0.3,
    fontFace: "Microsoft YaHei", fontSize: 22, bold: true, color: hex(colors.dark), margin: 0,
  });
  slide.addText(chapter.full_title, {
    x: 0.82, y: 1.34, w: 5.5, h: 0.8,
    fontFace: "Microsoft YaHei", fontSize: 28, bold: true, color: hex(colors.dark), margin: 0,
  });
  slide.addText("零基础教学版", {
    x: 0.84, y: 2.3, w: 3.8, h: 0.24,
    fontFace: "Microsoft YaHei", fontSize: 16, color: hex(colors.primary), margin: 0,
  });
  slide.addText(chapter.reader_question, {
    x: 0.84, y: 2.9, w: 5.2, h: 1.8,
    fontFace: "Microsoft YaHei", fontSize: 19, color: "334155", margin: 0, breakLine: false,
  });
  if (fs.existsSync(poster)) {
    slide.addImage({ path: poster, ...helpers.imageSizingContain(poster, 6.2, 1.0, 6.0, 5.5) });
  }
  addFooter(slide, `${book.title} / 章节导入`, chapter.sequence_no, colors.secondary);
  finalizeSlide(slide, ppt, helpers, true);

  slide = ppt.addSlide();
  slide.background = { color: hex(colors.bg) };
  addHeader(slide, "这章到底在讲什么", chapter.full_title, colors);
  slide.addText(chapter.core_message, {
    x: 0.82, y: 1.48, w: 5.2, h: 1.25,
    fontFace: "Microsoft YaHei", fontSize: 20, bold: true, color: hex(colors.dark), margin: 0, breakLine: false,
  });
  slide.addText(chapter.position_in_book, {
    x: 0.82, y: 3.0, w: 5.1, h: 1.0,
    fontFace: "Microsoft YaHei", fontSize: 16, color: "475569", margin: 0, breakLine: false,
  });
  slide.addText(`为什么重要：${chapter.action_takeaways[0]}`, {
    x: 0.82, y: 4.35, w: 5.1, h: 0.86,
    fontFace: "Microsoft YaHei", fontSize: 17, bold: true, color: hex(colors.primary), margin: 0, breakLine: false,
  });
  if (fs.existsSync(map)) {
    slide.addImage({ path: map, ...helpers.imageSizingContain(map, 6.2, 1.35, 6.15, 5.45) });
  }
  addFooter(slide, `${book.title} / 章节总览`, chapter.sequence_no, colors.primary);
  finalizeSlide(slide, ppt, helpers);

  slide = ppt.addSlide();
  slide.background = { color: "FFFFFF" };
  addHeader(slide, "三个关键词", chapter.full_title, colors);
  addConceptCard(slide, chapter.key_concepts[0], 0.8, 1.55, 3.9, 4.8, colors);
  addConceptCard(slide, chapter.key_concepts[1], 4.85, 1.55, 3.9, 4.8, colors);
  addConceptCard(slide, chapter.key_concepts[2], 8.9, 1.55, 3.65, 4.8, colors);
  addFooter(slide, `${book.title} / 关键词`, chapter.sequence_no, colors.primary);
  finalizeSlide(slide, ppt, helpers);

  slide = ppt.addSlide();
  slide.background = { color: "FFFDF8" };
  addHeader(slide, "通俗比喻", chapter.full_title, colors);
  slide.addText(chapter.plain_example, {
    x: 0.85, y: 1.55, w: 5.35, h: 2.9,
    fontFace: "Microsoft YaHei", fontSize: 21, color: hex(colors.dark), margin: 0, breakLine: false,
  });
  addBullets(slide, chapter.must_remember.slice(0, 3), 0.95, 4.65, 5.3, 1.5, 17, "334155");
  if (fs.existsSync(opener)) {
    slide.addImage({ path: opener, ...helpers.imageSizingContain(opener, 6.45, 1.4, 5.85, 5.25) });
  }
  addFooter(slide, `${book.title} / 通俗理解`, chapter.sequence_no, colors.primary);
  finalizeSlide(slide, ppt, helpers);

  slide = ppt.addSlide();
  slide.background = { color: hex(colors.bg) };
  addHeader(slide, "作者要你看到的结构", chapter.full_title, colors);
  slide.addText(chapter.report_paragraphs[0], {
    x: 0.82, y: 1.55, w: 5.8, h: 1.5,
    fontFace: "Microsoft YaHei", fontSize: 18, color: "334155", margin: 0, breakLine: false,
  });
  slide.addText(chapter.report_paragraphs[1], {
    x: 0.82, y: 3.25, w: 5.8, h: 1.5,
    fontFace: "Microsoft YaHei", fontSize: 18, color: "334155", margin: 0, breakLine: false,
  });
  slide.addText(`一句话：${chapter.one_line_review}`, {
    x: 0.82, y: 5.2, w: 5.5, h: 0.9,
    fontFace: "Microsoft YaHei", fontSize: 17, bold: true, color: hex(colors.primary), margin: 0, breakLine: false,
  });
  if (fs.existsSync(map)) {
    slide.addImage({ path: map, ...helpers.imageSizingContain(map, 6.75, 1.45, 5.5, 4.9) });
  }
  addFooter(slide, `${book.title} / 结构理解`, chapter.sequence_no, colors.primary);
  finalizeSlide(slide, ppt, helpers);

  slide = ppt.addSlide();
  slide.background = { color: "FFFFFF" };
  addHeader(slide, "放到现实里看", chapter.full_title, colors);
  slide.addShape("roundRect", {
    x: 0.82, y: 1.58, w: 5.6, h: 4.95,
    rectRadius: 0.05,
    fill: { color: "F8FAFC" },
    line: { color: hex(colors.secondary), width: 1.2 },
  });
  slide.addText("两个现实例子", {
    x: 1.02, y: 1.82, w: 2.0, h: 0.25,
    fontFace: "Microsoft YaHei", fontSize: 19, bold: true, color: hex(colors.primary), margin: 0,
  });
  addBullets(slide, chapter.real_examples, 1.02, 2.28, 5.0, 2.2, 18, "334155");
  slide.addText("立刻能做什么", {
    x: 1.02, y: 4.85, w: 2.0, h: 0.25,
    fontFace: "Microsoft YaHei", fontSize: 19, bold: true, color: hex(colors.primary), margin: 0,
  });
  addBullets(slide, chapter.action_takeaways.slice(0, 2), 1.02, 5.25, 5.0, 1.1, 17, "334155");
  if (fs.existsSync(poster)) {
    slide.addImage({ path: poster, ...helpers.imageSizingContain(poster, 6.8, 1.45, 5.4, 5.1) });
  }
  addFooter(slide, `${book.title} / 现实应用`, chapter.sequence_no, colors.primary);
  finalizeSlide(slide, ppt, helpers);

  slide = ppt.addSlide();
  slide.background = { color: "FFF7ED" };
  addHeader(slide, "容易误解的地方", chapter.full_title, colors);
  addBullets(slide, chapter.misunderstandings, 0.9, 1.65, 5.9, 4.5, 18, "7C2D12");
  slide.addText("读这一章时，别把一句响亮的话当成整套方法。真正重要的是：它让你以后怎么安排自己的判断结构。", {
    x: 7.0, y: 1.7, w: 5.3, h: 2.0,
    fontFace: "Microsoft YaHei", fontSize: 18, color: "334155", margin: 0, breakLine: false,
  });
  if (fs.existsSync(card)) {
    slide.addImage({ path: card, ...helpers.imageSizingContain(card, 7.0, 3.7, 5.2, 2.7) });
  }
  addFooter(slide, `${book.title} / 误区提醒`, chapter.sequence_no, colors.primary);
  finalizeSlide(slide, ppt, helpers);

  slide = ppt.addSlide();
  slide.background = { color: "FFFFFF" };
  addHeader(slide, "本章真正学到了什么", chapter.full_title, colors);
  addBullets(slide, chapter.must_remember, 0.95, 1.55, 6.2, 4.9, 18, "334155");
  slide.addText("把这 5 条说清楚，你就不是“看过”这章，而是开始“用”这章了。", {
    x: 7.35, y: 1.8, w: 4.8, h: 1.2,
    fontFace: "Microsoft YaHei", fontSize: 18, color: hex(colors.primary), margin: 0, breakLine: false,
  });
  if (fs.existsSync(card)) {
    slide.addImage({ path: card, ...helpers.imageSizingContain(card, 7.2, 3.0, 5.15, 3.1) });
  }
  addFooter(slide, `${book.title} / 核心收获`, chapter.sequence_no, colors.primary);
  finalizeSlide(slide, ppt, helpers);

  slide = ppt.addSlide();
  slide.background = { color: hex(colors.dark) };
  slide.addText("复盘问题", {
    x: 0.92, y: 0.92, w: 2.2, h: 0.36,
    fontFace: "Microsoft YaHei", fontSize: 28, bold: true, color: "FFFFFF", margin: 0,
  });
  slide.addText(chapter.one_line_review, {
    x: 0.92, y: 1.55, w: 5.2, h: 1.0,
    fontFace: "Microsoft YaHei", fontSize: 22, bold: true, color: hex(colors.secondary), margin: 0, breakLine: false,
  });
  addBullets(slide, chapter.study_questions, 0.95, 2.7, 5.7, 3.6, 18, "F8FAFC");
  if (fs.existsSync(poster)) {
    slide.addImage({ path: poster, ...helpers.imageSizingContain(poster, 6.75, 1.4, 5.45, 4.9) });
  }
  addFooter(slide, `${book.title} / 复盘收束`, chapter.sequence_no, colors.secondary);
  finalizeSlide(slide, ppt, helpers, true);

  return { ppt, fileBase };
}

async function buildBook(projectRoot) {
  const content = JSON.parse(fs.readFileSync(path.join(projectRoot, "_build", "chapter_content.json"), "utf8"));
  const helpers = loadHelpers(projectRoot);
  const PptxGenJS = loadPptxGen(projectRoot);
  const outputDir = path.join(projectRoot, "10_PPT");
  fs.mkdirSync(outputDir, { recursive: true });
  for (const file of fs.readdirSync(outputDir)) {
    if (file.toLowerCase().endsWith(".pptx")) {
      fs.unlinkSync(path.join(outputDir, file));
    }
  }
  for (const chapter of content.chapters) {
    const { ppt, fileBase } = buildChapterDeck(projectRoot, content.book, chapter, helpers, PptxGenJS);
    const output = path.join(outputDir, `${fileBase}.pptx`);
    await ppt.writeFile({ fileName: output });
  }
}

module.exports = { buildBook };

if (require.main === module) {
  const projectRoot = process.argv[2];
  if (!projectRoot) {
    console.error("Usage: node scripts/teaching_pack_ppt_builder.js <projectRoot>");
    process.exit(1);
  }
  buildBook(path.resolve(projectRoot)).catch((error) => {
    console.error(error);
    process.exit(1);
  });
}
