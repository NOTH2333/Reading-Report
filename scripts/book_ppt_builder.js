"use strict";

const fs = require("fs");
const path = require("path");
function loadPptxGen(projectRoot) {
  const localPath = path.join(projectRoot, "node_modules", "pptxgenjs");
  if (fs.existsSync(localPath)) {
    return require(localPath);
  }
  return require("pptxgenjs");
}

function loadHelpers(projectRoot) {
  const helperBase = path.join(projectRoot, "00_项目说明", "生成脚本", "pptxgenjs_helpers");
  return {
    imageSizingContain: require(path.join(helperBase, "image.js")).imageSizingContain,
    warnIfSlideHasOverlaps: require(path.join(helperBase, "layout.js")).warnIfSlideHasOverlaps,
    warnIfSlideElementsOutOfBounds: require(path.join(helperBase, "layout.js")).warnIfSlideElementsOutOfBounds,
  };
}

function hex(color) {
  return color.replace("#", "");
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
    w: 2.8,
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
    w: 10.0,
    h: 0.42,
    fontFace: "Microsoft YaHei",
    fontSize: 25,
    bold: true,
    color: hex(colors.dark),
    margin: 0,
  });
}

function addFooter(slide, leftText, rightText, color) {
  slide.addText(leftText, {
    x: 0.55,
    y: 7.02,
    w: 6.7,
    h: 0.18,
    fontFace: "Microsoft YaHei",
    fontSize: 9,
    color: "64748B",
    margin: 0,
  });
  slide.addText(rightText, {
    x: 11.8,
    y: 7.0,
    w: 0.9,
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

function addBullets(slide, items, x, y, w, fontSize, color) {
  const text = items.map((item) => ({ text: item, options: { bullet: { indent: 12 } } }));
  slide.addText(text, {
    x,
    y,
    w,
    h: 2.4,
    fontFace: "Microsoft YaHei",
    fontSize,
    color,
    margin: 0,
    breakLine: false,
    paraSpaceAfterPt: 8,
    valign: "top",
  });
}

function buildChapterDeck(projectRoot, content, chapter, helpers, PptxGenJS) {
  const ppt = baseDeck(PptxGenJS, `${content.book.title} ${chapter.full_title}`, `${content.book.title} 儿童可读版`);
  const colors = chapter.palette;
  const reportImg = path.join(projectRoot, "04_配图", "阅读报告");
  const summaryImg = path.join(projectRoot, "04_配图", "关键知识总结");
  const pptImg = path.join(projectRoot, "04_配图", "PPT");
  const sourceImg = path.join(projectRoot, "04_配图", "原书页图");
  const poster = path.join(pptImg, `${chapter.id}_poster.png`);
  const map = path.join(reportImg, `${chapter.id}_core_map.png`);
  const card = path.join(summaryImg, `${chapter.id}_memory_card.png`);
  const opener = path.join(sourceImg, `${chapter.id}_opener_page.png`);
  const fileBase = `第${chapter.number}章_${chapter.title}_儿童可读版`;

  let slide = ppt.addSlide();
  slide.background = { color: hex(colors.dark) };
  slide.addShape("roundRect", {
    x: 0.45, y: 0.4, w: 12.45, h: 6.7,
    rectRadius: 0.06,
    fill: { color: hex(colors.bg), transparency: 8 },
    line: { color: hex(colors.secondary), width: 2 },
  });
  slide.addText(content.book.title, { x: 0.8, y: 0.78, w: 3.8, h: 0.3, fontFace: "Microsoft YaHei", fontSize: 22, bold: true, color: "FFFFFF", margin: 0 });
  slide.addText(chapter.full_title, { x: 0.82, y: 1.35, w: 5.3, h: 0.8, fontFace: "Microsoft YaHei", fontSize: 28, bold: true, color: "FFFFFF", margin: 0 });
  slide.addText("儿童可读版学习 PPT", { x: 0.84, y: 2.34, w: 4.2, h: 0.24, fontFace: "Microsoft YaHei", fontSize: 16, color: hex(colors.secondary), margin: 0 });
  slide.addText(chapter.core_message, { x: 0.84, y: 3.0, w: 4.9, h: 1.4, fontFace: "Microsoft YaHei", fontSize: 20, color: "F8FAFC", margin: 0, breakLine: false });
  if (fs.existsSync(poster)) {
    slide.addImage({ path: poster, ...helpers.imageSizingContain(poster, 6.0, 1.0, 6.3, 5.5) });
  }
  addFooter(slide, `${content.book.title} 儿童可读版学习资料`, `第${chapter.number}章`, colors.primary);
  finalizeSlide(slide, ppt, helpers, true);

  slide = ppt.addSlide();
  slide.background = { color: hex(colors.bg) };
  addHeader(slide, "先知道", chapter.full_title, colors);
  slide.addText(chapter.core_message, { x: 0.75, y: 1.5, w: 5.4, h: 1.1, fontFace: "Microsoft YaHei", fontSize: 21, bold: true, color: hex(colors.dark), margin: 0, breakLine: false });
  slide.addText(`这章在书里的位置：${chapter.position_in_book}`, { x: 0.75, y: 2.9, w: 5.0, h: 1.0, fontFace: "Microsoft YaHei", fontSize: 16, color: "475569", margin: 0, breakLine: false });
  slide.addText(`孩子问题：${chapter.children_question}`, { x: 0.75, y: 4.3, w: 5.0, h: 0.8, fontFace: "Microsoft YaHei", fontSize: 17, bold: true, color: hex(colors.primary), margin: 0 });
  if (fs.existsSync(map)) {
    slide.addImage({ path: map, ...helpers.imageSizingContain(map, 6.25, 1.35, 6.2, 5.45) });
  }
  addFooter(slide, `${content.book.title} / 先知道`, chapter.number, colors.primary);
  finalizeSlide(slide, ppt, helpers);

  chapter.key_concepts.slice(0, 3).forEach((item, index) => {
    slide = ppt.addSlide();
    slide.background = { color: "FFFFFF" };
    addHeader(slide, `本章核心 ${index + 1}`, chapter.full_title, colors);
    slide.addShape("roundRect", {
      x: 0.78, y: 1.55, w: 4.1, h: 4.9,
      rectRadius: 0.06,
      fill: { color: hex(colors.bg) },
      line: { color: hex(colors.secondary), width: 1.8 },
    });
    slide.addText(item.name, { x: 1.02, y: 1.9, w: 3.6, h: 0.35, fontFace: "Microsoft YaHei", fontSize: 24, bold: true, color: hex(colors.primary), margin: 0 });
    slide.addText(item.explain, { x: 1.02, y: 2.42, w: 3.5, h: 2.6, fontFace: "Microsoft YaHei", fontSize: 18, color: "334155", margin: 0, breakLine: false });
    addBullets(slide, chapter.must_remember.slice(index, index + 2), 5.35, 1.72, 7.0, 18, "334155");
    addFooter(slide, `${content.book.title} / 核心 ${index + 1}`, chapter.number, colors.primary);
    finalizeSlide(slide, ppt, helpers);
  });

  slide = ppt.addSlide();
  slide.background = { color: hex(colors.bg) };
  addHeader(slide, "生活类比", chapter.full_title, colors);
  slide.addText(chapter.child_example, { x: 0.82, y: 1.55, w: 5.6, h: 2.8, fontFace: "Microsoft YaHei", fontSize: 22, color: hex(colors.dark), margin: 0, breakLine: false });
  if (chapter.stories && chapter.stories[0]) {
    slide.addText(`代表性故事：${chapter.stories[0].title}`, { x: 0.84, y: 4.55, w: 4.8, h: 0.25, fontFace: "Microsoft YaHei", fontSize: 17, bold: true, color: hex(colors.primary), margin: 0 });
    slide.addText(`${chapter.stories[0].summary} 启示：${chapter.stories[0].lesson}`, { x: 0.84, y: 4.92, w: 5.4, h: 1.2, fontFace: "Microsoft YaHei", fontSize: 16, color: "334155", margin: 0, breakLine: false });
  }
  if (fs.existsSync(opener)) {
    slide.addImage({ path: opener, ...helpers.imageSizingContain(opener, 6.55, 1.45, 5.8, 5.3) });
  }
  addFooter(slide, `${content.book.title} / 生活类比`, chapter.number, colors.primary);
  finalizeSlide(slide, ppt, helpers);

  slide = ppt.addSlide();
  slide.background = { color: "FFFFFF" };
  addHeader(slide, "投资里的真正意思", chapter.full_title, colors);
  slide.addText(chapter.investment_meaning, { x: 0.82, y: 1.55, w: 7.0, h: 2.4, fontFace: "Microsoft YaHei", fontSize: 20, color: "334155", margin: 0, breakLine: false });
  addBullets(slide, chapter.chapter_gains, 0.9, 4.35, 6.6, 18, "334155");
  if (fs.existsSync(map)) {
    slide.addImage({ path: map, ...helpers.imageSizingContain(map, 7.9, 1.55, 4.6, 4.6) });
  }
  addFooter(slide, `${content.book.title} / 真正意思`, chapter.number, colors.primary);
  finalizeSlide(slide, ppt, helpers);

  slide = ppt.addSlide();
  slide.background = { color: "FFF7ED" };
  addHeader(slide, "常见误区", chapter.full_title, colors);
  addBullets(slide, chapter.misunderstandings, 0.9, 1.6, 5.8, 18, "7C2D12");
  slide.addText("立刻能练的动作", { x: 7.1, y: 1.7, w: 3.2, h: 0.26, fontFace: "Microsoft YaHei", fontSize: 20, bold: true, color: hex(colors.primary), margin: 0 });
  addBullets(slide, chapter.child_actions, 7.1, 2.12, 5.2, 18, "334155");
  addFooter(slide, `${content.book.title} / 常见误区`, chapter.number, colors.primary);
  finalizeSlide(slide, ppt, helpers);

  slide = ppt.addSlide();
  slide.background = { color: hex(colors.dark) };
  slide.addText("一页带走", { x: 0.9, y: 0.9, w: 2.8, h: 0.36, fontFace: "Microsoft YaHei", fontSize: 28, bold: true, color: "FFFFFF", margin: 0 });
  slide.addText(chapter.one_line_review, { x: 0.92, y: 1.55, w: 5.2, h: 1.0, fontFace: "Microsoft YaHei", fontSize: 22, bold: true, color: hex(colors.secondary), margin: 0, breakLine: false });
  addBullets(slide, chapter.must_remember.slice(0, 5), 0.92, 2.65, 5.5, 17, "F8FAFC");
  if (fs.existsSync(card)) {
    slide.addImage({ path: card, ...helpers.imageSizingContain(card, 6.7, 1.45, 5.6, 4.8) });
  }
  addFooter(slide, `${content.book.title} / 一页带走`, chapter.number, colors.secondary);
  finalizeSlide(slide, ppt, helpers, true);

  return { ppt, fileBase };
}

function buildSpecialDeck(projectRoot, content, deck, helpers, PptxGenJS) {
  const ppt = baseDeck(PptxGenJS, `${content.book.title} ${deck.deck_title}`, `${content.book.title} 教学导读`);
  const reportImg = path.join(projectRoot, "04_配图", "阅读报告");
  const pptImg = path.join(projectRoot, "04_配图", "PPT");
  const poster = path.join(pptImg, `${deck.deck_id}_poster.png`);
  const overview = path.join(reportImg, "book_overview.png");
  const methods = path.join(reportImg, "common_methods.png");
  const colors = { primary: "#1D4ED8", secondary: "#93C5FD", dark: "#0F172A", bg: "#EFF6FF", accent: "#1E40AF" };
  const fileBase = `${deck.file_prefix}_${deck.deck_title}_儿童可读版`;

  const titles = ["封面", "为什么先看这一份", "这一份在讲什么", "结构总览", "分章路线", "三个阅读钥匙", "反复出现的方法", "最容易误会的地方", "读完会得到什么", "一页带走"];
  for (let i = 0; i < 10; i += 1) {
    const slide = ppt.addSlide();
    if (i === 0) {
      slide.background = { color: hex(colors.dark) };
      slide.addShape("roundRect", {
        x: 0.45, y: 0.4, w: 12.45, h: 6.7,
        rectRadius: 0.06,
        fill: { color: "102A43", transparency: 6 },
        line: { color: hex(colors.secondary), width: 2 },
      });
      slide.addText(content.book.title, { x: 0.85, y: 0.82, w: 4.2, h: 0.3, fontFace: "Microsoft YaHei", fontSize: 22, bold: true, color: "FFFFFF", margin: 0 });
      slide.addText(deck.deck_title, { x: 0.88, y: 1.4, w: 4.0, h: 0.8, fontFace: "Microsoft YaHei", fontSize: 30, bold: true, color: "FFFFFF", margin: 0 });
      slide.addText(deck.core_message, { x: 0.88, y: 2.45, w: 4.7, h: 1.3, fontFace: "Microsoft YaHei", fontSize: 20, color: "E2E8F0", margin: 0, breakLine: false });
      if (fs.existsSync(poster)) {
        slide.addImage({ path: poster, ...helpers.imageSizingContain(poster, 6.1, 1.0, 6.1, 5.4) });
      }
      addFooter(slide, `${content.book.title} / ${deck.deck_title}`, deck.file_prefix, colors.secondary);
      finalizeSlide(slide, ppt, helpers, true);
      continue;
    }
    slide.background = { color: i % 2 === 0 ? "FFFFFF" : hex(colors.bg) };
    addHeader(slide, titles[i], deck.deck_title, colors);
    if (i === 1) slide.addText(deck.why_it_matters, { x: 0.85, y: 1.55, w: 11.6, h: 4.8, fontFace: "Microsoft YaHei", fontSize: 21, color: "334155", margin: 0, breakLine: false });
    if (i === 2) slide.addText(deck.structure_summary, { x: 0.85, y: 1.55, w: 11.6, h: 4.8, fontFace: "Microsoft YaHei", fontSize: 21, color: "334155", margin: 0, breakLine: false });
    if (i === 3 && fs.existsSync(overview)) slide.addImage({ path: overview, ...helpers.imageSizingContain(overview, 0.85, 1.45, 11.8, 5.5) });
    if (i === 4) {
      const cards = deck.section_cards || [];
      cards.forEach((item, idx) => {
        const row = Math.floor(idx / 2);
        const col = idx % 2;
        const x = col === 0 ? 0.9 : 6.8;
        const y = row === 0 ? 1.6 : 4.1;
        slide.addShape("roundRect", { x, y, w: 5.5, h: 1.95, rectRadius: 0.05, fill: { color: idx % 2 === 0 ? "EFF6FF" : "FFFFFF" }, line: { color: hex(colors.secondary), width: 1.3 } });
        slide.addText(item.title, { x: x + 0.2, y: y + 0.18, w: 5.0, h: 0.22, fontFace: "Microsoft YaHei", fontSize: 19, bold: true, color: hex(colors.primary), margin: 0 });
        slide.addText(item.body, { x: x + 0.2, y: y + 0.58, w: 4.95, h: 0.95, fontFace: "Microsoft YaHei", fontSize: 16, color: "334155", margin: 0, breakLine: false });
      });
    }
    if (i === 5) addBullets(slide, deck.reading_keys, 1.0, 1.6, 11.0, 19, "334155");
    if (i === 6 && fs.existsSync(methods)) slide.addImage({ path: methods, ...helpers.imageSizingContain(methods, 0.85, 1.45, 11.8, 5.55) });
    if (i === 7) addBullets(slide, deck.misunderstandings, 1.0, 1.6, 11.0, 19, "7C2D12");
    if (i === 8) addBullets(slide, deck.gains, 1.0, 1.6, 11.0, 19, "334155");
    if (i === 9) {
      slide.addText(deck.one_line_review, { x: 0.95, y: 1.5, w: 5.2, h: 1.0, fontFace: "Microsoft YaHei", fontSize: 22, bold: true, color: hex(colors.primary), margin: 0, breakLine: false });
      addBullets(slide, deck.key_takeaways, 0.95, 2.6, 5.6, 18, "334155");
      addBullets(slide, deck.child_actions, 6.8, 2.6, 5.2, 18, "334155");
    }
    addFooter(slide, `${content.book.title} / ${deck.deck_title}`, deck.file_prefix, colors.primary);
    finalizeSlide(slide, ppt, helpers);
  }

  return { ppt, fileBase };
}

function stubSource(bookTitle, fileBase, projectRoot, deckType, deckLabel) {
  return `// ${bookTitle} / ${fileBase}
// 这是该 deck 的重建入口说明文件。
// 实际的版式模板与批量生成逻辑位于仓库根 scripts/book_ppt_builder.js
// 当前 deck 类型：${deckType}
// 当前 deck 标签：${deckLabel}
//
// 重新生成本项目全部 PPT：
//   node ../00_项目说明/生成脚本/build_ppts.js
//
// 如需查看原始教学内容数据：
//   ${path.relative(path.join(projectRoot, "03_PPT"), path.join(projectRoot, "05_中间素材", "chapter_content.json")).replace(/\\/g, "/")}
`;
}

async function buildBook(projectRoot) {
  const content = JSON.parse(fs.readFileSync(path.join(projectRoot, "05_中间素材", "chapter_content.json"), "utf8"));
  const helpers = loadHelpers(projectRoot);
  const PptxGenJS = loadPptxGen(projectRoot);
  const outputDir = path.join(projectRoot, "03_PPT");
  fs.mkdirSync(outputDir, { recursive: true });

  for (const deck of content.special_decks) {
    const { ppt, fileBase } = buildSpecialDeck(projectRoot, content, deck, helpers, PptxGenJS);
    const output = path.join(outputDir, `${fileBase}.pptx`);
    await ppt.writeFile({ fileName: output });
    fs.writeFileSync(path.join(outputDir, `${fileBase}.js`), stubSource(content.book.title, fileBase, projectRoot, "special", deck.deck_title), "utf8");
  }

  for (const chapter of content.chapters) {
    const { ppt, fileBase } = buildChapterDeck(projectRoot, content, chapter, helpers, PptxGenJS);
    const output = path.join(outputDir, `${fileBase}.pptx`);
    await ppt.writeFile({ fileName: output });
    fs.writeFileSync(path.join(outputDir, `${fileBase}.js`), stubSource(content.book.title, fileBase, projectRoot, "chapter", chapter.full_title), "utf8");
  }
}

module.exports = { buildBook };

if (require.main === module) {
  const projectRoot = process.argv[2];
  if (!projectRoot) {
    console.error("Usage: node scripts/book_ppt_builder.js <projectRoot>");
    process.exit(1);
  }
  buildBook(path.resolve(projectRoot)).catch((error) => {
    console.error(error);
    process.exit(1);
  });
}
