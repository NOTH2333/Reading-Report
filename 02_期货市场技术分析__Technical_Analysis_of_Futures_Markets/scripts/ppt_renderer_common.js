"use strict";

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const ROOT = path.resolve(__dirname, "..");
const helperDir = path.join(
  ROOT,
  "90_共用素材与底稿",
  "统一样式与图标",
  "pptxgenjs_helpers"
);
const {
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
} = require(path.join(helperDir, "layout.js"));

const shapeSource = new pptxgen();
const SHAPE = shapeSource._shapeType;

const COLORS = {
  bg: "F7F1E6",
  paper: "FFFDF8",
  ink: "253238",
  muted: "6B7280",
  line: "D9CCB8",
  band: "E9E1D4",
  teal: "0F766E",
  terracotta: "C84C31",
  blue: "1D4ED8",
  green: "2F8F5B",
  pale: "FFF8EE",
};

function addText(slide, text, opts) {
  slide.addText(text, {
    margin: 0,
    fontFace: "Microsoft YaHei",
    color: COLORS.ink,
    ...opts,
  });
}

function addChrome(slide, chapter, title, pageNo, total) {
  slide.background = { color: COLORS.bg };
  slide.addShape(SHAPE.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.62,
    fill: { color: COLORS.band },
    line: { color: COLORS.band, transparency: 100 },
  });
  slide.addShape(SHAPE.roundRect, {
    x: 0.58,
    y: 0.2,
    w: 0.95,
    h: 0.24,
    rectRadius: 0.08,
    fill: { color: COLORS.teal },
    line: { color: COLORS.teal, transparency: 100 },
  });
  addText(slide, `第${String(chapter.id).padStart(2, "0")}章`, {
    x: 0.66,
    y: 0.235,
    w: 0.8,
    h: 0.14,
    fontSize: 12,
    color: "FFFFFF",
    bold: true,
    align: "center",
  });
  addText(slide, title, {
    x: 0.6,
    y: 0.82,
    w: 8.9,
    h: 0.44,
    fontSize: 27,
    bold: true,
  });
  addText(slide, `${pageNo}/${total}`, {
    x: 12.15,
    y: 7.02,
    w: 0.6,
    h: 0.2,
    fontSize: 12,
    color: COLORS.muted,
    align: "right",
  });
  slide.addShape(SHAPE.line, {
    x: 0.6,
    y: 6.92,
    w: 12.1,
    h: 0,
    line: { color: COLORS.line, width: 1 },
  });
}

function addCard(slide, x, y, w, h, title, body, color = COLORS.teal) {
  slide.addShape(SHAPE.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    fill: { color: COLORS.paper },
    line: { color, width: 1.5 },
  });
  addText(slide, title, {
    x: x + 0.18,
    y: y + 0.14,
    w: w - 0.36,
    h: 0.2,
    fontSize: 16,
    bold: true,
    color,
  });
  addText(slide, body, {
    x: x + 0.18,
    y: y + 0.44,
    w: w - 0.36,
    h: h - 0.54,
    fontSize: 12.5,
    valign: "top",
  });
}

function bulletRuns(items) {
  return items.map((item, index) => ({
    text: item,
    options: {
      bullet: true,
      breakLine: index !== items.length - 1,
    },
  }));
}

function addImageFrame(slide, imagePath, x, y, w, h) {
  slide.addShape(SHAPE.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    fill: { color: COLORS.paper },
    line: { color: COLORS.line, width: 1.2 },
  });
  slide.addImage({ path: imagePath, x: x + 0.08, y: y + 0.08, w: w - 0.16, h: h - 0.16 });
}

function renderCover(slide, spec, chapter, chapterDir) {
  slide.background = { color: "F1E8D9" };
  slide.addShape(SHAPE.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 7.5,
    fill: { color: "F1E8D9" },
    line: { color: "F1E8D9", transparency: 100 },
  });
  slide.addShape(SHAPE.rect, {
    x: 8.95,
    y: 0,
    w: 4.383,
    h: 7.5,
    fill: { color: "E1D3BE" },
    line: { color: "E1D3BE", transparency: 100 },
  });
  addText(slide, `第${String(chapter.id).padStart(2, "0")}章`, {
    x: 0.82,
    y: 0.75,
    w: 1.2,
    h: 0.24,
    fontSize: 16,
    color: COLORS.teal,
    bold: true,
  });
  addText(slide, spec.title, {
    x: 0.82,
    y: 1.35,
    w: 6.9,
    h: 1.4,
    fontSize: 28,
    bold: true,
    color: COLORS.ink,
  });
  addText(slide, spec.subtitle, {
    x: 0.82,
    y: 3.15,
    w: 6.5,
    h: 0.34,
    fontSize: 16,
    color: COLORS.muted,
  });
  addText(slide, "儿童也能看懂的技术分析精讲", {
    x: 0.82,
    y: 4.2,
    w: 5.2,
    h: 0.36,
    fontSize: 18,
    color: COLORS.terracotta,
    bold: true,
  });
  addText(slide, "关键词：概念拆解｜图示重构｜误区提示｜本章收获", {
    x: 0.82,
    y: 4.68,
    w: 6.8,
    h: 0.3,
    fontSize: 14,
    color: COLORS.ink,
  });
  addImageFrame(slide, path.join(chapterDir, "assets", "02_核心图示.svg"), 9.45, 1.0, 3.15, 5.35);
}

function renderStandard(slide, spec, chapter, chapterDir, pageNo, total) {
  addChrome(slide, chapter, spec.title, pageNo, total);
  switch (spec.type) {
    case "objectives":
      spec.bullets.forEach((item, idx) => {
        const x = idx % 2 === 0 ? 0.8 : 6.95;
        const y = idx < 2 ? 1.55 : 3.95;
        addCard(slide, x, y, 5.55, 1.75, `目标 ${idx + 1}`, item, idx % 2 === 0 ? COLORS.teal : COLORS.terracotta);
      });
      break;
    case "map":
      spec.items.forEach((item, idx) => {
        const x = 0.95 + idx * 3.1;
        addCard(slide, x, 2.45, 2.55, 2.1, `模块 ${idx + 1}`, item, idx % 2 === 0 ? COLORS.teal : COLORS.blue);
        if (idx < spec.items.length - 1) {
          slide.addShape(SHAPE.line, {
            x: x + 2.55,
            y: 3.5,
            w: 0.42,
            h: 0,
            line: { color: COLORS.line, width: 2, endArrowType: "triangle" },
          });
        }
      });
      break;
    case "thesis":
      addCard(slide, 1.1, 2.1, 11.15, 2.6, "一句话主旨", spec.statement, COLORS.teal);
      addText(slide, "先抓主旨，再看细节，记忆会轻很多。", {
        x: 1.3,
        y: 5.15,
        w: 6.8,
        h: 0.28,
        fontSize: 14,
        color: COLORS.muted,
      });
      break;
    case "why":
      addCard(slide, 0.85, 1.65, 5.35, 4.7, "为什么重要", spec.statement, COLORS.terracotta);
      addCard(slide, 6.55, 1.65, 5.9, 4.7, "直观比喻", spec.analogy, COLORS.teal);
      break;
    case "keypoint":
      addCard(slide, 0.85, 1.65, 5.25, 4.7, spec.title, spec.idea, COLORS.teal);
      addCard(slide, 6.45, 1.65, 2.0, 2.2, "怎么看", spec.how, COLORS.blue);
      addCard(slide, 8.7, 1.65, 2.0, 2.2, "别踩坑", spec.mistake, COLORS.terracotta);
      addCard(slide, 10.95, 1.65, 1.55, 2.2, "带走", spec.harvest, COLORS.green);
      addImageFrame(slide, path.join(chapterDir, "assets", "01_概念地图.svg"), 6.45, 4.15, 6.05, 2.2);
      break;
    case "terms":
      spec.items.forEach((item, idx) => {
        addCard(slide, 1.0 + idx * 6.1, 2.0, 5.2, 3.8, item.term, item.explain, idx === 0 ? COLORS.teal : COLORS.terracotta);
      });
      break;
    case "diagram":
      addImageFrame(slide, path.join(chapterDir, spec.image), 0.95, 1.75, 6.35, 4.9);
      addCard(slide, 7.55, 1.95, 4.85, 2.2, "这张图在讲什么", spec.caption, COLORS.teal);
      addCard(slide, 7.55, 4.35, 4.85, 1.9, "阅读顺序", "先看背景，再看结构，再看确认，最后才考虑动作。", COLORS.terracotta);
      break;
    case "reference":
      addImageFrame(slide, path.join(chapterDir, spec.image), 0.95, 1.75, 6.0, 4.9);
      addCard(slide, 7.2, 2.0, 5.0, 1.85, "来源说明", spec.caption, COLORS.blue);
      addCard(slide, 7.2, 4.05, 5.0, 2.2, "看它的方式", "把原书页缩略图当作“原始证据”，重点观察结构，不要陷入逐字阅读。", COLORS.teal);
      break;
    case "deepdive":
      addText(slide, bulletRuns(spec.bullets), {
        x: 1.0,
        y: 1.8,
        w: 8.2,
        h: 4.9,
        fontSize: 18,
        breakLine: true,
      });
      addImageFrame(slide, path.join(chapterDir, "assets", "02_核心图示.svg"), 9.7, 2.0, 2.3, 2.7);
      break;
    case "cards":
      spec.items.forEach((item, idx) => {
        const x = idx % 2 === 0 ? 0.95 : 6.95;
        const y = idx < 2 ? 1.75 : 4.15;
        addCard(slide, x, y, 5.25, 1.9, `卡片 ${idx + 1}`, item, idx % 2 === 0 ? COLORS.teal : COLORS.blue);
      });
      break;
    case "misconceptions":
      spec.items.forEach((item, idx) => {
        addCard(
          slide,
          0.9,
          1.65 + idx * 1.7,
          11.6,
          1.35,
          `误区 ${idx + 1}：${item.question}`,
          item.answer,
          idx === 1 ? COLORS.terracotta : COLORS.teal
        );
      });
      break;
    case "application":
      spec.steps.forEach((step, idx) => {
        slide.addShape(SHAPE.ellipse, {
          x: 1.0,
          y: 1.8 + idx * 1.15,
          w: 0.45,
          h: 0.45,
          fill: { color: COLORS.pale },
          line: { color: COLORS.teal, width: 2 },
        });
        addText(slide, String(idx + 1), {
          x: 1.0,
          y: 1.915 + idx * 1.15,
          w: 0.45,
          h: 0.12,
          fontSize: 16,
          color: COLORS.teal,
          align: "center",
          bold: true,
        });
        addCard(slide, 1.75, 1.65 + idx * 1.15, 6.0, 0.72, `步骤 ${idx + 1}`, step, COLORS.teal);
      });
      addImageFrame(slide, path.join(chapterDir, "assets", "03_应用流程.svg"), 8.35, 1.85, 3.75, 4.45);
      break;
    case "takeaways":
      addText(slide, bulletRuns(spec.items), {
        x: 0.95,
        y: 1.8,
        w: 7.0,
        h: 4.6,
        fontSize: 18,
      });
      addCard(slide, 8.35, 2.0, 3.55, 2.8, "记忆方法", "把本章内容压缩成：背景、结构、验证、风险四步。", COLORS.terracotta);
      break;
    case "quiz":
      spec.items.forEach((item, idx) => {
        const x = idx % 2 === 0 ? 0.95 : 6.95;
        const y = idx < 4 ? 1.6 + Math.floor(idx / 2) * 1.8 : 5.2;
        const w = idx === 4 ? 11.25 : 5.2;
        addCard(slide, x, y, w, 1.35, `问题 ${idx + 1}`, item, idx % 2 === 0 ? COLORS.teal : COLORS.terracotta);
      });
      break;
    default:
      addText(slide, "未识别的幻灯片类型", { x: 1.0, y: 2.0, w: 6, h: 0.4, fontSize: 24, bold: true });
  }
}

async function renderChapterDeck(chapter, rootDir = ROOT) {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "OpenAI Codex";
  pptx.company = "OpenAI";
  pptx.subject = `${chapter.fullTitle} 精讲`;
  pptx.title = `${chapter.fullTitle} - ${chapter.targetSlides}页`;
  pptx.lang = "zh-CN";

  const chapterDir = path.join(rootDir, "02_章节PPT", chapter.dirName);
  fs.mkdirSync(chapterDir, { recursive: true });

  chapter.slides.forEach((spec, index) => {
    const slide = pptx.addSlide();
    if (spec.type === "cover") {
      renderCover(slide, spec, chapter, chapterDir);
    } else {
      renderStandard(slide, spec, chapter, chapterDir, index + 1, chapter.slides.length);
    }
    warnIfSlideHasOverlaps(slide, pptx);
    warnIfSlideElementsOutOfBounds(slide, pptx);
  });

  const outPath = path.join(chapterDir, chapter.pptFile);
  await pptx.writeFile({ fileName: outPath });
  return outPath;
}

module.exports = {
  renderChapterDeck,
};
