"use strict";

const fs = require("fs");
const path = require("path");
const PptxGenJS = require("pptxgenjs");
const { imageSizingContain } = require("../pptxgenjs_helpers/image");
const { safeOuterShadow } = require("../pptxgenjs_helpers/util");
const {
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
} = require("../pptxgenjs_helpers/layout");
const chapters = require("../src/chapter_data");
const SHAPE_TYPES = new PptxGenJS().ShapeType;

const ROOT = process.cwd();
const DELIVERABLES = path.join(ROOT, "deliverables");
const PPT_DIR = path.join(DELIVERABLES, "ppt");
const REPORT_DIR = path.join(DELIVERABLES, "reports");
const SUMMARY_DIR = path.join(DELIVERABLES, "summaries");
const ASSET_DIR = path.join(DELIVERABLES, "assets");
const WORK_DIR = path.join(ROOT, "_work");

function ensureDir(target) {
  fs.mkdirSync(target, { recursive: true });
}

function cleanOutput() {
  if (fs.existsSync(DELIVERABLES)) {
    fs.rmSync(DELIVERABLES, { recursive: true, force: true });
  }
  const checkDir = path.join(WORK_DIR, "checks");
  if (fs.existsSync(checkDir)) {
    fs.rmSync(checkDir, { recursive: true, force: true });
  }
  ensureDir(PPT_DIR);
  ensureDir(REPORT_DIR);
  ensureDir(SUMMARY_DIR);
  ensureDir(ASSET_DIR);
  ensureDir(WORK_DIR);
  ensureDir(checkDir);
}

function escapeXml(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function wrapText(text, width = 16) {
  const chars = text.split("");
  const lines = [];
  let line = "";
  for (const ch of chars) {
    line += ch;
    if (line.length >= width && "，。；：！？）】》 ".includes(ch)) {
      lines.push(line.trim());
      line = "";
    }
  }
  if (line.trim()) {
    lines.push(line.trim());
  }
  return lines;
}

function svgTag(inner, bg = "#ffffff", width = 1600, height = 900) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
<rect width="${width}" height="${height}" fill="${bg}"/>
${inner}
</svg>`;
}

function writeAsset(filePath, content) {
  ensureDir(path.dirname(filePath));
  fs.writeFileSync(filePath, content, "utf8");
}

function buildSceneSvg(chapter) {
  const lines = wrapText(chapter.sceneMessage, 18);
  const bubbleText = lines
    .map(
      (line, index) =>
        `<text x="930" y="${280 + index * 48}" font-size="34" font-family="Microsoft YaHei" fill="#1f2933">${escapeXml(
          line
        )}</text>`
    )
    .join("");
  const keywordBlocks = chapter.sceneKeywords
    .map((keyword, index) => {
      const x = 150 + (index % 2) * 230;
      const y = 610 + Math.floor(index / 2) * 94;
      return `
        <rect x="${x}" y="${y}" width="190" height="62" rx="20" fill="#ffffff" opacity="0.92"/>
        <text x="${x + 95}" y="${y + 39}" text-anchor="middle" font-size="28" font-family="Microsoft YaHei" fill="#243b53">${escapeXml(
          keyword
        )}</text>`;
    })
    .join("");
  return svgTag(
    `
    <defs>
      <linearGradient id="sky" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stop-color="#${chapter.palette.sky}"/>
        <stop offset="100%" stop-color="#${chapter.palette.soft}"/>
      </linearGradient>
      <linearGradient id="hill" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%" stop-color="#${chapter.palette.base}"/>
        <stop offset="100%" stop-color="#${chapter.palette.accent}"/>
      </linearGradient>
    </defs>
    <rect width="1600" height="900" fill="url(#sky)"/>
    <circle cx="1320" cy="110" r="64" fill="#ffffff" opacity="0.6"/>
    <path d="M0 640 C220 540 430 760 620 650 C760 570 860 520 980 600 C1140 710 1290 650 1600 570 L1600 900 L0 900 Z" fill="url(#hill)" opacity="0.22"/>
    <rect x="96" y="92" width="620" height="716" rx="44" fill="#ffffff" opacity="0.82"/>
    <text x="146" y="180" font-size="58" font-family="Microsoft YaHei" font-weight="700" fill="#${chapter.palette.deep}">第 ${chapter.num} 章</text>
    <text x="146" y="250" font-size="44" font-family="Microsoft YaHei" font-weight="700" fill="#243b53">${escapeXml(
      chapter.sceneTitle
    )}</text>
    <rect x="150" y="320" width="460" height="220" rx="28" fill="#${chapter.palette.soft}" opacity="0.95"/>
    <circle cx="250" cy="408" r="66" fill="#${chapter.palette.base}"/>
    <rect x="205" y="476" width="92" height="146" rx="26" fill="#${chapter.palette.base}"/>
    <rect x="330" y="350" width="240" height="170" rx="18" fill="#ffffff" stroke="#${chapter.palette.accent}" stroke-width="6"/>
    <line x1="330" y1="398" x2="570" y2="398" stroke="#${chapter.palette.accent}" stroke-width="6"/>
    <path d="M364 443 L424 404 L483 446 L533 414" stroke="#${chapter.palette.base}" stroke-width="10" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
    ${keywordBlocks}
    <rect x="820" y="180" width="650" height="360" rx="40" fill="#ffffff" opacity="0.92" stroke="#${chapter.palette.base}" stroke-width="6"/>
    ${bubbleText}
    <rect x="880" y="610" width="520" height="150" rx="34" fill="#${chapter.palette.base}" opacity="0.12"/>
    <text x="920" y="670" font-size="36" font-family="Microsoft YaHei" font-weight="700" fill="#${chapter.palette.deep}">本章主线</text>
    <text x="920" y="722" font-size="30" font-family="Microsoft YaHei" fill="#1f2933">${escapeXml(
      chapter.elevator
    )}</text>
  `,
    `#${chapter.palette.sky}`
  );
}

function buildConceptSvg(chapter) {
  const positions = [
    [800, 170],
    [480, 340],
    [1120, 340],
    [500, 610],
    [1100, 610],
  ];
  const nodes = chapter.coreTerms
    .slice(0, 5)
    .map((term, index) => {
      const [x, y] = positions[index];
      const descLines = wrapText(term[1], 11)
        .slice(0, 3)
        .map(
          (line, lineIndex) =>
            `<text x="${x}" y="${y + 52 + lineIndex * 28}" text-anchor="middle" font-size="22" font-family="Microsoft YaHei" fill="#334e68">${escapeXml(
              line
            )}</text>`
        )
        .join("");
      return `
      <circle cx="${x}" cy="${y}" r="92" fill="#ffffff" stroke="#${index === 0 ? chapter.palette.base : chapter.palette.accent}" stroke-width="6"/>
      <text x="${x}" y="${y - 10}" text-anchor="middle" font-size="30" font-family="Microsoft YaHei" font-weight="700" fill="#${chapter.palette.deep}">${escapeXml(
        term[0]
      )}</text>
      ${descLines}`;
    })
    .join("");
  const links = [
    [800, 170, 480, 340],
    [800, 170, 1120, 340],
    [480, 340, 500, 610],
    [1120, 340, 1100, 610],
  ]
    .map(
      ([x1, y1, x2, y2]) =>
        `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="#${chapter.palette.base}" stroke-width="8" opacity="0.35"/>`
    )
    .join("");
  return svgTag(
    `
    <rect x="80" y="60" width="1440" height="780" rx="40" fill="#ffffff" opacity="0.94"/>
    <text x="120" y="140" font-size="46" font-family="Microsoft YaHei" font-weight="700" fill="#${chapter.palette.deep}">本章概念地图</text>
    <text x="120" y="190" font-size="28" font-family="Microsoft YaHei" fill="#486581">把书里的抽象术语翻成能看得见的关系图。</text>
    ${links}
    ${nodes}
  `,
    `#${chapter.palette.sky}`
  );
}

function buildMisconceptionSvg(chapter) {
  const rows = chapter.misconceptions
    .map((item, index) => {
      const y = 180 + index * 190;
      const mythLines = wrapText(item.myth, 18)
        .map(
          (line, lineIndex) =>
            `<text x="175" y="${y + 72 + lineIndex * 34}" font-size="28" font-family="Microsoft YaHei" fill="#7f1d1d">${escapeXml(
              line
            )}</text>`
        )
        .join("");
      const truthLines = wrapText(item.truth, 18)
        .map(
          (line, lineIndex) =>
            `<text x="865" y="${y + 72 + lineIndex * 34}" font-size="28" font-family="Microsoft YaHei" fill="#14532d">${escapeXml(
              line
            )}</text>`
        )
        .join("");
      return `
      <rect x="110" y="${y}" width="580" height="136" rx="28" fill="#FEE2E2"/>
      <rect x="810" y="${y}" width="580" height="136" rx="28" fill="#DCFCE7"/>
      <text x="155" y="${y + 40}" font-size="26" font-family="Microsoft YaHei" font-weight="700" fill="#991B1B">常见误区</text>
      <text x="855" y="${y + 40}" font-size="26" font-family="Microsoft YaHei" font-weight="700" fill="#166534">作者真正想说</text>
      ${mythLines}
      ${truthLines}
      <path d="M710 ${y + 68} L790 ${y + 68}" stroke="#${chapter.palette.base}" stroke-width="8" stroke-linecap="round"/>
      <path d="M772 ${y + 46} L798 ${y + 68} L772 ${y + 90}" fill="none" stroke="#${chapter.palette.base}" stroke-width="8" stroke-linecap="round" stroke-linejoin="round"/>`;
    })
    .join("");
  return svgTag(
    `
    <text x="110" y="100" font-size="46" font-family="Microsoft YaHei" font-weight="700" fill="#${chapter.palette.deep}">误区与修正</text>
    <text x="110" y="145" font-size="28" font-family="Microsoft YaHei" fill="#486581">把“听上去很像对的话”翻回书里真正的意思。</text>
    ${rows}
  `,
    `#${chapter.palette.sky}`
  );
}

function buildActionSvg(chapter) {
  const steps = chapter.checklist
    .map((item, index) => {
      const y = 165 + index * 150;
      const lines = wrapText(item, 22)
        .map(
          (line, lineIndex) =>
            `<text x="340" y="${y + 48 + lineIndex * 30}" font-size="28" font-family="Microsoft YaHei" fill="#243b53">${escapeXml(
              line
            )}</text>`
        )
        .join("");
      return `
      <circle cx="195" cy="${y + 28}" r="42" fill="#${chapter.palette.base}"/>
      <text x="195" y="${y + 38}" text-anchor="middle" font-size="34" font-family="Georgia" font-weight="700" fill="#ffffff">${
        index + 1
      }</text>
      <rect x="260" y="${y - 18}" width="1080" height="102" rx="30" fill="#ffffff" opacity="0.92"/>
      ${lines}`;
    })
    .join("");
  return svgTag(
    `
    <text x="120" y="105" font-size="46" font-family="Microsoft YaHei" font-weight="700" fill="#${chapter.palette.deep}">行动清单</text>
    <text x="120" y="150" font-size="28" font-family="Microsoft YaHei" fill="#486581">把本章道理压缩成可以直接执行的动作。</text>
    <rect x="110" y="190" width="1300" height="560" rx="42" fill="#${chapter.palette.soft}" opacity="0.72"/>
    ${steps}
  `,
    `#${chapter.palette.sky}`
  );
}

function buildOverviewRoadmap() {
  const chapterCards = chapters
    .map((chapter, index) => {
      const row = Math.floor(index / 4);
      const col = index % 4;
      const x = 110 + col * 355;
      const y = 170 + row * 210;
      return `
      <rect x="${x}" y="${y}" width="300" height="136" rx="24" fill="#ffffff" opacity="0.95" stroke="#${chapter.palette.base}" stroke-width="4"/>
      <text x="${x + 28}" y="${y + 46}" font-size="28" font-family="Microsoft YaHei" font-weight="700" fill="#${chapter.palette.deep}">第 ${chapter.num} 章</text>
      <text x="${x + 28}" y="${y + 88}" font-size="24" font-family="Microsoft YaHei" fill="#243b53">${escapeXml(chapter.titleZh)}</text>`;
    })
    .join("");
  return svgTag(
    `
    <rect x="70" y="50" width="1460" height="800" rx="42" fill="#FFF8F1"/>
    <text x="120" y="132" font-size="52" font-family="Microsoft YaHei" font-weight="700" fill="#2D3748">《交易心理分析》全书路线图</text>
    <text x="120" y="180" font-size="28" font-family="Microsoft YaHei" fill="#486581">从分析工具，走到心理结构，再走到概率与信念的重建。</text>
    ${chapterCards}
  `,
    "#FFFDF9"
  );
}

function buildFiveTruthsSvg() {
  const truths = [
    "你不需要知道下一步会发生什么，仍然可以赚钱。",
    "任何事情都可能发生。",
    "每个时刻都是独特的。",
    "优势只是表示某种结果概率更高，不是保证。",
    "每一次样本都只是整体统计的一部分。",
  ];
  const items = truths
    .map((truth, index) => {
      const y = 170 + index * 125;
      return `
      <circle cx="180" cy="${y}" r="32" fill="#2A9D8F"/>
      <text x="180" y="${y + 11}" text-anchor="middle" font-size="26" font-family="Georgia" font-weight="700" fill="#ffffff">${
        index + 1
      }</text>
      <rect x="250" y="${y - 38}" width="1110" height="78" rx="24" fill="#ffffff"/>
      <text x="285" y="${y + 11}" font-size="30" font-family="Microsoft YaHei" fill="#1f2933">${escapeXml(
        truth
      )}</text>`;
    })
    .join("");
  return svgTag(
    `
    <text x="120" y="110" font-size="50" font-family="Microsoft YaHei" font-weight="700" fill="#234E52">五个基础事实</text>
    <text x="120" y="155" font-size="28" font-family="Microsoft YaHei" fill="#486581">第 7-8 章的核心。真正的交易者心态，是围绕这些事实建立起来的。</text>
    ${items}
  `,
    "#F4FFFD"
  );
}

function buildPsychologyLoopSvg() {
  return svgTag(
    `
    <text x="120" y="110" font-size="50" font-family="Microsoft YaHei" font-weight="700" fill="#7D2E22">普通交易者的心理回路</text>
    <text x="120" y="155" font-size="28" font-family="Microsoft YaHei" fill="#486581">全书反复想打断的，就是这条会自我强化的循环。</text>
    <circle cx="330" cy="420" r="120" fill="#FCE4CC"/><text x="330" y="405" text-anchor="middle" font-size="32" font-family="Microsoft YaHei" font-weight="700" fill="#8B4E1B">想预测对</text><text x="330" y="448" text-anchor="middle" font-size="24" font-family="Microsoft YaHei" fill="#7C5A38">把交易等同于证明自己</text>
    <circle cx="780" cy="260" r="120" fill="#FEE2E2"/><text x="780" y="245" text-anchor="middle" font-size="32" font-family="Microsoft YaHei" font-weight="700" fill="#991B1B">遇到不确定</text><text x="780" y="288" text-anchor="middle" font-size="24" font-family="Microsoft YaHei" fill="#7C5A38">市场不照“应该”走</text>
    <circle cx="1230" cy="420" r="120" fill="#E6DDEA"/><text x="1230" y="405" text-anchor="middle" font-size="32" font-family="Microsoft YaHei" font-weight="700" fill="#4C1D95">情绪放大</text><text x="1230" y="448" text-anchor="middle" font-size="24" font-family="Microsoft YaHei" fill="#7C5A38">恐惧、贪婪、报复</text>
    <circle cx="780" cy="610" r="120" fill="#DDE6EF"/><text x="780" y="595" text-anchor="middle" font-size="32" font-family="Microsoft YaHei" font-weight="700" fill="#17313B">行为扭曲</text><text x="780" y="638" text-anchor="middle" font-size="24" font-family="Microsoft YaHei" fill="#7C5A38">拖损、抢跑、乱加仓</text>
    <path d="M446 372 C550 300 620 280 655 270" stroke="#E76F51" stroke-width="10" fill="none" stroke-linecap="round"/>
    <path d="M900 270 C1010 285 1085 330 1125 360" stroke="#E76F51" stroke-width="10" fill="none" stroke-linecap="round"/>
    <path d="M1155 540 C1070 590 965 622 900 620" stroke="#E76F51" stroke-width="10" fill="none" stroke-linecap="round"/>
    <path d="M660 622 C565 615 470 565 405 520" stroke="#E76F51" stroke-width="10" fill="none" stroke-linecap="round"/>
  `,
    "#FFF9F3"
  );
}

function buildConsistencyLadderSvg() {
  const steps = [
    "认识到问题不只在分析，更在心理结构",
    "接受市场不确定且每个样本独立",
    "建立与事实匹配的信念",
    "用规则和练习把信念变成动作",
    "在长期样本中形成一致性",
  ];
  const blocks = steps
    .map((step, index) => {
      const y = 650 - index * 110;
      const width = 320 + index * 150;
      const x = 800 - width / 2;
      return `
      <rect x="${x}" y="${y}" width="${width}" height="78" rx="24" fill="${
        ["#D8F1EC", "#FCE4CC", "#E7F2E3", "#E6DDEA", "#DDE6EF"][index]
      }" stroke="#355070" stroke-width="3"/>
      <text x="800" y="${y + 47}" text-anchor="middle" font-size="28" font-family="Microsoft YaHei" fill="#1f2933">${escapeXml(
        step
      )}</text>`;
    })
    .join("");
  return svgTag(
    `
    <text x="120" y="110" font-size="50" font-family="Microsoft YaHei" font-weight="700" fill="#355070">从读懂到做到的一致性阶梯</text>
    <text x="120" y="155" font-size="28" font-family="Microsoft YaHei" fill="#486581">全书真正的成长路径，不是“看准更多”，而是“思维结构越来越匹配市场”。</text>
    ${blocks}
  `,
    "#F8FBFF"
  );
}

function buildAssets() {
  for (const chapter of chapters) {
    const chapterDir = path.join(ASSET_DIR, `ch${chapter.id}`);
    ensureDir(chapterDir);
    writeAsset(path.join(chapterDir, `ch${chapter.id}-scene.svg`), buildSceneSvg(chapter));
    writeAsset(path.join(chapterDir, `ch${chapter.id}-concepts.svg`), buildConceptSvg(chapter));
    writeAsset(path.join(chapterDir, `ch${chapter.id}-myths.svg`), buildMisconceptionSvg(chapter));
    writeAsset(path.join(chapterDir, `ch${chapter.id}-actions.svg`), buildActionSvg(chapter));
  }
  const commonDir = path.join(ASSET_DIR, "common");
  ensureDir(commonDir);
  writeAsset(path.join(commonDir, "book-roadmap.svg"), buildOverviewRoadmap());
  writeAsset(path.join(commonDir, "five-truths.svg"), buildFiveTruthsSvg());
  writeAsset(path.join(commonDir, "psychology-loop.svg"), buildPsychologyLoopSvg());
  writeAsset(path.join(commonDir, "consistency-ladder.svg"), buildConsistencyLadderSvg());
}

function summaryMarkdown(chapter) {
  const imageBase = `../assets/ch${chapter.id}`;
  const sectionBlocks = chapter.sections
    .map(
      (section, index) => `### ${index + 1}. ${section.title}

${section.summary}

孩子也能懂的说法：
${section.child}

放回交易里看：
${section.market}
`
    )
    .join("\n");
  const myths = chapter.misconceptions
    .map((item) => `- 误区：${item.myth}\n- 修正：${item.truth}`)
    .join("\n");
  const cards = chapter.memoryCards.map((item) => `- ${item}`).join("\n");
  const actions = chapter.checklist.map((item) => `- ${item}`).join("\n");
  return `# 第 ${chapter.num} 章｜${chapter.titleZh}

## 一句话主旨

${chapter.elevator}

![第 ${chapter.num} 章场景图](${imageBase}/ch${chapter.id}-scene.svg)

## 这章到底在解决什么问题

${chapter.question}

为什么这章重要：
${chapter.whyImportant}

## 关键知识点

![第 ${chapter.num} 章概念图](${imageBase}/ch${chapter.id}-concepts.svg)

${chapter.coreTerms.map((term) => `- **${term[0]}**：${term[1]}`).join("\n")}

## 按章节内容展开

${sectionBlocks}

## 孩子也能记住的类比

**${chapter.childAnalogy.title}**

${chapter.childAnalogy.story}

这个类比想说明：
${chapter.childAnalogy.lesson}

## 常见错误

![第 ${chapter.num} 章误区图](${imageBase}/ch${chapter.id}-myths.svg)

${myths}

## 记忆卡片

${cards}

## 行动清单

![第 ${chapter.num} 章行动图](${imageBase}/ch${chapter.id}-actions.svg)

${actions}
`;
}

function overviewNarrative() {
  return `# 《交易心理分析》阅读报告

![全书路线图](../assets/common/book-roadmap.svg)

## 导论：这本书为什么值得做成一整套学习包

马克·道格拉斯写这本书时，真正想解决的并不是“哪一种指标更好”，也不是“怎样在下一次交易里精准押中方向”。他想回答的，是一个更深也更痛的问题：为什么很多聪明、勤奋、知识丰富、职业上很成功的人，一旦走进交易市场，却会反复输给自己？作者在前言和引子里不断强调，市场之所以让人挫败，不是因为机会太少，而是因为市场的运行方式和大多数人从小被训练出来的思考方式并不匹配。学校奖励确定答案，职场奖励可控推进，社交奖励说服与管理；市场却要求人接受不确定、允许亏损、把单笔结果看成样本，并在没有外部监督的情况下仍能执行规则。

从这个意义上说，《交易心理分析》并不是一本教人“猜涨跌”的书，而是一本教人重建认知与行为结构的书。它从基本面与技术面的争论切入，逐步把焦点挪到心理结构，再进一步进入概率思维、信念机制与行为训练，最后落到一种可以通过机械练习逐步安装的交易者心智。它真正挑战读者的，不是智力，而是旧有信念：我必须看对，我不能亏，我的判断应该被市场尊重，我只要更努力分析就会解决问题。作者几乎整本书都在拆这些信念，再把一套与市场现实更匹配的新信念装回去。

## 作者的问题意识：为什么技术问题最后会变成心理问题

作者先让读者看到一个常被忽略的事实：市场分析有价值，但分析能力并不能自动变成稳定盈利。基本面分析之所以难以直接转化为一致结果，是因为价格真正的推动者是人，而不是模型；技术分析之所以比基本面更接近实战，是因为它能读取市场参与者反复出现的行为模式；但仅有技术分析仍然不够，因为“看懂机会”和“执行机会”之间仍然隔着一层心理落差。很多人会在图上看见信号，却在真实下单时变形：该进不进、该损不损、该收不收、该停不停。这不是知识缺口，而是结构缺口。

![心理回路图](../assets/common/psychology-loop.svg)

作者的问题意识因此非常清晰：交易中的绝大多数持续痛苦，并不是来自市场本身，而是来自人把市场个人化、道德化、人格化之后产生的内在冲突。你越要求市场配合你的期待，越会痛；你越把单笔输赢和自我价值绑定，越会扭曲；你越想用日常社会里的控制技能去管理市场，越会发现这些技能在这里频频失灵。作者真正要教的，是怎样把自己从这种失配中解放出来。

## 全书总论点：稳定盈利来自一种与市场现实匹配的心智结构

全书的总论点可以压缩成一句话：稳定盈利并不首先来自更高明的预测，而来自一种与市场现实相匹配的心智结构。这种结构有几个基础特征。第一，它承认市场不确定，而且这种不确定不是缺点，而是环境特征。第二，它承认优势只意味着概率更高，而不是结果保证。第三，它接受亏损是经营成本，而不是人格羞辱。第四，它把单笔交易放回长期样本中理解，不要求每一笔都证明自己正确。第五，它愿意通过纪律和练习，把这些看似抽象的原则变成大脑的默认反应。

![五个基础事实](../assets/common/five-truths.svg)

因此，这本书反复在做两件事：一件是拆除旧心智，另一件是安装新心智。拆除的对象包括对确定性的迷恋、对亏损的羞耻、对市场的对抗姿态、对随机奖励的上瘾，以及对自我形象的过度保护。安装的对象，则包括概率思维、样本意识、接受风险、去人格化、内部控制、规则执行和信念重塑。作者并不把这些当作鸡汤，而是当作一套可以被训练、被验证、被复盘的心理工程。

## 11 章论证链：作者如何一步步把读者带到“像交易者一样思考”
`;
}

function chapterReportBlock(chapter) {
  const intro = `### 第 ${chapter.num} 章：${chapter.titleZh}

![第 ${chapter.num} 章图解](../assets/ch${chapter.id}/ch${chapter.id}-scene.svg)

这一章围绕“${chapter.question.replace(/？$/, "")}”展开。作者先抛出一个让多数交易者不舒服的前提：问题很可能不在市场本身，而在你理解市场、理解风险和理解自己的方式。${chapter.whyImportant}`;
  const sections = chapter.sections
    .map(
      (section) => `- **${section.title}**：${section.summary} 这部分如果用孩子也能懂的话来翻译，就是：${section.child} 放回交易场景里，作者真正想逼读者看到的是：${section.market}`
    )
    .join("\n");
  const conclusion = `这一章最后落到三个层面。第一，概念层面，作者把 ${chapter.coreTerms
    .slice(0, 3)
    .map((term) => `“${term[0]}”`)
    .join("、")} 等关键词重新定义；第二，行为层面，他提醒读者别再掉进 ${chapter.misconceptions
    .map((item) => `“${item.myth}”`)
    .join("、")} 这些看似合理却会持续伤害执行的误区；第三，训练层面，他给出的方向是：${chapter.checklist
    .slice(0, 2)
    .join("；")}。这意味着本章不是让读者多背一个概念，而是让读者换一种默认反应。`;
  return `${intro}

${sections}

${conclusion}
`;
}

function reportClosing() {
  return `
## 关键概念之间的关系：从“看懂市场”到“管理自己”

![一致性阶梯](../assets/common/consistency-ladder.svg)

如果把全书当成一个完整模型，它至少包含四层。第一层是市场层，作者要求读者承认市场具有不确定、独特、无限表达的特性。第二层是认知层，人并不是直接看见市场，而是通过信念、记忆、联想和自我评估在看市场。第三层是行为层，一旦认知带着冲突，行为就会变形，表现为拖延止损、报复交易、过度期待、随意加减仓等。第四层是训练层，要想从不稳定走向一致，就必须通过规则、纪律、观察、样本练习和信念重塑，把新的理解安装成新的默认行为。

换句话说，作者想让交易者完成的，不是“更会分析市场”这一件事，而是“更会在分析之后不毁掉自己”。这也是为什么全书不断提醒：你不必知道下一步会发生什么，仍然可以赚钱；任何事情都可能发生；每个时刻都是独特的；优势只是一种偏向；长期结果来自一连串样本。只要这些事实没有真正进入你的行为系统，你就会在压力来时回到旧人格、旧防御和旧习惯。

## 可执行训练框架：把书里的思想变成可以练习的一周流程

作者在第 11 章给出机械阶段练习，其实可以整理成更现代的周训练框架。第一步，选择一个清晰、有限、可检验的优势，不同时改太多变量。第二步，在进场前写清楚风险、失效条件和退出条件，确保最坏结果已经被接受。第三步，连续执行一组样本，中途不因为单笔输赢随意改规则。第四步，复盘时同时看两张图：市场图和心理图。市场图看机会是否成立，心理图看自己是否因为恐惧、希望、自尊和冲动而偏离。第五步，每周只修正一两个最关键的行为偏差，而不是在情绪里整套系统推倒重来。

这种训练框架的重要意义在于，它把“改变自己”从模糊愿望变成一套可被重复的流程。你不再只是希望自己更冷静、更自律、更像高手，而是在每天的具体动作里一点点降低旧信念的控制权，增加新结构的存在感。久而久之，纪律会从外在支架慢慢变成自然反应，市场信息也会因为内部冲突减少而变得更清晰。

## 局限与误读提醒：这本书不是万能钥匙

这本书极有力量，但也容易被误读。第一种误读是把它当成“只要心态好就不需要系统”。作者其实明确假设你已经有某种优势，至少有一套可定义的模式识别方法；没有优势，心态再好也只是在稳定地乱做。第二种误读是把“接受风险”理解成对亏损麻木。作者要的不是麻木，而是事先清楚、事中不扭曲、事后能学习。第三种误读是把概率思维理解成“什么都无所谓”。恰恰相反，概率思维要求更严格的样本意识和更诚实的执行纪律。第四种误读是把“状态”神秘化。作者说的状态，本质上是当事实被接受、冲突被减少后，一种高质量、低内耗的专注工作状态，不是玄学，不是灵感附体。

## 结论：全书真正想改造的不是方法，而是身份

读完整本书会发现，作者真正想改造的不是一两条操作规则，而是交易者对自己的身份理解。普通交易者常把自己理解成一个必须不断证明自己正确的人，所以市场一旦不给面子，自尊、防御、报复和逃避就会一起出现。作者想建立的则是另一种身份：我是一个在不确定环境中经营优势、管理风险、长期训练自己的人。这个身份更谦逊，也更强大。它不靠每一笔都赢来维持体面，却能在长期里积累真正的稳定。

如果说这本书留下的最重要收获是什么，那大概就是：市场从来不是最难的部分，真正困难的是让自己变成一个配得上市场现实的人。只有当人停止向市场索要确定性，转而训练自己与不确定性合作，交易才会从一场不断受伤的个人战争，变成一门可以长期经营的专业工作。
`;
}

function buildReport() {
  let content = overviewNarrative();
  for (const chapter of chapters) {
    content += `\n${chapterReportBlock(chapter)}\n`;
  }
  content += reportClosing();
  fs.writeFileSync(path.join(REPORT_DIR, "阅读报告.md"), content, "utf8");
}

function buildSummaries() {
  const indexItems = chapters
    .map((chapter) => `- [第 ${chapter.num} 章｜${chapter.titleZh}](./${chapter.id}-${chapter.titleZh}.md)`)
    .join("\n");
  fs.writeFileSync(
    path.join(SUMMARY_DIR, "00-章节索引.md"),
    `# 章节知识索引

这套索引按照原书章节顺序排列，可直接跳转到对应章节知识总结。

${indexItems}
`,
    "utf8"
  );
  for (const chapter of chapters) {
    fs.writeFileSync(
      path.join(SUMMARY_DIR, `${chapter.id}-${chapter.titleZh}.md`),
      summaryMarkdown(chapter),
      "utf8"
    );
  }
}

function buildReadme() {
  const deckLinks = chapters
    .map(
      (chapter) =>
        `- 第 ${chapter.num} 章：\`ppt/${chapter.id}-${chapter.titleZh}.pptx\` 与 \`ppt/${chapter.id}-${chapter.titleZh}.js\``
    )
    .join("\n");
  const summaryLinks = chapters
    .map((chapter) => `- \`summaries/${chapter.id}-${chapter.titleZh}.md\``)
    .join("\n");
  const content = `# 《交易心理分析》学习包

## 目录

- 总阅读报告：\`reports/阅读报告.md\`
- 章节索引：\`summaries/00-章节索引.md\`
- 章节 PPT：
${deckLinks}
- 章节知识总结：
${summaryLinks}
- 配图资产：\`assets/ch01\` 至 \`assets/ch11\`，以及 \`assets/common\`

## 设计说明

- 语言：中文为主，关键英文术语只在需要时补充。
- 风格：图解启蒙风，适合零基础读者阅读，也适合二次讲解。
- 配图：全部为本地原创 SVG 图解与场景图，不依赖外部版权图片。

## 生成链

- 文本抽取：\`python scripts/extract_book.py\`
- 统一生成：\`node scripts/generate_learning_pack.js\`
- 抽样校验：\`powershell -ExecutionPolicy Bypass -File scripts/validate_decks.ps1\`
`;
  fs.writeFileSync(path.join(DELIVERABLES, "README.md"), content, "utf8");
}

function chapterAsset(chapter, suffix) {
  return path.join(ASSET_DIR, `ch${chapter.id}`, `ch${chapter.id}-${suffix}.svg`);
}

function bulletLines(slide, items, box, options = {}) {
  const runs = [];
  items.forEach((item, index) => {
    runs.push({
      text: `${item}`,
      options: {
        bullet: { indent: options.bulletIndent || 16 },
        hanging: options.hanging || 2,
        breakLine: index !== items.length - 1,
      },
    });
  });
  slide.addText(runs, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    fontFace: "Microsoft YaHei",
    fontSize: options.fontSize || 22,
    color: options.color || "243B53",
    margin: 0.08,
    valign: "top",
    paraSpaceAfterPt: 10,
  });
}

function addShell(slide, chapter, title, subtitle) {
  slide.addShape(SHAPE_TYPES.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 7.5,
    fill: { color: chapter.palette.sky },
    line: { color: chapter.palette.sky },
  });
  slide.addShape(SHAPE_TYPES.rect, {
    x: 0.5,
    y: 0.35,
    w: 12.3,
    h: 6.8,
    fill: { color: "FFFFFF", transparency: 4 },
    line: { color: chapter.palette.soft, width: 1.2 },
    radius: 0.2,
    shadow: safeOuterShadow("A0AEC0", 0.14, 45, 2, 1),
  });
  slide.addShape(SHAPE_TYPES.line, {
    x: 0.85,
    y: 0.92,
    w: 11.7,
    h: 0,
    line: { color: chapter.palette.soft, width: 1.2 },
  });
  slide.addText(`第 ${chapter.num} 章`, {
    x: 0.95,
    y: 0.55,
    w: 1.4,
    h: 0.35,
    fontFace: "Georgia",
    fontSize: 20,
    bold: true,
    color: chapter.palette.base,
  });
  slide.addText(title, {
    x: 0.95,
    y: 1.0,
    w: 5.4,
    h: 0.42,
    fontFace: "Microsoft YaHei",
    fontSize: 24,
    bold: true,
    color: chapter.palette.deep,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.95,
      y: 1.58,
      w: 5.45,
      h: 0.28,
      fontFace: "Microsoft YaHei",
      fontSize: 12.5,
      color: "486581",
    });
  }
}

function createDeck(chapter) {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "OpenAI Codex";
  pptx.company = "OpenAI";
  pptx.subject = `《交易心理分析》第 ${chapter.num} 章学习包`;
  pptx.title = `第 ${chapter.num} 章 - ${chapter.titleZh}`;
  pptx.lang = "zh-CN";
  pptx.theme = {
    headFontFace: "Microsoft YaHei",
    bodyFontFace: "Microsoft YaHei",
    lang: "zh-CN",
  };

  const scene = chapterAsset(chapter, "scene");
  const concepts = chapterAsset(chapter, "concepts");
  const myths = chapterAsset(chapter, "myths");
  const actions = chapterAsset(chapter, "actions");

  const cover = pptx.addSlide();
  cover.background = { color: chapter.palette.sky };
  cover.addImage({ path: scene, ...imageSizingContain(scene, 0.75, 1.3, 11.8, 3.95) });
  cover.addShape(SHAPE_TYPES.rect, {
    x: 0.85,
    y: 0.72,
    w: 2.1,
    h: 0.52,
    fill: { color: chapter.palette.base, transparency: 8 },
    line: { color: chapter.palette.base },
    radius: 0.15,
  });
  cover.addText(`第 ${chapter.num} 章`, {
    x: 1.05,
    y: 0.86,
    w: 0.9,
    h: 0.2,
    fontFace: "Georgia",
    fontSize: 20,
    bold: true,
    color: "FFFFFF",
  });
  cover.addText(chapter.titleZh, {
    x: 0.95,
    y: 5.42,
    w: 8.6,
    h: 0.34,
    fontFace: "Microsoft YaHei",
    fontSize: 22,
    bold: true,
    color: chapter.palette.deep,
  });
  cover.addText(chapter.elevator, {
    x: 0.95,
    y: 5.92,
    w: 8.7,
    h: 0.42,
    fontFace: "Microsoft YaHei",
    fontSize: 12.5,
    color: "334E68",
    margin: 0.05,
  });
  cover.addText("给零基础读者的图解启蒙版", {
    x: 9.65,
    y: 6.42,
    w: 2.2,
    h: 0.3,
    align: "right",
    fontFace: "Microsoft YaHei",
    fontSize: 12,
    color: chapter.palette.base,
  });

  const whySlide = pptx.addSlide();
  addShell(whySlide, chapter, "这一章在回答什么问题？", "先抓住问题、价值和本章阅读地图，再进入分节理解。");
  whySlide.addText(chapter.question, {
    x: 1.0,
    y: 2.1,
    w: 5.0,
    h: 1.1,
    fontFace: "Microsoft YaHei",
    fontSize: 24,
    bold: true,
    color: chapter.palette.deep,
  });
  whySlide.addText(chapter.whyImportant, {
    x: 1.0,
    y: 3.25,
    w: 5.0,
    h: 2.6,
    fontFace: "Microsoft YaHei",
    fontSize: 15,
    color: "334E68",
    margin: 0.04,
    valign: "top",
  });
  whySlide.addImage({ path: concepts, ...imageSizingContain(concepts, 6.45, 1.7, 5.7, 4.95) });

  chapter.sections.forEach((section, index) => {
    const slide = pptx.addSlide();
    addShell(slide, chapter, `分节 ${index + 1} · ${section.title}`, "每一节都按“白话说明 -> 孩子也能懂的类比 -> 交易场景”来拆解。");
    slide.addShape(SHAPE_TYPES.roundRect, {
      x: 1.0,
      y: 2.0,
      w: 5.4,
      h: 1.45,
      fill: { color: chapter.palette.soft, transparency: 3 },
      line: { color: chapter.palette.base, width: 1.2 },
      radius: 0.14,
    });
    slide.addText("书中意思", {
      x: 1.18,
      y: 2.12,
      w: 1.2,
      h: 0.25,
      fontFace: "Microsoft YaHei",
      fontSize: 16,
      bold: true,
      color: chapter.palette.deep,
    });
    slide.addText(section.summary, {
      x: 1.18,
      y: 2.45,
      w: 5.0,
      h: 0.8,
      fontFace: "Microsoft YaHei",
      fontSize: 15,
      color: "334E68",
      margin: 0.04,
    });
    slide.addShape(SHAPE_TYPES.roundRect, {
      x: 1.0,
      y: 3.75,
      w: 5.4,
      h: 1.15,
      fill: { color: "FFF8E8" },
      line: { color: chapter.palette.accent, width: 1.1 },
      radius: 0.14,
    });
    slide.addText("孩子版类比", {
      x: 1.18,
      y: 3.88,
      w: 1.35,
      h: 0.25,
      fontFace: "Microsoft YaHei",
      fontSize: 16,
      bold: true,
      color: chapter.palette.deep,
    });
    slide.addText(section.child, {
      x: 1.18,
      y: 4.17,
      w: 5.0,
      h: 0.55,
      fontFace: "Microsoft YaHei",
      fontSize: 14.5,
      color: "334E68",
      margin: 0.04,
    });
    slide.addShape(SHAPE_TYPES.roundRect, {
      x: 1.0,
      y: 5.2,
      w: 5.4,
      h: 1.1,
      fill: { color: "EFFCF6" },
      line: { color: "8AB17D", width: 1.1 },
      radius: 0.14,
    });
    slide.addText("交易场景", {
      x: 1.18,
      y: 5.33,
      w: 1.2,
      h: 0.25,
      fontFace: "Microsoft YaHei",
      fontSize: 16,
      bold: true,
      color: chapter.palette.deep,
    });
    slide.addText(section.market, {
      x: 1.18,
      y: 5.62,
      w: 5.0,
      h: 0.52,
      fontFace: "Microsoft YaHei",
      fontSize: 14.5,
      color: "334E68",
      margin: 0.04,
    });
    slide.addImage({ path: scene, ...imageSizingContain(scene, 6.8, 2.0, 5.0, 4.3) });
    slide.addText(`这一节的关键词：${chapter.sceneKeywords[index % chapter.sceneKeywords.length]}`, {
      x: 6.95,
      y: 6.05,
      w: 4.7,
      h: 0.25,
      fontFace: "Microsoft YaHei",
      fontSize: 13,
      color: chapter.palette.base,
      align: "center",
    });
  });

  const mythSlide = pptx.addSlide();
  addShell(mythSlide, chapter, "常见误区对照", "把最容易误解本章的地方直接翻译成“别再这样想”。");
  mythSlide.addImage({ path: myths, ...imageSizingContain(myths, 0.92, 2.05, 11.5, 4.65) });

  const analogySlide = pptx.addSlide();
  addShell(analogySlide, chapter, `儿童级类比 · ${chapter.childAnalogy.title}`, "用日常故事把抽象交易心理翻译成任何人都能感受到的经验。");
  analogySlide.addText(chapter.childAnalogy.story, {
    x: 1.02,
    y: 2.0,
    w: 5.2,
    h: 2.55,
    fontFace: "Microsoft YaHei",
    fontSize: 16,
    color: "334E68",
    margin: 0.06,
  });
  analogySlide.addShape(SHAPE_TYPES.roundRect, {
    x: 1.02,
    y: 5.05,
    w: 5.2,
    h: 1.35,
    fill: { color: chapter.palette.soft, transparency: 10 },
    line: { color: chapter.palette.base, width: 1.2 },
    radius: 0.14,
  });
  analogySlide.addText("这一则故事想帮助你记住：", {
    x: 1.18,
    y: 5.22,
    w: 2.4,
    h: 0.3,
    fontFace: "Microsoft YaHei",
    fontSize: 16,
    bold: true,
    color: chapter.palette.deep,
  });
  analogySlide.addText(chapter.childAnalogy.lesson, {
    x: 1.18,
    y: 5.55,
    w: 4.86,
    h: 0.58,
    fontFace: "Microsoft YaHei",
    fontSize: 15,
    color: "334E68",
    margin: 0.04,
  });
  analogySlide.addImage({ path: concepts, ...imageSizingContain(concepts, 6.55, 1.95, 5.35, 4.55) });

  const actionSlide = pptx.addSlide();
  addShell(actionSlide, chapter, "实战启示与行动清单", "把概念压缩成下一次开盘前就可以执行的动作。");
  actionSlide.addImage({ path: actions, ...imageSizingContain(actions, 0.95, 2.05, 11.45, 4.7) });

  const recapSlide = pptx.addSlide();
  addShell(recapSlide, chapter, "复盘清单", "读完整章后，至少要能把这三件事说清楚。");
  recapSlide.addText("1. 这章在讲什么？", {
    x: 1.05,
    y: 2.0,
    w: 3.5,
    h: 0.35,
    fontFace: "Microsoft YaHei",
    fontSize: 21,
    bold: true,
    color: chapter.palette.deep,
  });
  recapSlide.addText(chapter.elevator, {
    x: 1.15,
    y: 2.42,
    w: 11.0,
    h: 0.72,
    fontFace: "Microsoft YaHei",
    fontSize: 15,
    color: "334E68",
    margin: 0.04,
  });
  recapSlide.addText("2. 为什么很多人会做错？", {
    x: 1.05,
    y: 3.45,
    w: 4.0,
    h: 0.35,
    fontFace: "Microsoft YaHei",
    fontSize: 21,
    bold: true,
    color: chapter.palette.deep,
  });
  recapSlide.addText(chapter.misconceptions.map((item) => item.myth).join("；"), {
    x: 1.15,
    y: 3.87,
    w: 11.0,
    h: 0.84,
    fontFace: "Microsoft YaHei",
    fontSize: 15,
    color: "334E68",
    margin: 0.04,
  });
  recapSlide.addText("3. 读完后应该怎么做？", {
    x: 1.05,
    y: 5.05,
    w: 3.9,
    h: 0.35,
    fontFace: "Microsoft YaHei",
    fontSize: 21,
    bold: true,
    color: chapter.palette.deep,
  });
  bulletLines(recapSlide, chapter.checklist, { x: 1.15, y: 5.45, w: 11.1, h: 1.0 }, { fontSize: 14.5 });

  pptx._slides.forEach((slide) => {
    warnIfSlideHasOverlaps(slide, pptx, { muteContainment: true, ignoreDecorativeShapes: true });
    warnIfSlideElementsOutOfBounds(slide, pptx);
  });
  return pptx;
}

function writePptSource(chapter, slideCount) {
  const source = `"use strict";

const path = require("path");
const { buildSingleDeckFile } = require("../../scripts/generate_learning_pack");

buildSingleDeckFile("${chapter.id}", __dirname)
  .then((result) => {
    console.log(\`Generated ${chapter.id}-${chapter.titleZh} with \${result.slideCount} slides at \${result.output}\`);
  })
  .catch((error) => {
    console.error(error);
    process.exit(1);
  });

// Current slide count at generation time: ${slideCount}
`;
  fs.writeFileSync(path.join(PPT_DIR, `${chapter.id}-${chapter.titleZh}.js`), source, "utf8");
}

async function buildSingleDeckFile(chapterId, outDir = PPT_DIR) {
  const chapter = chapters.find((item) => item.id === chapterId);
  if (!chapter) {
    throw new Error(`Unknown chapter id: ${chapterId}`);
  }
  const pptx = createDeck(chapter);
  ensureDir(outDir);
  const output = path.join(outDir, `${chapter.id}-${chapter.titleZh}.pptx`);
  await pptx.writeFile({ fileName: output });
  return { chapter, output, slideCount: pptx._slides.length };
}

async function buildDecks() {
  const deckStats = [];
  for (const chapter of chapters) {
    const result = await buildSingleDeckFile(chapter.id, PPT_DIR);
    writePptSource(chapter, result.slideCount);
    deckStats.push({ id: chapter.id, title: chapter.titleZh, slides: result.slideCount, output: result.output });
  }
  fs.writeFileSync(path.join(WORK_DIR, "checks", "deck-stats.json"), JSON.stringify(deckStats, null, 2), "utf8");
}

async function main() {
  cleanOutput();
  buildAssets();
  buildReport();
  buildSummaries();
  buildReadme();
  await buildDecks();
}

if (require.main === module) {
  main().catch((error) => {
    console.error(error);
    process.exit(1);
  });
}

module.exports = {
  buildAssets,
  buildReport,
  buildSummaries,
  buildReadme,
  buildSingleDeckFile,
  buildDecks,
  main,
};
