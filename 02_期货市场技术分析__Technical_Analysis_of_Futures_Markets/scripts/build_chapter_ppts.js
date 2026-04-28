"use strict";

const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const deckData = require(path.join(ROOT, "90_共用素材与底稿", "chapter_decks.json"));
const { renderChapterDeck } = require(path.join(__dirname, "ppt_renderer_common.js"));

function writeChapterSource(chapter) {
  const chapterDir = path.join(ROOT, "02_章节PPT", chapter.dirName);
  const sourcePath = path.join(chapterDir, chapter.sourceFile);
  const wrapper = `"use strict";

const path = require("path");
const ROOT = path.resolve(__dirname, "..", "..");
const deckData = require(path.join(ROOT, "90_共用素材与底稿", "chapter_decks.json"));
const { renderChapterDeck } = require(path.join(ROOT, "scripts", "ppt_renderer_common.js"));

const chapter = deckData.chapters.find((item) => item.id === ${chapter.id});

renderChapterDeck(chapter, ROOT)
  .then((file) => console.log("generated:", file))
  .catch((error) => {
    console.error(error);
    process.exit(1);
  });
`;
  fs.writeFileSync(sourcePath, wrapper, "utf8");
}

async function main() {
  const arg = process.argv[2] ? Number(process.argv[2]) : null;
  const chapters = arg
    ? deckData.chapters.filter((chapter) => chapter.id === arg)
    : deckData.chapters;
  if (chapters.length === 0) {
    throw new Error(`No chapter matched argument: ${process.argv[2]}`);
  }
  for (const chapter of chapters) {
    writeChapterSource(chapter);
    const out = await renderChapterDeck(chapter, ROOT);
    console.log(out);
  }
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
