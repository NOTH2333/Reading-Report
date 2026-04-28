"use strict";

const path = require("path");
const ROOT = path.resolve(__dirname, "..", "..");
const deckData = require(path.join(ROOT, "90_共用素材与底稿", "chapter_decks.json"));
const { renderChapterDeck } = require(path.join(ROOT, "scripts", "ppt_renderer_common.js"));

const chapter = deckData.chapters.find((item) => item.id === 3);

renderChapterDeck(chapter, ROOT)
  .then((file) => console.log("generated:", file))
  .catch((error) => {
    console.error(error);
    process.exit(1);
  });
