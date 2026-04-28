"use strict";

const path = require("path");
const { buildDeck, writeDeck } = require(path.resolve(__dirname, "..", "00_项目说明", "生成脚本", "build_ppts.js"));

writeDeck("ch04").catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
