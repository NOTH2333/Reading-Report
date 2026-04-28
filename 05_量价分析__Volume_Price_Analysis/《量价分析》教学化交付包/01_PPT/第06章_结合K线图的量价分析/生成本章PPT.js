"use strict";
const { buildSingle } = require("../../scripts/build_chapter_ppts");
buildSingle(6).catch((error) => {
  console.error(error);
  process.exit(1);
});
