"use strict";
const { buildSingle } = require("../../scripts/build_chapter_ppts");
buildSingle(5).catch((error) => {
  console.error(error);
  process.exit(1);
});
