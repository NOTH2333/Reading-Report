"use strict";
const { buildSingle } = require("../../scripts/build_chapter_ppts");
buildSingle(1).catch((error) => {
  console.error(error);
  process.exit(1);
});
