"use strict";

const path = require("path");
const { buildBook } = require(path.resolve(__dirname, "../../../scripts/teaching_pack_ppt_builder.js"));

buildBook(path.resolve(__dirname, "../..")).catch((error) => {
  console.error(error);
  process.exit(1);
});
