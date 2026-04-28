"use strict";

const path = require("path");
const { buildSingleDeckFile } = require("../../scripts/generate_learning_pack");

buildSingleDeckFile("01", __dirname)
  .then((result) => {
    console.log(`Generated 01-成功之路：基本面、技术面或思想分析？ with ${result.slideCount} slides at ${result.output}`);
  })
  .catch((error) => {
    console.error(error);
    process.exit(1);
  });

// Current slide count at generation time: 9
