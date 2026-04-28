"use strict";

const path = require("path");
const { buildSingleDeckFile } = require("../../scripts/generate_learning_pack");

buildSingleDeckFile("04", __dirname)
  .then((result) => {
    console.log(`Generated 04-持续一致性：一种思想状态 with ${result.slideCount} slides at ${result.output}`);
  })
  .catch((error) => {
    console.error(error);
    process.exit(1);
  });

// Current slide count at generation time: 9
