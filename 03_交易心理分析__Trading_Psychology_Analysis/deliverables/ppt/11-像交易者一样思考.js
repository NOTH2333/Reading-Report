"use strict";

const path = require("path");
const { buildSingleDeckFile } = require("../../scripts/generate_learning_pack");

buildSingleDeckFile("11", __dirname)
  .then((result) => {
    console.log(`Generated 11-像交易者一样思考 with ${result.slideCount} slides at ${result.output}`);
  })
  .catch((error) => {
    console.error(error);
    process.exit(1);
  });

// Current slide count at generation time: 13
