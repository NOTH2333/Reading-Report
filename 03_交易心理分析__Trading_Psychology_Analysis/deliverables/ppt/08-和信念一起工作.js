"use strict";

const path = require("path");
const { buildSingleDeckFile } = require("../../scripts/generate_learning_pack");

buildSingleDeckFile("08", __dirname)
  .then((result) => {
    console.log(`Generated 08-和信念一起工作 with ${result.slideCount} slides at ${result.output}`);
  })
  .catch((error) => {
    console.error(error);
    process.exit(1);
  });

// Current slide count at generation time: 10
