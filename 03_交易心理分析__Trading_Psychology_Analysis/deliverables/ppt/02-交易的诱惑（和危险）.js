"use strict";

const path = require("path");
const { buildSingleDeckFile } = require("../../scripts/generate_learning_pack");

buildSingleDeckFile("02", __dirname)
  .then((result) => {
    console.log(`Generated 02-交易的诱惑（和危险） with ${result.slideCount} slides at ${result.output}`);
  })
  .catch((error) => {
    console.error(error);
    process.exit(1);
  });

// Current slide count at generation time: 13
