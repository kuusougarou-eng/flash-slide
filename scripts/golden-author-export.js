"use strict";
// Build all 32 layouts from the golden sample data into debug/golden-native/authored/NN.pptx
// (one slide each) so PowerPoint can render them for pixel comparison with the reference.
//   node scripts/golden-author-export.js
const fs = require("fs"), path = require("path");
const author = require("../server/golden-native/author");
const { SAMPLES } = require("../server/golden-native/author/samples");
const out = path.resolve("debug/golden-native/authored");
fs.mkdirSync(out, { recursive: true });
const times = [];
for (const id of Object.keys(author.LAYOUTS).sort()) {
  const t = process.hrtime.bigint();
  const { buffer } = author.build(id, SAMPLES[id], { page: String(Number(id)) });
  times.push({ id, ms: Number(process.hrtime.bigint() - t) / 1e6 });
  fs.writeFileSync(path.join(out, id + ".pptx"), buffer);
}
fs.writeFileSync(path.join(out, "build-times.json"), JSON.stringify(times, null, 2));
console.log("built", times.length, "slides; total ms", times.reduce((n, t) => n + t.ms, 0).toFixed(1));
