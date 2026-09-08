"use strict";
// Run the one-inference golden generator on novel inputs (real LLM from .env) and
// save the compiled slides + plans for PowerPoint rendering and review.
//   node scripts/golden-author-llm.js [inputs.json] [--only id1,id2] [--model m] [--layout NN]
require("dotenv").config();
const fs = require("fs"), path = require("path");
const gen = require("../server/golden-native/author/generate");
const args = process.argv.slice(2);
const opt = (k, d) => (args.includes(k) ? args[args.indexOf(k) + 1] : d);
const file = args.find((a) => a.endsWith(".json")) || "test/golden-author/inputs.json";
const only = opt("--only", "").split(",").filter(Boolean);
const inputs = JSON.parse(fs.readFileSync(file, "utf8")).filter((x) => !only.length || only.includes(x.id));
const out = path.resolve("debug/golden-native/llm");
fs.mkdirSync(out, { recursive: true });
(async () => {
  const summary = [];
  for (const it of inputs) {
    const t = Date.now();
    try {
      const r = await gen.generate(it.text, { model: opt("--model"), layout: opt("--layout", it.layout || "") });
      fs.writeFileSync(path.join(out, it.id + ".pptx"), r.buffer);
      fs.writeFileSync(path.join(out, it.id + ".json"), JSON.stringify({ layout: r.layout, plan: r.plan, warnings: r.warnings, inferenceMs: r.inferenceMs, compileMs: r.compileMs, usage: r.usage, model: r.model }, null, 2));
      summary.push({ id: it.id, layout: r.layout, inferenceMs: r.inferenceMs, compileMs: r.compileMs, totalMs: Date.now() - t, promptTokens: r.usage?.prompt_tokens, completionTokens: r.usage?.completion_tokens, warnings: r.warnings.length });
      console.log(it.id, "→ layout", r.layout, "inference", r.inferenceMs, "ms; compile", r.compileMs, "ms; warnings", r.warnings.length);
    } catch (e) {
      summary.push({ id: it.id, error: e.message, totalMs: Date.now() - t });
      console.log(it.id, "ERROR", e.message);
      if (e.plan) fs.writeFileSync(path.join(out, it.id + ".error.json"), JSON.stringify(e.plan, null, 2));
    }
  }
  fs.writeFileSync(path.join(out, "summary.json"), JSON.stringify(summary, null, 2));
})();
