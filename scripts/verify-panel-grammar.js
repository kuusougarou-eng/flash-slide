"use strict";
// Manual QA with the configured LLM. Saves raw responses and deterministic layouts.
require("dotenv").config({ quiet: true });
const fs = require("fs");
const path = require("path");
const llm = require("../server/llm");
const { buildMessages, extractJson, shapeResponse } = require("../server/prompt");
const L = require("../public/layout");
const cases = require("../test/cases.json");
const selected = process.argv.slice(2);
const names = selected.length ? selected : ["sales-decline", "steps", "schedule", "compare-vendors", "matrix-4x5"];
const dir = path.join(__dirname, "../debug/panel-grammar");
fs.mkdirSync(dir, { recursive: true });
(async () => {
  const result = await Promise.allSettled(names.map(async (name) => {
    const test = cases.find((c) => c.name === name);
    if (!test) throw new Error("Unknown case: " + name);
    const response = await llm.chat({ messages: buildMessages({ prompt: test.prompt, maxSlides: 1 }), maxTokens: 4000, temperature: 0.2, jsonMode: true });
    const raw = extractJson(response.content);
    fs.writeFileSync(path.join(dir, name + "-raw.json"), JSON.stringify(raw, null, 2));
    const shaped = shapeResponse(raw, L.normalizeGeneratedSpec, 1);
    const layouts = shaped.slides.map((s) => L.layout(s, { width: 960, height: 540, palette: { accent: "none" } }));
    fs.writeFileSync(path.join(dir, name + "-generated.json"), JSON.stringify({ input: test.prompt, raw, shaped, layouts }, null, 2));
    console.log(name, shaped.slides.map((s) => `${s.panelCount} panels / ${s.body.type || "text"}`), layouts.map((l) => ({ fonts: l.fonts, warnings: l.warnings })));
    return name;
  }));
  for (let i = 0; i < result.length; i++) if (result[i].status === "rejected") {
    console.error(names[i], result[i].reason.message);
    process.exitCode = 1;
  }
})();
