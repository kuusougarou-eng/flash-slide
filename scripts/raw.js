"use strict";
/**
 * デバッグ: サーバを経由せずに LLM を直接呼び、生の応答と normalizeSpec 後の body を並べて表示する。
 *   node scripts/raw.js <model> <case name in test/cases.json> [repeat]
 */
require("dotenv").config();
const fs = require("fs");
const path = require("path");
const llm = require("../server/llm");
const { buildMessages, extractJson } = require("../server/prompt");
const SlideLayout = require("../public/layout.js");

(async () => {
  const [model, name, rep] = process.argv.slice(2);
  const cases = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "test", "cases.json"), "utf8"));
  const c = cases.find((x) => x.name === name);
  if (!c) throw new Error("case not found: " + name);
  for (let i = 0; i < Number(rep || 1); i++) {
    const t0 = Date.now();
    const r = await llm.chat({ messages: buildMessages({ prompt: c.prompt, context: "", hint: "", maxSlides: 2 }), model, maxTokens: 3000, temperature: 0.2, jsonMode: true });
    const ms = Date.now() - t0;
    const out = path.join(__dirname, "..", "debug", `raw-${model.replace(/[^a-z0-9.-]/gi, "_")}-${name}-${Date.now()}.json`);
    fs.writeFileSync(out, r.content);
    let parsed = null,
      err = null;
    try {
      parsed = extractJson(r.content);
    } catch (e) {
      err = String(e.message);
    }
    const norm = parsed ? SlideLayout.normalizeSpec(parsed.slides ? parsed.slides[0] : parsed) : null;
    console.log(`#${i + 1} ${ms}ms out=${r.usage && r.usage.completion_tokens} saved=${out}`);
    console.log("  raw head:", r.content.slice(0, 700).replace(/\n/g, " "));
    if (err) console.log("  PARSE ERROR:", err);
    if (norm) console.log("  normalized body:", JSON.stringify(norm.body).slice(0, 500));
  }
})();
