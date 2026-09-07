"use strict";
/** 検証: 既定モデルがスライド画像を読んで書き起こせるか(Remix の実現可能性)。node scripts/_vision-probe.js <png> */
const fs = require("fs");
const path = require("path");
require("dotenv").config();
const llm = require(path.join(__dirname, "..", "server/llm"));
(async () => {
  const png = fs.readFileSync(process.argv[2]).toString("base64");
  const t0 = Date.now();
  const r = await llm.chat({
    messages: [
      { role: "system", content: "あなたはスライドの書き起こし係。画像のスライドの内容を、構造(タイトル/リード/見出し/箇条書き/表のセル)を保ったまま日本語のプレーンテキストに書き起こす。装飾語は足さない。" },
      { role: "user", content: [{ type: "text", text: "このスライドを書き起こして。" }, { type: "image_url", image_url: { url: "data:image/png;base64," + png } }] },
    ],
    maxTokens: 1200,
    temperature: 0.1,
  });
  fs.writeFileSync(path.join(__dirname, "..", "debug", "vision-probe.txt"), `model=${r.model} ms=${Date.now() - t0} usage=${JSON.stringify(r.usage)}\n\n${r.content}`, "utf8");
  console.log("ok", r.model, Date.now() - t0, "ms");
})().catch((e) => {
  console.log("ERROR", e.message);
  process.exit(1);
});
