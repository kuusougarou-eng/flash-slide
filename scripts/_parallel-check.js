"use strict";
/** 検証: N 案を並列生成したときの実測待ち時間・失敗率・構成の重複。結果は debug/parallel-check.json */
const fs = require("fs");
const path = require("path");
const ROOT = path.join(__dirname, "..");
require("dotenv").config();
const llm = require(path.join(ROOT, "server/llm"));
const { buildMessages, extractJson, shapeResponse } = require(path.join(ROOT, "server/prompt"));
const SlideLayout = require(path.join(ROOT, "public/layout.js"));

const PROMPTS = {
  transit: [
    "地域交通の再編について、経営会議向けに",
    "・利用者数は2019年比で22%減。通学利用は41%減、65歳以上は9%増",
    "・87%が運転手不足と回答。平均年齢54.2歳、今後5年で3割が退職見込み",
    "・212路線のうち168路線が赤字。自治体補助は年38億円で5年前の1.6倍",
    "・住民の困りごとは 本数が少ない63%、最寄り停留所まで遠い38%、乗り継ぎが不便29%",
    "・隣県A市はデマンド交通へ切替え、運行コスト27%削減、満足度4.1/5",
  ].join("\n"),
  margin: [
    "2026 上期レビュー",
    "・受注は前年比 +18%(大型 2 件が牽引。うち 1 件は一過性)",
    "・粗利率は ▲2.0pt。値引き拡大(▲1.2pt)と物流費の上昇(▲0.8pt)が要因",
    "・下期は値引き決裁基準の統一と配送ルート見直しで +1.5pt を狙う",
    "・リスク: 大型案件の反動減。新規パイプラインは前年比 92%",
  ].join("\n"),
};

function sig(node) {
  if (!node || typeof node !== "object") return "";
  if (Array.isArray(node.rows) && !node.type) return "rows[" + node.rows.map(sig).join(",") + "]";
  if (Array.isArray(node.cols)) return "cols[" + node.cols.map(sig).join(",") + "]";
  if (node.type === "table") return `table${(node.rows || []).length}x${(node.colHeaders || []).length}${node.headShape ? "/" + node.headShape : ""}${node.axes ? "/axes" : ""}`;
  if (node.type === "cell") return `cell(${node.head ? "h+" : ""}${node.items ? "items" + node.items.length : node.text ? "text" : "head"}${node.shape ? "/" + node.shape : ""}${node.fill ? "/" + node.fill : ""})`;
  return node.type || "?";
}

async function one(prompt) {
  const t0 = Date.now();
  const r = await llm.chat({ messages: buildMessages({ prompt, context: "", hint: "", maxSlides: 2 }), maxTokens: 3000, temperature: 0.2, jsonMode: true });
  const shaped = shapeResponse(extractJson(r.content), SlideLayout.normalizeSpec, 2);
  const s = (shaped.slides || [])[0] || {};
  return { ms: Date.now() - t0, in: (r.usage || {}).prompt_tokens, out: (r.usage || {}).completion_tokens, slides: (shaped.slides || []).length, title: s.title || "", sig: sig(s.body) };
}

async function batch(prompt, n) {
  const t0 = Date.now();
  const rs = await Promise.all(
    Array.from({ length: n }, () => one(prompt).catch((e) => ({ error: String(e.message || e) })))
  );
  return { n, wallMs: Date.now() - t0, results: rs };
}

(async () => {
  const out = [];
  for (const [name, prompt] of Object.entries(PROMPTS)) {
    for (const n of [1, 4]) {
      for (let rep = 0; rep < 2; rep++) {
        const b = await batch(prompt, n);
        const ok = b.results.filter((r) => !r.error);
        const uniq = new Set(ok.map((r) => r.sig)).size;
        const each = ok.map((r) => r.ms);
        out.push({ case: name, n, rep: rep + 1, wallMs: b.wallMs, eachMs: each, errors: b.results.length - ok.length, uniqueSigs: uniq, titles: ok.map((r) => r.title), sigs: ok.map((r) => r.sig), tokens: ok.map((r) => `${r.in}/${r.out}`) });
        console.log(`${name} n=${n} rep${rep + 1}: wall=${b.wallMs}ms each=[${each.join(",")}] err=${b.results.length - ok.length} uniq=${uniq}/${ok.length}`);
      }
    }
  }
  fs.writeFileSync(path.join(ROOT, "debug", "parallel-check.json"), JSON.stringify(out, null, 2), "utf8");
  console.log("written debug/parallel-check.json");
})();
