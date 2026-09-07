"use strict";
/**
 * モデル選定ベンチ: 実プロンプト(test/cases.json)× モデルで /api/generate を叩き、速度・トークン・構造の妥当性・レイアウト品質を集計する。
 *   node scripts/bench.js --models <model-a>,<model-b> [--cases test/cases.json] [--parallel 3] [--repeat 1]
 * 出力: debug/bench-<ts>.json(全 spec 保存)と、標準出力にモデル別サマリ(Markdown 表)。
 * 品質判定は決定論: JSON 妥当性 / 空でない本文 / 960×540 でのはみ出し・縮小ゼロ / タイトル規約 / ケースごとの期待構造(正規表現)。
 */
const fs = require("fs");
const path = require("path");
const https = require("https");

const ROOT = path.join(__dirname, "..");
const SlideLayout = require(path.join(ROOT, "public", "layout.js"));
const BASE = process.env.FS_BASE || "https://localhost:3455";
const agent = new https.Agent({ rejectUnauthorized: false });

// 検証デッキ(プロジェクトアプローチ.pptx)相当の参照プロファイル。clone 時の本文領域でレイアウト品質を測る
const PROFILE = {
  title: { x: 32, y: 32, w: 896, h: 40, fontSize: 32, bold: false, color: "#000000", align: "left", reuse: true },
  lead: { x: 32, y: 77, w: 896, h: 62, fontSize: 24, color: "#000000", align: "left", reuse: true },
  body: { x: 32, y: 149, w: 896, h: 357 },
  footnote: null,
  fontName: "Arial",
  accent: "#D04A02",
  bodyFontSize: 14,
  headFontSize: 14,
};

// ケースごとの期待構造(spec JSON 文字列に対する正規表現。どれか 1 つ一致で OK)
const EXPECT = {
  // composition v1(panelCount / 主図 1 つ)での期待構造。判定は正規化後の spec(JSON 文字列)に対する正規表現
  approach: [/"type":"(table|sequence)"/],
  transport: [/"type":"table"/, /"panelCount":[23]/],
  "compare-vendors": [/"type":"table"[\s\S]*"colHeaders":\[[^\]]*[ABC]\s?社/],
  interview: [/"panelCount":[23]/],
  "sales-decline": [/"panelCount":2/, /"type":"(bars|stacked|table)"/],
  schedule: [/"type":"gantt"/],
  "kpi-trend": [/"type":"line"/],
  quadrant: [/"axes"/, /"type":"table"/],
  "cause-effect": [/"type":"sequence"/, /"panelCount":2/],
  steps: [/"type":"sequence"/],
  "org-chart": [/"type":"org"/],
  "gantt-detail": [/"type":"gantt"/],
  "matrix-4x5": [/"type":"table"[\s\S]*"colHeaders":\[[^\]]*リスク/],
};

function post(p, body) {
  return new Promise((resolve, reject) => {
    const u = new URL(BASE + p);
    const r = https.request(u, { method: "POST", headers: { "Content-Type": "application/json" }, agent }, (res) => {
      let d = "";
      res.on("data", (c) => (d += c));
      res.on("end", () => {
        try {
          resolve({ status: res.statusCode, json: JSON.parse(d) });
        } catch (e) {
          resolve({ status: res.statusCode, json: null, text: d });
        }
      });
    });
    r.on("error", reject);
    r.write(JSON.stringify(body));
    r.end();
  });
}

function judge(name, gen) {
  const out = { ok: false, reasons: [], warnings: [], shrunk: 0, leaves: 0, types: {}, titleLen: 0, expect: null };
  if (!gen || !gen.slides || !gen.slides.length) {
    out.reasons.push(gen && gen.sections ? "sections(分割提案)" : "no slides");
    return out;
  }
  const spec = gen.slides[0];
  const json = JSON.stringify(spec);
  out.titleLen = (spec.title || "").length;
  if (out.titleLen < 12 || out.titleLen > 60) out.reasons.push("title length " + out.titleLen);
  if (/(です|ます)$/.test(spec.title || "")) out.reasons.push("title 敬体");
  (function walk(n) {
    if (!n || typeof n !== "object") return;
    if (!n.type && (n.rows || n.cols)) return (n.rows || n.cols).forEach(walk);
    out.leaves++;
    out.types[n.type] = (out.types[n.type] || 0) + 1;
  })(spec.body);
  if (out.leaves === 0) out.reasons.push("empty body");
  for (const prof of [null, PROFILE]) {
    const lay = SlideLayout.layout(spec, { width: 960, height: 540, profile: prof });
    const bad = lay.warnings.filter((w) => /overflow|failed/.test(w));
    out.warnings.push(...bad.map((w) => (prof ? "[clone] " : "[layout] ") + w));
    out.shrunk += lay.prims.filter((p) => p.shrunk).length;
    for (const p of lay.prims) {
      if (p.kind !== "rect" || p.rotation) continue;
      if (p.x < -0.5 || p.y < -0.5 || p.x + p.w > 960.5 || p.y + p.h > 540.5) out.reasons.push("out of slide");
    }
  }
  if (out.warnings.length) out.reasons.push("layout warnings " + out.warnings.length);
  if (out.shrunk) out.reasons.push("shrunk " + out.shrunk);
  const ex = EXPECT[name];
  if (ex) {
    out.expect = ex.some((re) => re.test(json));
    if (!out.expect) out.reasons.push("structure mismatch");
  }
  out.ok = out.reasons.length === 0;
  return out;
}

(async () => {
  const args = process.argv.slice(2);
  const opt = (k, d) => {
    const i = args.indexOf(k);
    return i >= 0 ? args[i + 1] : d;
  };
  // --rejudge <bench.json>: 保存済みの spec をエンジンの現行版で再判定する(LLM を呼ばない)
  if (opt("--rejudge", "")) {
    const prev = JSON.parse(fs.readFileSync(opt("--rejudge", ""), "utf8"));
    const rows = prev.map((r) => {
      if (!r.spec || r.spec.sections) return r;
      const j = judge(r.case, { slides: [r.spec] });
      return Object.assign({}, r, { ok: j.ok, reasons: j.reasons, expect: j.expect });
    });
    const models = Array.from(new Set(rows.map((r) => r.model)));
    console.log("| model | cases | OK | 構造一致 | 平均 ms | NG 理由 |");
    console.log("|---|---|---|---|---|---|");
    for (const model of models) {
      const rs = rows.filter((r) => r.model === model);
      const reasons = {};
      rs.forEach((r) => r.reasons.forEach((x) => (reasons[x.replace(/\d+/g, "n")] = (reasons[x.replace(/\d+/g, "n")] || 0) + 1)));
      console.log(`| ${model} | ${rs.length} | ${rs.filter((r) => r.ok).length} | ${rs.filter((r) => r.expect === true).length}/${rs.filter((r) => r.expect != null).length} | ${Math.round(rs.reduce((a, r) => a + r.ms, 0) / rs.length)} | ${Object.entries(reasons).map(([k, v]) => k + "×" + v).join("; ")} |`);
    }
    return;
  }
  const models = String(opt("--models", "")).split(",").map((s) => s.trim()).filter(Boolean);
  const cases = JSON.parse(fs.readFileSync(opt("--cases", path.join(ROOT, "test", "cases.json")), "utf8"));
  const parallel = Number(opt("--parallel", 3));
  const repeat = Number(opt("--repeat", 1));
  if (!models.length) {
    console.error("--models を指定してください");
    process.exit(1);
  }
  const results = [];
  const queue = models.slice();
  async function worker() {
    while (queue.length) {
      const model = queue.shift();
      for (let rep = 0; rep < repeat; rep++) {
        for (const c of cases) {
          const t0 = Date.now();
          let r;
          try {
            r = await post("/api/generate", { prompt: c.prompt, hint: c.hint || "", model, mock: false, maxSlides: 2 });
          } catch (e) {
            r = { status: 0, json: null, text: String(e) };
          }
          const ms = Date.now() - t0;
          const gen = r.json;
          const err = r.status !== 200 ? (gen && gen.error) || r.text || "http " + r.status : null;
          const j = err ? { ok: false, reasons: ["error: " + String(err).slice(0, 120)], warnings: [], shrunk: 0, leaves: 0, types: {} } : judge(c.name, gen);
          const usage = (gen && gen.usage) || {};
          results.push({ model, case: c.name, rep, ms, llmMs: gen && gen.ms, in: usage.prompt_tokens, out: usage.completion_tokens, reasoning: usage.completion_tokens_details && usage.completion_tokens_details.reasoning_tokens, ok: j.ok, reasons: j.reasons, types: j.types, leaves: j.leaves, titleLen: j.titleLen, expect: j.expect, spec: gen && gen.slides ? gen.slides[0] : gen && gen.sections ? { sections: gen.sections } : null, error: err });
          console.log(`${model} | ${c.name} | ${ms}ms | out=${usage.completion_tokens || "?"} | ${j.ok ? "OK" : "NG: " + j.reasons.join(", ")}`);
        }
      }
    }
  }
  await Promise.all(Array.from({ length: Math.min(parallel, models.length) }, worker));

  const out = path.join(ROOT, "debug", `bench-${Date.now()}.json`);
  fs.writeFileSync(out, JSON.stringify(results, null, 2));
  console.log("\n| model | cases | OK | 構造一致 | 平均 ms | 中央値 ms | 最大 ms | 平均 out tok | NG 理由 |");
  console.log("|---|---|---|---|---|---|---|---|---|");
  for (const model of models) {
    const rs = results.filter((r) => r.model === model);
    const ok = rs.filter((r) => r.ok).length;
    const ex = rs.filter((r) => r.expect === true).length;
    const exN = rs.filter((r) => r.expect != null).length;
    const mss = rs.map((r) => r.ms).sort((a, b) => a - b);
    const avg = Math.round(mss.reduce((a, b) => a + b, 0) / Math.max(1, mss.length));
    const med = mss[Math.floor(mss.length / 2)] || 0;
    const outTok = rs.map((r) => r.out || 0);
    const avgOut = Math.round(outTok.reduce((a, b) => a + b, 0) / Math.max(1, outTok.length));
    const reasons = {};
    rs.forEach((r) => r.reasons.forEach((x) => (reasons[x.replace(/\d+/g, "n")] = (reasons[x.replace(/\d+/g, "n")] || 0) + 1)));
    console.log(`| ${model} | ${rs.length} | ${ok} | ${ex}/${exN} | ${avg} | ${med} | ${mss[mss.length - 1] || 0} | ${avgOut} | ${Object.entries(reasons).map(([k, v]) => k + "×" + v).join("; ")} |`);
  }
  console.log("\nsaved " + out);
})();
