"use strict";
/**
 * 開発検証: DEV_TRIGGER=1 のサーバに対してトリガを順番に投入し、タスクペインの完了ログ([taskpane])を待って結果を集計する。
 *   node scripts/run-triggers.js reanalyze swimlane insight ...        … hint 名(モック)を順に実行
 *   node scripts/run-triggers.js --llm "プロンプト文" [--model X]        … 本物の LLM で 1 回
 *   node scripts/run-triggers.js --file test/cases.json [--model X]     … JSON の配列 [{name, prompt}] を本物の LLM で順に実行
 * 出力: 各ケースの mode / 描画 prims / 警告 / LLM ms / tokens。PNG は DEBUG_SNAPSHOT=1 なら debug/ に保存される。
 */
const fs = require("fs");
const path = require("path");
const https = require("https");

const ROOT = path.join(__dirname, "..");
const LOG = path.join(ROOT, "server.log");
const BASE = process.env.FS_BASE || "https://localhost:3455";
const agent = new https.Agent({ rejectUnauthorized: false });

function post(p, body) {
  return new Promise((resolve, reject) => {
    const u = new URL(BASE + p);
    const r = https.request(u, { method: "POST", headers: { "Content-Type": "application/json" }, agent }, (res) => {
      let d = "";
      res.on("data", (c) => (d += c));
      res.on("end", () => resolve(d));
    });
    r.on("error", reject);
    r.write(JSON.stringify(body));
    r.end();
  });
}
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
function logSize() {
  try {
    return fs.statSync(LOG).size;
  } catch (_) {
    return 0;
  }
}
function readFrom(offset) {
  const fd = fs.openSync(LOG, "r");
  const size = fs.fstatSync(fd).size;
  const buf = Buffer.alloc(Math.max(0, size - offset));
  if (buf.length) fs.readSync(fd, buf, 0, buf.length, offset);
  fs.closeSync(fd);
  return buf.toString("utf8");
}
async function waitFor(offset, re, timeoutMs) {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    const txt = readFrom(offset);
    const lines = txt.split("\n").filter((l) => re.test(l));
    if (lines.length) return { lines, txt };
    await sleep(500);
  }
  return { lines: [], txt: readFrom(offset) };
}

async function runOne(trigger, timeoutMs) {
  // TRIGGER_EXTRA(例 {"doc":"zeroshot"})は全経路で混ぜる(複数ペインが同じサーバをポーリングしているときの取り合い防止)
  if (process.env.TRIGGER_EXTRA) trigger = Object.assign({}, JSON.parse(process.env.TRIGGER_EXTRA), trigger);
  const off = logSize();
  await post("/api/debug/trigger", trigger);
  // 完了判定: 生成トリガは自分の tag の完了ログだけを待つ(同じサーバを使う別ペイン/別エージェントのログを拾わない)
  const esc = (x) => String(x).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const doneRe = trigger.prompt && trigger.tag ? new RegExp('\\[taskpane\\] \\{"(slide":"' + esc(trigger.tag) + '"|error")') : /\[taskpane\] \{"(slide|reanalyze|makeref|error)/;
  const done = await waitFor(off, doneRe, timeoutMs);
  const gen = done.txt.split("\n").filter((l) => /^\[generate\]/.test(l));
  const snaps = done.txt.split("\n").filter((l) => /^\[snapshot\]/.test(l)).map((l) => l.replace(/^\[snapshot\] /, ""));
  const errs = done.txt.split("\n").filter((l) => /\[taskpane\] \{"error"/.test(l) || /\[generate\]/.test(l) && /Error|error/.test(l));
  return { lines: done.lines, gen, snaps, errs };
}

(async () => {
  const args = process.argv.slice(2);
  let model = "";
  const mi = args.indexOf("--model");
  if (mi >= 0) {
    model = args[mi + 1];
    args.splice(mi, 2);
  }
  const results = [];
  if (args[0] === "--llm") {
    const prompt = args.slice(1).join(" ");
    const r = await runOne({ hint: "", prompt, mock: false, tag: "llm", model }, 180000);
    results.push({ name: "llm", ...r });
  } else if (args[0] === "--file") {
    const cases = JSON.parse(fs.readFileSync(args[1], "utf8"));
    for (const c of cases) {
      process.stdout.write(`▶ ${c.name} ... `);
      const r = await runOne({ hint: c.hint || "", prompt: c.prompt, mock: false, tag: (model ? model.replace(/[^a-z0-9.-]/gi, "_") + "-" : "") + c.name, model }, 240000);
      const g = r.gen[r.gen.length - 1] || "";
      console.log(r.lines.length ? "ok " + g.replace(/^\[generate\] /, "") : "TIMEOUT");
      results.push({ name: c.name, ...r });
      await sleep(800);
    }
  } else {
    for (const hint of args) {
      process.stdout.write(`▶ ${hint} ... `);
      if (hint === "reload") {
        await post("/api/debug/trigger", Object.assign({}, process.env.TRIGGER_EXTRA ? JSON.parse(process.env.TRIGGER_EXTRA) : {}, { hint: "reload" }));
        await sleep(6000);
        console.log("sent");
        continue;
      }
      const r = await runOne({ hint, prompt: "", mock: true, tag: hint }, 60000);
      console.log(r.lines.length ? "ok" : "TIMEOUT");
      results.push({ name: hint, ...r });
      await sleep(800);
    }
  }
  const out = path.join(ROOT, "debug", `triggers-${Date.now()}.json`);
  fs.writeFileSync(out, JSON.stringify(results, null, 2));
  for (const r of results) {
    console.log(`\n== ${r.name}`);
    for (const l of r.lines) console.log("  " + l.slice(0, 400));
    for (const g of r.gen) console.log("  " + g);
    for (const s of r.snaps) console.log("  png: " + s);
    for (const e of r.errs) console.log("  ERR: " + e.slice(0, 300));
  }
  console.log("\nsaved " + out);
})();
