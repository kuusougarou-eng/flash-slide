"use strict";
/**
 * 品質ゲート用の fixture を、本物の LLM から 1 回の推論で取ってくる。
 *   node scripts/capture-specs.js                 … 全ケース
 *   node scripts/capture-specs.js proposal-m …    … 名前を指定
 *   FS_BASE=https://localhost:3455 node scripts/capture-specs.js --model <name>
 *
 * サーバを起動しておくこと(既定は https://localhost:3455)。既定は単回推論
 * (MULTI_PASS を立てていない状態)。保存先は test/fixtures/specs/<name>.json。
 * LLM は毎回同じ答えを返さないので、fixture を更新したときは中身を目で見てから commit する。
 */
const fs = require("fs");
const path = require("path");
const https = require("https");
const http = require("http");

const ROOT = path.join(__dirname, "..");
const OUT = path.join(ROOT, "test", "fixtures", "specs");
const BASE = process.env.FS_BASE || "https://localhost:3455";
const agent = BASE.startsWith("https") ? new https.Agent({ rejectUnauthorized: false }) : undefined;

function post(p, body) {
  return new Promise((resolve, reject) => {
    const u = new URL(BASE + p);
    const mod = u.protocol === "https:" ? https : http;
    const req = mod.request(u, { method: "POST", headers: { "Content-Type": "application/json" }, agent }, (res) => {
      let d = "";
      res.on("data", (c) => (d += c));
      res.on("end", () => {
        try {
          resolve(JSON.parse(d));
        } catch (e) {
          reject(new Error("bad JSON from server: " + d.slice(0, 200)));
        }
      });
    });
    req.on("error", reject);
    req.write(JSON.stringify(body));
    req.end();
  });
}

(async () => {
  const args = process.argv.slice(2);
  let model = "";
  const mi = args.indexOf("--model");
  if (mi >= 0) {
    model = args[mi + 1];
    args.splice(mi, 2);
  }
  const inputs = require(path.join(ROOT, "test", "fixtures", "inputs.json"));
  const targets = args.length ? inputs.filter((c) => args.includes(c.name)) : inputs;
  fs.mkdirSync(OUT, { recursive: true });

  for (const c of targets) {
    process.stdout.write(`▶ ${c.name} (${c.prompt.length} 字) ... `);
    const t0 = Date.now();
    let res;
    try {
      res = await post("/api/generate", { prompt: c.prompt, hint: "", model, mock: false, maxSlides: 2 });
    } catch (e) {
      console.log("ERROR " + e.message);
      continue;
    }
    if (res.error) {
      console.log("ERROR " + res.error);
      continue;
    }
    const ms = Date.now() - t0;
    const out = {
      name: c.name,
      genre: c.genre,
      prompt: c.prompt,
      capturedAt: new Date().toISOString().slice(0, 10),
      model: res.model,
      roundTrips: res.roundTrips || 1,
      serverMs: res.ms,
      response: res.sections ? { sections: res.sections } : { slides: res.slides },
    };
    fs.writeFileSync(path.join(OUT, c.name + ".json"), JSON.stringify(out, null, 1) + "\n", "utf8");
    const shape = res.sections ? `sections=${res.sections.length}` : `slides=${(res.slides || []).length}`;
    console.log(`${shape} x${out.roundTrips} ${ms}ms`);
  }
})();
