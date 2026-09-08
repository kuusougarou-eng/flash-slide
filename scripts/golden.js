/**
 * ゴールデンスライドの描画・書き出し。
 *   node scripts/golden.js [--deck golden] [--only 01-exec-summary,08-roadmap-gantt] [--catalog docs/golden/catalog.json]
 *
 * docs/golden/catalog.json の各 spec を、開いている参照デッキ(既定 debug/golden.pptx = ユーザーのマスタ)の
 * タスクペインに renderspec トリガで送り、実機描画 → COM で PNG(debug/golden/<id>.png)に書き出す。
 * 最後に一覧シート debug/golden/sheet.png を作る。LLM は使わない(手書きの正解を、そのまま描く)。
 *
 * 前提: DEV_TRIGGER=1 のサーバ、対象デッキのペインが doc 付きでポーリングしていること。
 */
const fs = require("fs");
const path = require("path");
const https = require("https");
const { execFileSync } = require("child_process");

const ROOT = path.join(__dirname, "..");
const args = process.argv.slice(2);
const opt = (k, d) => (args.includes(k) ? args[args.indexOf(k) + 1] : d);
const DECK = opt("--deck", "golden");
const ONLY = opt("--only", "").split(",").filter(Boolean);
const CATALOG = path.join(ROOT, opt("--catalog", "docs/golden/catalog.json"));
const PORT = process.env.PORT || 3455;
const BASE = `https://localhost:${PORT}`;
const LOG = path.join(ROOT, "server.log");
const OUT = path.join(ROOT, "debug", "golden");
const agent = new https.Agent({ rejectUnauthorized: false });

function post(p, body) {
  return new Promise((resolve, reject) => {
    const r = https.request(new URL(BASE + p), { method: "POST", headers: { "Content-Type": "application/json" }, agent }, (res) => {
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
const logSize = () => (fs.existsSync(LOG) ? fs.statSync(LOG).size : 0);
async function waitLog(off, re, timeoutMs) {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    const txt = fs.existsSync(LOG) ? fs.readFileSync(LOG, "utf8").slice(off) : "";
    const m = txt.match(re);
    if (m) return { line: m[0], txt };
    await sleep(400);
  }
  return null;
}

(async () => {
  const cat = JSON.parse(fs.readFileSync(CATALOG, "utf8"));
  fs.mkdirSync(OUT, { recursive: true });
  const slides = cat.slides.filter((g) => !ONLY.length || ONLY.includes(g.id));
  const results = [];
  for (const g of slides) {
    // 参照は常にデッキの 1 枚目(マスタの見本)。生成物が参照にならないように毎回選び直す
    try {
      execFileSync("powershell.exe", ["-NoProfile", "-File", path.join(ROOT, "scripts", "ppt-activate.ps1"), DECK, "1"], { env: { ...process.env, PPT_NO_ACTIVATE: "1" }, stdio: "ignore" });
    } catch (_) {}
    const tag = "golden-" + g.id;
    const off = logSize();
    await post("/api/debug/trigger", { hint: "renderspec", doc: DECK, tag, specs: [g.spec], accent: "none" });
    const done = await waitLog(off, new RegExp('\\[taskpane\\] \\{"renderedSpecs":"' + tag.replace(/[-]/g, "\\-") + '"[^\\n]*'), 90000);
    if (!done) {
      console.log(`TIMEOUT ${g.id}`);
      results.push({ id: g.id, ok: false });
      continue;
    }
    const ids = [...done.line.matchAll(/"slideId":"([^"]+)"/g)].map((m) => m[1]);
    const warns = [...done.line.matchAll(/"warnings":\[([^\]]*)\]/g)].map((m) => m[1]).filter(Boolean);
    const pngs = [];
    ids.forEach((id, i) => {
      const png = path.join(OUT, `${g.id}${ids.length > 1 ? "-" + (i + 1) : ""}.png`);
      try {
        execFileSync("powershell.exe", ["-NoProfile", "-File", path.join(ROOT, "scripts", "ppt-export-by-id.ps1"), DECK, id, png], { encoding: "utf8" });
        pngs.push(png);
      } catch (e) {
        console.log(`  export failed: ${(e.stdout || e.message).toString().slice(0, 120)}`);
      }
    });
    console.log(`${ids.length ? "ok" : "NG"} ${g.id.padEnd(20)} ${g.pattern}${warns.length ? "  警告: " + warns.join(" | ").slice(0, 160) : ""}`);
    results.push({ id: g.id, pattern: g.pattern, ok: !!ids.length, slideIds: ids, pngs, warnings: warns });
    await sleep(500);
  }
  fs.writeFileSync(path.join(OUT, "results.json"), JSON.stringify(results, null, 1));
  console.log(`\n${results.filter((r) => r.ok).length}/${results.length} rendered → ${OUT}`);
})();
