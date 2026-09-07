/**
 * 入力の取りこぼし計測。
 *   node scripts/coverage.js [--cases test/cases.json] [--model X] [--repeat 1] [--quiet]
 *
 * 入力プロンプトから「落としてはいけない原子」(数値・割合・期間・固有名詞らしき語)を機械的に抜き、
 * 生成された spec にそれが現れるかを数える。コンテキストの追加は許されないが、与えられた内容は
 * なるべく漏れなく載せる、という要件を数字で見るための道具。
 *
 * 出力: モデル別・ケース別の被覆率と、落ちた原子の一覧(debug/coverage-<ts>.json に全文)。
 */
const fs = require("fs");
const path = require("path");
const https = require("https");
const http = require("http");

const ROOT = path.join(__dirname, "..");
const args = process.argv.slice(2);
const opt = (k, d) => (args.includes(k) ? args[args.indexOf(k) + 1] : d);
const PORT = process.env.PORT || 3455;
const BASE = process.env.BASE || `https://localhost:${PORT}`;
const agent = new https.Agent({ rejectUnauthorized: false });

function post(p, body) {
  return new Promise((resolve, reject) => {
    const u = new URL(BASE + p);
    const mod = u.protocol === "https:" ? https : http;
    const r = mod.request(u, { method: "POST", headers: { "Content-Type": "application/json" }, agent: u.protocol === "https:" ? agent : undefined }, (res) => {
      let d = "";
      res.on("data", (c) => (d += c));
      res.on("end", () => {
        try {
          resolve(JSON.parse(d));
        } catch (e) {
          reject(new Error(d.slice(0, 200)));
        }
      });
    });
    r.on("error", reject);
    r.write(JSON.stringify(body));
    r.end();
  });
}

const cov = require("../server/coverage");

/** 全角→半角、空白と桁区切りを落として比較しやすくする */
function fold(s) {
  return String(s || "")
    .replace(/[Ａ-Ｚａ-ｚ０-９％．，－]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0))
    .replace(/[,，\s]/g, "")
    .toUpperCase();
}

(async () => {
  const cases = JSON.parse(fs.readFileSync(path.join(ROOT, opt("--cases", "test/cases.json")), "utf8"));
  const model = opt("--model", "");
  const repeat = Number(opt("--repeat", 1));
  const quiet = args.includes("--quiet");
  const rows = [];
  for (let rep = 0; rep < repeat; rep++) {
    for (const c of cases) {
      const want = cov.atoms(c.prompt);
      const wantT = cov.terms(c.prompt);
      let gen;
      try {
        gen = await post("/api/generate", { prompt: c.prompt, model: model || undefined, maxSlides: 2 });
      } catch (e) {
        rows.push({ case: c.name, rep, error: String(e.message).slice(0, 120), covered: 0, total: want.length, missing: want });
        continue;
      }
      const text = JSON.stringify(gen.slides || gen.sections || gen);
      const missing = cov.missing(want, text);
      const missingT = cov.missingTerms(wantT, text);
      rows.push({ case: c.name, rep, total: want.length, covered: want.length - missing.length, missing, totalTerms: wantT.length, coveredTerms: wantT.length - missingT.length, missingTerms: missingT, slides: (gen.slides || []).length, ms: gen.ms, spec: gen.slides });
      if (!quiet)
        console.log(
          `${c.name.padEnd(16)} 数値 ${String(want.length - missing.length).padStart(3)}/${String(want.length).padEnd(3)} 語 ${String(wantT.length - missingT.length).padStart(3)}/${String(wantT.length).padEnd(3)} 枚=${(gen.slides || []).length}` +
            (missing.length ? "\n    数値落ち: " + missing.join(" / ") : "") +
            (missingT.length ? "\n    語落ち: " + missingT.join(" / ") : "")
        );
    }
  }
  const sum = (k) => rows.reduce((a, r) => a + (r[k] || 0), 0);
  const pct = (a, b) => ((a / Math.max(1, b)) * 100).toFixed(1);
  console.log(`\n数値・固有名詞: ${sum("covered")}/${sum("total")} = ${pct(sum("covered"), sum("total"))}%`);
  console.log(`日本語の内容語: ${sum("coveredTerms")}/${sum("totalTerms")} = ${pct(sum("coveredTerms"), sum("totalTerms"))}%`);
  console.log(`取りこぼしのあったケース: ${rows.filter((r) => (r.missing || []).length || (r.missingTerms || []).length).length}/${rows.length}`);
  const out = path.join(ROOT, "debug", `coverage-${Date.now()}.json`);
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, JSON.stringify(rows, null, 1));
  console.log("saved " + out);
})();
