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

/** 全角→半角、空白と桁区切りを落として比較しやすくする */
function fold(s) {
  return String(s || "")
    .replace(/[Ａ-Ｚａ-ｚ０-９％．，－]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0))
    .replace(/[,，\s]/g, "")
    .toUpperCase();
}

/**
 * 落としてはいけない原子を抜く。
 *  - 数値 + 単位(22%, 3,860, 54.2歳, 212路線, 38億円, 1.6倍, 4.1/5, 18か月, 2週間 …)
 *  - 英数の固有名詞らしき語(BI, FAQ, PoC, PMO, A社 …)
 */
function atoms(prompt) {
  const out = new Set();
  const t = String(prompt || "");
  const numRe = /\d[\d,，.]*\s*(?:%|％|割|倍|件|名|社|人|歳|億円|万円|円|路線|本|テーブル|か月|ヶ月|カ月|か年|年|月|週間|週|日|回|時間|pt|ポイント|\/\s*\d+)?/g;
  let m;
  while ((m = numRe.exec(t))) {
    const raw = m[0].trim();
    if (fold(raw).replace(/[^0-9]/g, "").length === 0) continue;
    // 単位の無い裸の数字は、桁が小さいと箇条書きの番号などと紛れるので 3 桁以上だけ拾う
    const bare = /^\d[\d,，.]*$/.test(raw);
    if (bare && fold(raw).replace(/[^0-9]/g, "").length < 3) continue;
    out.add(raw);
  }
  for (const w of t.match(/\b[A-Z][A-Za-z]{1,9}\b/g) || []) out.add(w);
  for (const w of t.match(/[A-Z]\s?社/g) || []) out.add(w);
  return [...out];
}

function has(hay, atom) {
  return fold(hay).includes(fold(atom));
}

(async () => {
  const cases = JSON.parse(fs.readFileSync(path.join(ROOT, opt("--cases", "test/cases.json")), "utf8"));
  const model = opt("--model", "");
  const repeat = Number(opt("--repeat", 1));
  const quiet = args.includes("--quiet");
  const rows = [];
  for (let rep = 0; rep < repeat; rep++) {
    for (const c of cases) {
      const want = atoms(c.prompt);
      let gen;
      try {
        gen = await post("/api/generate", { prompt: c.prompt, model: model || undefined, maxSlides: 2 });
      } catch (e) {
        rows.push({ case: c.name, rep, error: String(e.message).slice(0, 120), covered: 0, total: want.length, missing: want });
        continue;
      }
      const text = JSON.stringify(gen.slides || gen.sections || gen);
      const missing = want.filter((a) => !has(text, a));
      rows.push({ case: c.name, rep, total: want.length, covered: want.length - missing.length, missing, slides: (gen.slides || []).length, ms: gen.ms, spec: gen.slides });
      if (!quiet) console.log(`${c.name.padEnd(16)} ${String(want.length - missing.length).padStart(3)}/${String(want.length).padEnd(3)} 枚=${(gen.slides || []).length}${missing.length ? "  落ち: " + missing.join(" / ") : ""}`);
    }
  }
  const tot = rows.reduce((a, r) => a + r.total, 0);
  const cov = rows.reduce((a, r) => a + r.covered, 0);
  console.log(`\n被覆率: ${cov}/${tot} = ${((cov / Math.max(1, tot)) * 100).toFixed(1)}%  (取りこぼしのあったケース ${rows.filter((r) => r.missing.length).length}/${rows.length})`);
  const out = path.join(ROOT, "debug", `coverage-${Date.now()}.json`);
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, JSON.stringify(rows, null, 1));
  console.log("saved " + out);
})();
