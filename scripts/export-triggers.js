/**
 * run-triggers.js の結果(debug/triggers-*.json)から slideId を拾い、開いている PowerPoint の該当スライドを PNG に書き出す。
 *   node scripts/export-triggers.js [debug/triggers-XXXX.json] [--deck zeroshot] [--prefix t-]
 * 既定は最新の triggers-*.json、デッキ名 "zeroshot"、出力 debug/t-<name>-<n>.png
 */
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const ROOT = path.join(__dirname, "..");
const args = process.argv.slice(2);
const opt = (k, d) => (args.includes(k) ? args[args.indexOf(k) + 1] : d);
const deck = opt("--deck", "zeroshot");
const prefix = opt("--prefix", "t-");
let file = args.find((a) => a.endsWith(".json"));
if (!file) {
  const dir = path.join(ROOT, "debug");
  file = fs
    .readdirSync(dir)
    .filter((f) => /^triggers-\d+\.json$/.test(f))
    .sort()
    .pop();
  file = path.join(dir, file);
}
const results = JSON.parse(fs.readFileSync(file, "utf8"));
const out = [];
for (const r of results) {
  const ids = [];
  for (const ln of r.lines || []) {
    // 自分の tag(= ケース名で終わる)の完了ログだけから slideId を拾う(別ペインのログを混ぜない)
    const m = ln.match(/"slide":"([^"]+)","slideId":"([^"]+)"/);
    if (m && m[1].endsWith(r.name)) ids.push(m[2]);
  }
  ids.forEach((id, i) => {
    const png = path.join(ROOT, "debug", `${prefix}${r.name}-${i + 1}.png`);
    try {
      const res = execFileSync("powershell.exe", ["-NoProfile", "-File", path.join(ROOT, "scripts", "ppt-export-by-id.ps1"), deck, id, png], { encoding: "utf8" });
      out.push(`${r.name}[${i + 1}] ${res.trim()}`);
    } catch (e) {
      out.push(`${r.name}[${i + 1}] FAILED ${(e.stdout || e.message || "").toString().trim()}`);
    }
  });
  if (!ids.length) out.push(`${r.name} no slideId (${(r.lines || []).length} lines)`);
}
console.log(out.join("\n"));
