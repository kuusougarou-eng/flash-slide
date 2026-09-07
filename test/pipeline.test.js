"use strict";
/**
 * 正規化パイプラインの契約テスト。
 *
 * 生成応答は必ず `SlideLayout.normalizeSpec` を入口にする、という 1 点だけを守れば、
 * 文法の正規化(`normalizeGeneratedSpec`)と決定論的な後処理(`refineComposition`)の
 * 両方が適用される。過去に server.js が `normalizeGeneratedSpec` を直接呼んだことで
 * 後処理が丸ごと飛ぶ事故が起きたため、その配線をテストで固定する。
 *
 * 複数人・複数エージェントで同じファイルを触るため、「誰が文法側を書き換えても
 * 後処理が外れない」ことを機械的に保証するのがこのファイルの役割。
 */
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const L = require("../public/layout");

const root = path.join(__dirname, "..");
let checks = 0;
const ok = (cond, msg) => {
  assert.ok(cond, msg);
  checks++;
};

// --- 1. 配線: 生成応答の正規化は normalizeSpec を通す ---
const serverSrc = fs.readFileSync(path.join(root, "server.js"), "utf8");
const direct = serverSrc.match(/SlideLayout\.normalizeGeneratedSpec/g) || [];
ok(direct.length === 0, `server.js は normalizeGeneratedSpec を直接呼ばない(後処理が飛ぶ)。見つかった箇所: ${direct.length}`);
ok(/SlideLayout\.normalizeSpec/.test(serverSrc), "server.js は SlideLayout.normalizeSpec を使う");

const paneSrc = fs.readFileSync(path.join(root, "public", "taskpane.js"), "utf8");
ok(!/normalizeGeneratedSpec/.test(paneSrc), "タスクペインも normalizeGeneratedSpec を直接呼ばない");

// --- 2. normalizeSpec が composition 文法を受け取り、後処理まで適用する ---
// ネイティブ表 → セル合成の格子(10 行・5 列以下)
const ntable = L.normalizeSpec({
  panelCount: 1,
  title: "評価軸ごとに 3 案を比較する",
  body: { type: "ntable", colHeaders: ["A 案", "B 案"], rows: [{ head: "費用", cells: ["1", "2"] }, { head: "期間", cells: ["3", "4"] }] },
});
ok(ntable.body.type === "table", "小さなネイティブ表は格子に落ちる(refineComposition)");
ok((ntable.compositionRepairs || []).includes("ntable → matrix"), "組み替えが記録される");

// 全行が空の列は落とす
const emptyCol = L.normalizeSpec({
  panelCount: 1,
  title: "調査結果を観点ごとに整理する",
  body: { type: "table", colHeaders: ["結果", "備考"], rows: [{ head: "r1", cells: ["x", "—"] }, { head: "r2", cells: ["y", ""] }, { head: "r3", cells: ["z", "—"] }] },
});
ok(emptyCol.body.colHeaders.length === 1, "全行が空の列は落ちる");

// 同じ前置きの繰り返しは列見出しへ昇格
const prefix = L.normalizeSpec({
  panelCount: 1,
  title: "調査で見えた実態と対応を整理する",
  body: { type: "table", colHeaders: [], rows: [{ cells: ["実態: 利用者 22% 減", "対応: 幹線維持"] }, { cells: ["実態: 運転手不足", "対応: 待遇改善"] }] },
});
ok(prefix.body.colHeaders[0] === "実態" && prefix.body.colHeaders[1] === "対応", "繰り返す前置きは列見出しに昇格する");

// 文章に埋まった系列は折れ線へ
const series = L.normalizeSpec({
  panelCount: 2,
  title: "問合せ増加を自動化で吸収する",
  panels: [
    { head: "推移", text: "月間問合せ件数は 2023 年 1,200 件、2024 年 1,500 件、2025 年 1,900 件、2026 年 2,160 件と増加した。残業で吸収している" },
    { head: "方針", text: "FAQ ボットで定型質問を自動回答する" },
  ],
});
ok(series.body.type === "line", "系列を含む文章は折れ線になる");
ok(series.body.labels.length === 4 && series.body.series[0].values[3] === 2160, "系列の値が保たれる");
ok(series.note && /残業で吸収/.test(series.note.text), "図に載らない文は補足として残る(情報を落とさない)");

// 箱フローの二重番号は外す
const seq = L.normalizeSpec({
  panelCount: 1,
  title: "移行は段階を追って進める",
  body: { type: "sequence", steps: [{ head: "01 現状把握", text: "棚卸しする" }, { head: "02 移行", text: "並行稼働する" }] },
});
ok(seq.body.steps[0].head === "現状把握", "LLM が付けた番号は外す(エンジンが振る)");

// ガント: 使われない期間を落とし、年度四半期で組み直す
const gantt = L.normalizeSpec({
  panelCount: 1,
  title: "刷新は年度をまたいで進む",
  body: {
    type: "gantt",
    groups: [{ text: "Q4", span: 6 }, { text: "Q1", span: 7 }],
    periods: ["4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", "1月", "2月", "3月", "4月"],
    tasks: [{ label: "要件定義", start: "4月", end: "5月" }, { label: "開発", start: "6月", end: "9月" }],
  },
});
ok(gantt.body.periods.length === 6, `使われない期間は落とす(残り ${gantt.body.periods.length})`);
ok(gantt.body.groups[0].text === "Q1", "四半期は年度基準(4〜6 月 = Q1)で組み直す");

// --- 3. 強調は結論に出る対象だけ(旧形式の予算) ---
const hl = L.normalizeSpec({
  title: "A 社が最適である",
  body: { cols: [{ type: "cell", head: "A 社", text: "実績が多い", highlight: true }, { type: "cell", head: "B 社", text: "安価", highlight: true }] },
});
ok(hl.body.cols[0].highlight === true, "title に名前が出る対象の強調は残る");
ok(!hl.body.cols[1].highlight, "結論に出てこない対象の強調は落ちる");

// --- 4. 描画まで通る(契約が守られていれば版面に収まる) ---
for (const spec of [ntable, emptyCol, prefix, series, seq, gantt]) {
  const lay = L.layout(spec, { width: 960, height: 540, palette: { accent: "none" } });
  ok(lay.prims.length > 0, "描画プリミティブが生成される");
  for (const p of lay.prims.filter((p) => p.body && p.kind === "rect")) {
    ok(p.x >= -0.5 && p.y >= -0.5 && p.x + p.w <= 960.5 && p.y + p.h <= 540.5, `版面内に収まる: ${(p.text || "").slice(0, 12)}`);
  }
}

console.log(`pipeline tests OK (${checks} checks)`);
