"use strict";
const assert = require("assert");
const L = require("../public/layout");
const { schemaDoc } = require("../server/prompt");
const norm = L.normalizeGeneratedSpec;
const draw = (s, opts = {}) => L.layout(s, { width: 960, height: 540, palette: { accent: "none" }, ...opts });
function bounds(l) {
  assert.ok(!l.warnings.some((w) => /overflow|failed|shrunk/.test(w)), l.warnings.join("; "));
  for (const p of l.prims.filter((p) => p.body && p.kind !== "line")) {
    const rotated = (p.rotation || 0) % 180 === 90;
    const w = rotated ? p.h : p.w, h = rotated ? p.w : p.h;
    const x = p.x + (p.w - w) / 2, y = p.y + (p.h - h) / 2;
    assert.ok(x >= l.bodyRect.x - .01 && x + w <= l.bodyRect.x + l.bodyRect.w + .01);
    assert.ok(y >= l.bodyRect.y - .01 && y + h <= l.bodyRect.y + l.bodyRect.h + .01);
  }
}
const panels = norm({ panelCount: 2, title: "原因を把握して打ち手を選ぶ", panels: [
  { head: "原因", items: ["内部要因 50%", ["商品の陳腐化 15%", "品質問題 10%"], "その他 20%"], highlight: true },
  { head: "打ち手", text: "商品力への投資を優先する", fill: "dark", highlight: true }
] });
assert.deepStrictEqual(norm(panels), panels, "new grammar is idempotent when re-rendering saved JSON");
assert.deepStrictEqual(panels.body.cols[0].items[1], ["商品の陳腐化 15%", "品質問題 10%"], "nested content survives");
for (const width of [720, 960, 1280]) {
  const l = draw(panels, { width }); bounds(l);
  const rules = l.prims.filter((p) => p.body && p.kind === "line" && p.weight === L.STYLE.rule.thick);
  assert.strictEqual(rules.length, 2);
  assert.strictEqual(rules[0].y1, rules[1].y1);
  assert.strictEqual(rules[0].x2 - rules[0].x1, rules[1].x2 - rules[1].x1);
  assert.ok(!l.prims.some((p) => p.body && p.fill), "text panels have no arbitrary fills");
  assert.ok(l.prims.filter((p) => p.body && p.text && !p.role).every((p) => p.valign === "middle"));
  assert.ok(l.fonts.head > l.fonts.body && l.fonts.body >= 18);
}
const short = norm({ panelCount: 1, body: { type: "sequence", steps: [
  { head: "調査", text: "課題を確認する", highlight: true },
  { head: "設計", text: "方針を決める" }, { head: "実行", text: "施策を実行する" }
] } });
let l = draw(short); bounds(l);
assert.strictEqual(l.prims.filter((p) => p.body && p.shape === "homePlate" && !p.rotation).length, 3);
assert.ok(l.prims.some((p) => /^01/.test(p.text)), "ordered steps retain numbers");
const long = norm({ panelCount: 1, body: { type: "sequence", steps: short.body.steps.map((s) => ({ ...s, text: "業務の依存関係を確認し、関係者と合意した判断基準を満たしてから次の段階へ移行する。".repeat(2) })), note: "稼働率 99.5% 以上、再処理 週 3 件以下" } });
assert.ok(long.note.text.includes("99.5%"), "misplaced LLM note must not lose thresholds");
l = draw(long); bounds(l);
assert.strictEqual(l.prims.filter((p) => p.rotation === 90 && p.shape === "homePlate").length, 3);
assert.ok(l.prims.some((p) => (p.text || "").includes("99.5%") && !p.fill), "supplement is one unboxed full-width text");
assert.deepStrictEqual(norm(long), long);
const table = { type: "table", colHeaders: ["金額", "期間"], rows: [
  { head: "A社", cells: ["4.2 億円", "18 か月"] }, { head: "B社", cells: ["3.6 億円", "12 か月"] }
] };
const small = norm({ panelCount: 1, body: table });
assert.strictEqual(small.body.colHeaders.length, 0);
assert.ok(small.body.rows.every((r) => !r.head));
for (const text of ["金額", "期間", "A社", "B社", "4.2 億円", "12 か月"]) assert.ok(JSON.stringify(small).includes(text), text + " preserved inline");
assert.deepStrictEqual(norm(small), small);
const directional = norm({ panelCount: 1, body: { ...table, headShape: "chevron" } });
assert.ok(draw(directional).prims.some((p) => /^01\s+A社/.test(p.text)), "directional matrix headers are numbered");
const nested = norm({ panelCount: 2, panels: [{ rows: [{ type: "cell", head: "比較" }, table] }, { head: "打ち手", text: "商品力を改善する" }] });
assert.ok(nested.body.cols.every((p) => p.type === "cell"));
assert.ok(JSON.stringify(nested).includes("4.2 億円"), "nested matrix becomes text without dropping values");
const dense = norm({ panelCount: 1, body: { ...table, colHeaders: ["金額", "期間", "実績"], rows: Array.from({ length: 8 }, (_, i) => ({ head: String(i), cells: ["2億円", "1年", "3社"] })) } });
assert.strictEqual(dense.body.type, "ntable");
l = draw(dense); bounds(l);
assert.strictEqual(l.prims.filter((p) => p.kind === "table").length, 1, "dense matrix uses actual native table primitive");
const unnamedDense = norm({ panelCount: 1, body: { type: "table", rows: Array.from({ length: 8 }, (_, i) => ({ head: "項目" + i, cells: ["内容" + i, "条件" + i] })) } });
assert.strictEqual(draw(unnamedDense).prims.find((p) => p.kind === "table").cells[8][2].text, "条件7", "native table keeps data columns even without supplied column names");
assert.strictEqual(norm({ panelCount: 1, panels: [{ head: "内容", text: "本文を保持する" }] }).body.items[0], "本文を保持する");
const gantt = norm({ panelCount: 1, body: { type: "gantt", periods: ["4月", "5月", "6月"], groups: [{ text: "2026", span: 3 }, { text: "2026", span: 3 }], tasks: [{ label: "設計", start: "4月", end: "6月", note: "IT 部門が実行責任者", highlight: true }], milestones: [{ label: "最終月末にリリースする", at: "6月" }] }, note: "PM 1 名、IT 5 名" });
assert.strictEqual(gantt.body.groups.reduce((n, g) => n + g.span, 0), 3);
assert.ok(!gantt.body.tasks[0].highlight);
assert.deepStrictEqual(norm(gantt), gantt);
bounds(draw(gantt));
assert.throws(() => norm({ panelCount: 1, body: { rows: [table, { type: "bars", items: [{ label: "売上", value: 10 }] }] } }), /独立した図/);
assert.throws(() => norm({ panelCount: 3, panels: [{ head: "A" }, { head: "B" }] }), /一致/);
assert.throws(() => norm({ panelCount: 2, panels: [{ head: "A" }, { head: "B" }], note: "別の箱" }), /補足/);
assert.ok(!schemaDoc().includes('"highlight":true'), "prompt examples must not contradict neutral emphasis rule");
console.log("composition tests OK");
