"use strict";
// The JS port of scripts/make-golden.py must reproduce every golden slide from
// its data, and must keep working when counts, statuses and text lengths change.
const assert = require("assert"), fs = require("fs"), path = require("path"), AdmZip = require("adm-zip");
const P = require("../server/golden-native/package");
const source = path.resolve("docs/golden/golden.pptx");
if (!fs.existsSync(source)) { console.log("golden-author skipped: local GoldenPPTX is not installed"); process.exit(0); }
const author = require("../server/golden-native/author");
const { SAMPLES } = require("../server/golden-native/author/samples");
const gen = require("../server/golden-native/author/generate");
const g = author.getCompiler();
const strip = (doc) => { for (const n of P.els(doc, "p", "cNvPr")) { n.removeAttribute("id"); n.removeAttribute("name"); } return doc; };
const canon = (xml) => P.xml(strip(P.parse(xml)).documentElement); // ignore the XML declaration

// 1. Every layout reproduces its reference slide XML from the sample data.
assert.equal(Object.keys(author.LAYOUTS).length, 32);
for (const id of Object.keys(author.LAYOUTS)) {
  const { xml } = author.build(id, SAMPLES[id], { page: String(Number(id)) });
  assert.equal(canon(xml), canon(P.read(new AdmZip(g.reference(id)), g.get(id).part)), "layout " + id + " must equal the golden slide XML");
}

// 2. Variable counts re-flow the body area (no golden coordinates baked in).
const nine = { ...SAMPLES["09"], items: SAMPLES["09"].items.slice(0, 3) };
const three = author.render("09", nine).xml;
assert.equal((three.match(/FS_Rectangle/g) || []).length, 3, "one number card per item");
const risks = { ...SAMPLES["27"], rows: [...SAMPLES["27"].rows, ...SAMPLES["27"].rows.slice(0, 2)] };
assert.equal((author.render("27", risks).xml.match(/FS_Rectangle/g) || []).length, 4 + 6, "bands + one badge per row");
assert.throws(() => author.render("27", { ...SAMPLES["27"], rows: SAMPLES["27"].rows.slice(0, 1) }), /2–6 items/);
const gantt = { ...SAMPLES["13"], periods: SAMPLES["13"].periods.slice(0, 6), tasks: SAMPLES["13"].tasks.filter((t) => t.end < 6).slice(0, 4), milestones: [{ label: "A", at: 5 }], today: null };
assert.ok(author.render("13", gantt).xml.includes("A"));
assert.throws(() => author.render("13", { ...gantt, tasks: [{ label: "x", owner: "y", start: 0, end: 9, critical: false }] }), /out of range/);

// 3. Semantic values are normalised, geometry-only fields are dropped, statuses accept words.
const kpi = JSON.parse(JSON.stringify(SAMPLES["17"]));
kpi.tiles[0].status = "順調"; kpi.tiles[1].status = "遅延"; kpi.statuses[0].status = "要注意";
assert.ok(author.render("17", kpi).xml.includes("2E7D32"));
assert.throws(() => author.render("17", { ...kpi, tiles: [{ ...kpi.tiles[0], status: "purple" }, kpi.tiles[1]] }), /status must be/);
const plan = { layout: "3", data: JSON.parse(JSON.stringify(SAMPLES["03"])) };
plan.data.annotation.line = true;
const compiled = gen.compile(plan);
assert.equal(compiled.layout, "03");
assert.ok(!plan.data.annotation.line, "model cannot place the calibrated pointer line");
assert.throws(() => gen.compile({ layout: "99", data: { title: "x" } }), /Unknown layout/);
assert.throws(() => gen.compile({ layout: "01", data: { lead: "x" } }), /title is required/);

// 4. Overlong text shrinks stepwise instead of overflowing silently.
const long = JSON.parse(JSON.stringify(SAMPLES["27"]));
long.rows[0].action = "対応".repeat(120);
const warnings = [];
author.build("27", long, { fit: { floor: 8, log: (w) => warnings.push(w) } });
assert.ok(warnings.length >= 1 && warnings[0].k < 1, "fit must report the shrink factor");

// 5. One inference call, catalog contains all layouts, single-layout mode narrows it.
assert.equal((gen.messages("x")[0].content.match(/## layout /g) || []).length, 32);
assert.equal((gen.messages("x", { layout: "13" })[0].content.match(/## layout /g) || []).length, 1);
let calls = 0;
gen.generate("検証入力", { chat: async () => { calls++; return { content: JSON.stringify({ layout: "27", data: SAMPLES["27"] }), model: "fake", usage: {} }; } })
  .then((r) => { assert.equal(calls, 1); assert.equal(r.inferenceCalls, 1); assert.ok(r.buffer.length > 1000); console.log("golden author: 32/32 XML-identical, variable counts, normalisation, fitting, one inference verified"); })
  .catch((e) => { console.error(e); process.exitCode = 1; });
