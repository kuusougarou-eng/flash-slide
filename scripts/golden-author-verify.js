"use strict";
// Proves the JS port: for every layout with golden sample data, the rendered
// slide XML must equal the reference slide XML (ids/names normalized), and the
// bound chart XML must equal the reference chart XML.
//   node scripts/golden-author-verify.js [01,02,...]
const AdmZip = require("adm-zip");
const P = require("../server/golden-native/package");
const author = require("../server/golden-native/author");
const { SAMPLES } = require("../server/golden-native/author/samples");
const only = (process.argv[2] || "").split(",").filter(Boolean);
const g = author.getCompiler();
const strip = (doc) => { for (const n of P.els(doc, "p", "cNvPr")) { n.removeAttribute("id"); n.removeAttribute("name"); } return doc; };
const parts = (xml) => { const d = strip(P.parse(xml)); const tree = P.els(d, "p", "spTree")[0]; return Array.from(tree.childNodes).filter((n) => n.nodeType === 1).map((n) => P.xml(n)); };
const chartParts = (zip, part) => P.relationships(zip, part).filter((r) => r.getAttribute("Type").endsWith("/chart")).map((r) => P.target(part, r.getAttribute("Target")));
let ok = 0, bad = 0;
for (const id of Object.keys(author.LAYOUTS).sort()) {
  if (only.length && !only.includes(id)) continue;
  const sample = SAMPLES[id];
  if (!sample) { console.log(id, "no sample"); continue; }
  try {
    const { buffer, xml } = author.build(id, sample, { page: String(Number(id)) });
    const slide = g.get(id), ref = new AdmZip(g.reference(id)), out = new AdmZip(buffer);
    const a = parts(P.read(ref, slide.part)), b = parts(xml);
    let diff = null;
    const win = (x, y) => { let k = 0; while (k < x.length && x[k] === y[k]) k++; return [x.slice(Math.max(0, k - 250), k + 250), y.slice(Math.max(0, k - 250), k + 250)]; };
    for (let i = 0; i < Math.max(a.length, b.length); i++) if (a[i] !== b[i]) { const [r, o] = win(a[i] || "", b[i] || ""); diff = { i, ref: r, out: o }; break; }
    if (!diff) {
      for (const cp of chartParts(ref, slide.part)) {
        const norm = (t) => P.xml(P.parse(t)).replace(/ xmlns:[ac]="[^"]+"/g, "").replace(/<c:v>(-?\d+(?:\.\d+)?)<\/c:v>/g, (m, v) => "<c:v>" + String(+(+v).toFixed(6)) + "</c:v>");
        const ra = norm(P.read(ref, cp)), rb = norm(P.read(out, cp));
        if (ra !== rb) { let k = 0; while (k < ra.length && ra[k] === rb[k]) k++; diff = { chart: cp, ref: ra.slice(Math.max(0, k - 200), k + 300), out: rb.slice(Math.max(0, k - 200), k + 300) }; break; }
      }
    }
    if (!diff) { ok++; console.log(id, "OK", a.length, "shapes"); }
    else { bad++; console.log(id, "DIFF", diff.chart ? "chart " + diff.chart : "shape#" + diff.i + " of " + a.length + "/" + b.length); console.log("  ref:", (diff.ref || "(none)").slice(0, 900)); console.log("  out:", (diff.out || "(none)").slice(0, 900)); }
  } catch (e) { bad++; console.log(id, "ERROR", e.message); if (process.env.STACK) console.log(e.stack); }
}
console.log(`identical ${ok}, differing ${bad}`);
process.exitCode = bad ? 1 : 0;
