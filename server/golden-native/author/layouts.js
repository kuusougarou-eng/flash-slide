"use strict";
// Layout functions ported from scripts/make-golden.py. Each takes semantic data
// (no coordinates) and draws on a Slide. Counts are variable; positions are
// computed from the body area exactly like the Python reference.
const X = require("./xml");
const { T, BX, BY, BW, BB, BH, T_H1, T_H2, T_BODY, T_SMALL, T_NOTE, T_KPI, bullets, band, panelHead, rowhead, vchevron, harvey, dot, source, rowsFill, colsFill, pad2 } = X;
const LAYOUTS = {};
const def = (id, spec) => { LAYOUTS[id] = spec; };
const req = (cond, msg) => { if (!cond) throw Error(msg); };
const arr = (v, lo, hi, name) => { req(Array.isArray(v) && v.length >= lo && v.length <= hi, `${name} must have ${lo}–${hi} items`); return v; };

// ---- 01 executive summary: claim × facts × impact grid + total band ----
def("01", {
  name: "エグゼクティブサマリー(主張×根拠×インパクトの格子+合計帯)",
  build(s, d) {
    const rows = arr(d.rows, 2, 5, "rows");
    const cols = [["#", 40], [d.heads?.[0] || "主張", 250], [d.heads?.[1] || "根拠となる事実", 396], [d.heads?.[2] || "利益インパクト", 210]];
    let x = BX;
    for (const [lab, w] of cols) { band(s, x, BY, w - 4, 22, lab, { align: lab === "#" ? "ctr" : "l" }); x += w; }
    const totH = 40;
    rowsFill(BY + 28, BB - totH - 8, rows.length, 6).forEach(([y, rh], i) => {
      const r = rows[i]; let x = BX;
      s.text(x, y, 36, rh, String(i + 1), { size: 22, bold: true, color: T.ACCENT, align: "ctr" }); x += 40;
      s.text(x, y, 246, rh, r.claim, { size: T_H2, bold: true }); x += 250;
      s.text(x, y, 392, rh, bullets(r.facts), { size: T_BODY }); x += 396;
      s.text(x, y, 206, rh, [{ t: r.pt, size: 18, bold: true, align: "ctr" }, { t: r.yen, size: T_BODY, color: T.DARK, align: "ctr" }], { anchor: "ctr" });
      s.line(BX, y + rh + 3, BX + BW, y + rh + 3, { color: T.LIGHT });
    });
    const y = BB - totH;
    s.rect(BX, y, BW, totH, { fill: T.PALE });
    s.text(BX + 8, y, 640, totH, d.total.text, { size: T_BODY, anchor: "ctr" });
    s.text(BX + BW - 220, y, 212, totH, [{ t: d.total.pt, size: 18, bold: true, color: T.ACCENT, align: "r" }, { t: d.total.yen, size: T_BODY, color: T.DARK, align: "r" }], { anchor: "ctr" });
    source(s, d.source);
  },
});

// ---- 02 two panels (issues / actions) + correspondence band ----
def("02", {
  name: "2パネル(黒下線+中央見出し、2階層の箇条書き)+対応関係の帯",
  build(s, d) {
    const panels = arr(d.panels, 2, 2, "panels");
    const [[x1, w], [x2]] = colsFill(BX, BW, 2, 24);
    const y0 = panelHead(s, x1, BY, w, panels[0].head);
    panelHead(s, x2, BY, w, panels[1].head);
    const footH = d.band ? 92 : 0;
    const sz = d.bodySize || 12.5;
    s.text(x1, y0 + 8, w, BB - footH - y0 - 16, bullets(panels[0].items, sz), { size: sz });
    s.text(x2, y0 + 8, w, BB - footH - y0 - 16, bullets(panels[1].items, sz), { size: sz });
    if (d.band) {
      const y = BB - footH;
      band(s, BX, y, BW, 22, d.band.head, { align: "l" });
      s.text(BX, y + 26, BW, footH - 26, bullets(d.band.items, 12), { size: 12 });
    }
    source(s, d.source);
  },
});

// ---- 09 vertical enumeration: number card + name + long description + owner/KPI/effect ----
def("09", {
  name: "縦列挙(番号カード+施策名+横長の説明+担当/KPI/効果)",
  build(s, d) {
    const items = arr(d.items, 2, 7, "items");
    const [xNum, xT, wT, xD, xO, wO, xK, wK, xE, wE] = [BX, BX + 40, 170, BX + 218, BX + 610, 90, BX + 706, 130, BX + 842, 54];
    const heads = d.heads || ["施策", "内容", "担当 / 期限", "KPI(目標)", "効果\n億円/年"];
    [[heads[0], xT, wT, "l"], [heads[1], xD, xO - xD - 8, "l"], [heads[2], xO, wO, "l"], [heads[3], xK, wK, "l"], [heads[4], xE, wE, "r"]].forEach(([lab, xx, ww, al]) =>
      s.text(xx, BY, ww, 24, lab, { size: T_SMALL, bold: true, color: T.DARK, align: al, anchor: "b", margins: [0, 0, 0, 2] }));
    s.line(BX, BY + 26, BX + BW, BY + 26, { color: T.INK, w: 1 });
    const top = d.highlight ?? 2;
    rowsFill(BY + 32, BB, items.length, 6).forEach(([y, rh], k) => {
      const it = items[k], dark = k < top;
      s.rect(xNum, y, 30, 30, { fill: dark ? T.DARK : T.LIGHT, text: { content: pad2(k + 1), size: 12, bold: true, color: dark ? T.WHITE : T.INK } });
      s.text(xT, y, wT, rh, it.name, { size: 13, bold: true, margins: [0, 0, 4, 0] });
      s.text(xD, y, xO - xD - 8, rh, it.desc, { size: 12, margins: [0, 0, 0, 0] });
      s.text(xO, y, wO, rh, it.owner, { size: T_BODY, color: T.DARK, margins: [0, 0, 0, 0] });
      s.text(xK, y, wK, rh, it.kpi, { size: T_BODY, margins: [0, 0, 0, 0] });
      s.text(xE, y, wE, rh, it.effect, { size: 18, bold: true, align: "r", anchor: "ctr", margins: [0, 0, 0, 0] });
      s.line(BX, y + rh + 3, BX + BW, y + rh + 3, { color: T.LIGHT });
    });
    source(s, d.source);
  },
});

require("./layouts-b")(def);
for (const f of ["./layouts-c", "./layouts-d", "./layouts-e", "./layouts-f"]) { try { require(f)(def); } catch (e) { if (e.code !== "MODULE_NOT_FOUND" || !String(e.message).includes(f.slice(2))) throw e; } }
module.exports = { LAYOUTS };
