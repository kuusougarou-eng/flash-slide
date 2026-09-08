"use strict";
// Layouts 16–20 (ported from scripts/make-golden.py).
const X = require("./xml");
const { T, BX, BY, BW, BB, BH, T_H1, T_H2, T_BODY, T_SMALL, T_NOTE, T_KPI, bullets, band, panelHead, dot, source, rowsFill, colsFill } = X;
const req = (cond, msg) => { if (!cond) throw Error(msg); };
const arr = (v, lo, hi, name) => { req(Array.isArray(v) && v.length >= lo && v.length <= hi, `${name} must have ${lo}–${hi} items`); return v; };
const STATUS = { green: T.GREEN, amber: T.AMBER, red: T.RED, yellow: T.AMBER, orange: T.AMBER, ok: T.GREEN, warn: T.AMBER, ng: T.RED };
const STATUS_WORDS = [[/順調|計画どおり|計画通り|完了|良|達成|on ?track|good|done/i, "green"], [/遅延|遅れ|未達|悪|停止|critical|late|behind|bad/i, "red"], [/注意|懸念|要確認|やや|warning|risk|at ?risk/i, "amber"]];
/** Accepts green|amber|red plus common Japanese/English status words; the slide only shows the colour. */
function normalizeStatus(v) {
  const key = String(v ?? "").trim().toLowerCase();
  if (STATUS[key]) return key === "yellow" || key === "orange" || key === "warn" ? "amber" : key === "ok" ? "green" : key === "ng" ? "red" : key;
  for (const [re, name] of STATUS_WORDS) if (re.test(key)) return name;
  throw Error("status must be green|amber|red (got " + v + ")");
}
const statusColor = (s) => STATUS[normalizeStatus(s)];

module.exports = function register(def) {
  // ---- 16 action plan: native table with wave separator rows and a total row ----
  def("16", {
    name: "実行計画(ネイティブ表: 波の区切り行、合計行)",
    build(s, d) {
      const hdr = arr(d.heads, 4, 9, "heads"), n = hdr.length;
      const widths = d.widths || Array.from({ length: n }, () => BW / n);
      req(widths.length === n, "widths must match heads");
      const numeric = new Set(d.numericCols || []), centered = new Set(d.centerCols || []);
      const rows = arr(d.rows, 3, 16, "rows");
      const nSec = rows.filter((r) => r.section).length, nItem = rows.length - nSec;
      const secH = 20, hdrH = 28, itemH = (BH - hdrH - secH * nSec) / nItem;
      const heights = [hdrH, ...rows.map((r) => (r.section ? secH : itemH))];
      const cells = [hdr.map((h, j) => ({ t: h, size: T_H2, bold: true, color: T.WHITE, fill: T.DARK, align: centered.has(j) || numeric.has(j) ? "ctr" : "l" }))];
      for (const r of rows) {
        if (r.section) { cells.push([{ t: r.section, bold: true, fill: T.PALE, span: n }, ...Array.from({ length: n - 1 }, () => ({ merged: true, raw: true }))]); continue; }
        const total = !!r.total, vals = total ? [r.cells[0], "", ...r.cells.slice(2)] : r.cells;
        req(vals.length === n, "row must have " + n + " cells");
        cells.push(vals.map((v, j) => ({ t: v, bold: total, fill: total ? T.PALE : null, align: numeric.has(j) ? "r" : centered.has(j) ? "ctr" : "l", span: total && j === 0 ? 2 : 1, merged: total && j === 1 })));
      }
      s.table(BX, BY, BW, BH, widths, heights, cells);
      source(s, d.source);
    },
  });

  // ---- 17 KPI dashboard: 4 tiles + monthly line chart + small table + status list ----
  def("17", {
    name: "KPI ダッシュボード(4 タイル RAG 色+月次折れ線+小表+施策の状態と対応)",
    build(s, d) {
      const tiles = arr(d.tiles, 2, 4, "tiles"), th = 78;
      colsFill(BX, BW, tiles.length, 12).forEach(([x, tw], i) => {
        const t = tiles[i], col = statusColor(t.status);
        s.rect(x, BY, tw, th, { fill: T.PALE }); s.rect(x, BY, 4, th, { fill: col });
        s.text(x + 10, BY + 4, tw - 14, 14, t.label, { size: T_SMALL, color: T.DARK, margins: [0, 0, 0, 0] });
        s.text(x + 10, BY + 18, tw - 14, 34, t.value, { size: T_KPI, bold: true, margins: [0, 0, 0, 0] });
        s.text(x + 10, BY + 50, tw - 14, 12, t.plan, { size: 8, color: T.MUTED, margins: [0, 0, 0, 0] });
        s.text(x + tw - 110, BY + 50, 104, 24, t.delta, { size: T_SMALL, bold: true, color: col, align: "r", anchor: "ctr", margins: [0, 0, 4, 0] });
      });
      const y1 = BY + th + 12, cw = 470, rx = BX + cw + 24, rw = BX + BW - rx;
      band(s, BX, y1, cw, 20, d.chartHead, { align: "l", size: T_BODY });
      band(s, rx, y1, rw, 20, d.statusHead, { align: "l", size: T_BODY });
      const chh = BB - y1 - 24 - 74;
      const series17 = arr(d.series, 1, 2, "series");
      s.chart(BX, y1 + 22, cw, chh, { categories: d.categories, series: series17.map((x) => ({ name: x.name, values: x.values })), domain: require("./layouts-b").usableDomain(d.domain, series17.flatMap((x) => x.values)) });
      const ty = y1 + 22 + chh + 2;
      s.text(BX, ty, cw, 14, d.tableHead, { size: T_SMALL, bold: true, margins: [0, 0, 0, 0] });
      s.line(BX, ty + 15, BX + cw, ty + 15, { color: T.INK, w: 0.75 });
      const table = arr(d.table, 2, 5, "table"), nc = table[0].length, widths = [cw - 68 * (nc - 1), ...Array.from({ length: nc - 1 }, () => 68)];
      let yy = ty + 17;
      table.forEach((r, i) => {
        let xx = BX;
        r.forEach((v, j) => { s.text(xx, yy, widths[j], 13, v, { size: 8.5, bold: i === 0 || i === table.length - 1, color: i === 0 ? T.DARK : T.INK, align: j === 0 ? "l" : "r", margins: [0, 0, 0, 0], anchor: "ctr" }); xx += widths[j]; });
        s.line(BX, yy + 14, BX + cw, yy + 14, { color: T.PALE }); yy += 14;
      });
      const st = arr(d.statuses, 2, 6, "statuses");
      rowsFill(y1 + 26, BB, st.length, 3).forEach(([y, rh], i) => {
        const it = st[i], col = statusColor(it.status);
        dot(s, rx + 7, y + rh / 2, 10, col);
        s.text(rx + 16, y, 100, rh, it.name, { size: T_SMALL + 0.5, bold: true, anchor: "ctr", margins: [0, 0, 0, 0] });
        s.text(rx + 118, y, 82, rh, it.state, { size: T_SMALL, color: col, bold: true, anchor: "ctr", margins: [0, 0, 0, 0] });
        s.text(rx + 202, y, rw - 202, rh, it.action, { size: T_SMALL, anchor: "ctr", margins: [0, 0, 0, 0] });
        s.line(rx, y + rh + 1, rx + rw, y + rh + 1, { color: T.PALE });
      });
      source(s, d.source);
    },
  });

  // ---- 18 issue tree: root → groups → issues with evidence and status badge ----
  def("18", {
    name: "論点ツリー(根→2 側面→論点+検証結果+状態バッジ)+未検証論点の扱い",
    build(s, d) {
      const tbox = (x, y, w, h, t, sub, dark, fill) => s.rect(x, y, w, h, { fill: dark ? T.DARK : fill || T.PALE, text: sub ? { content: [{ t, size: T_H2, bold: true, color: dark ? T.WHITE : T.INK }, { t: sub, size: T_SMALL, color: dark ? T.WHITE : T.DARK }], size: T_H2, bold: true, color: dark ? T.WHITE : T.INK } : { content: t, size: T_BODY, bold: true, color: dark ? T.WHITE : T.INK } });
      const x0 = BX, w0 = 120, x1 = BX + 150, w1 = 120, x2 = BX + 300, w2 = 250, x3 = BX + 570, w3 = BX + BW - x3;
      band(s, x2, BY, w2, 18, d.heads[0], { fill: T.LIGHT, color: T.INK, size: T_SMALL });
      band(s, x3, BY, w3, 18, d.heads[1], { fill: T.LIGHT, color: T.INK, size: T_SMALL, align: "l" });
      const footH = d.foot ? 40 : 0, groups = arr(d.groups, 2, 3, "groups"), total = groups.reduce((n, g) => n + g.items.length, 0);
      req(total >= 3 && total <= 8, "issue tree needs 3–8 issues in total");
      const rows = rowsFill(BY + 24, BB - footH - 8, total, 6);
      const verified = d.verifiedLabel || "検証済", unverified = d.unverifiedLabel || "未検証";
      let k = 0; const centers = [];
      for (const g of groups) {
        const ys = rows.slice(k, k + g.items.length); k += g.items.length;
        const gy = ys[0][0], gh = ys.at(-1)[0] + ys.at(-1)[1] - gy, gc = gy + gh / 2;
        centers.push(gc);
        tbox(x1, gc - 20, w1, 40, g.label);
        s.line(x1 + w1, gc, x2 - 12, gc, { color: T.DARK, w: 1 });
        s.line(x2 - 12, ys[0][0] + ys[0][1] / 2, x2 - 12, ys.at(-1)[0] + ys.at(-1)[1] / 2, { color: T.DARK, w: 1 });
        ys.forEach(([yy, rh], i) => {
          const it = g.items[i];
          s.line(x2 - 12, yy + rh / 2, x2, yy + rh / 2, { color: T.DARK, w: 1 });
          tbox(x2, yy, w2, rh, it.question, null, false, it.verified ? T.PALE : T.WHITE);
          if (!it.verified) s.rect(x2, yy, w2, rh, { line: T.DARK, lw: 1 });
          s.text(x3, yy, w3 - 60, rh, it.evidence, { size: T_BODY, anchor: "ctr", margins: [4, 0, 0, 0] });
          s.rect(x3 + w3 - 56, yy + rh / 2 - 10, 52, 20, { fill: it.verified ? T.DARK : T.WHITE, line: T.DARK, text: { content: it.verified ? verified : unverified, size: T_NOTE, bold: true, color: it.verified ? T.WHITE : T.INK } });
        });
      }
      const mid = (rows[0][0] + rows.at(-1)[0] + rows.at(-1)[1]) / 2;
      tbox(x0, mid - 26, w0, 52, d.root.label, d.root.sub, true);
      s.line(x0 + w0, mid, x1 - 12, mid, { color: T.DARK, w: 1 });
      s.line(x1 - 12, centers[0], x1 - 12, centers.at(-1), { color: T.DARK, w: 1 });
      for (const c of centers) s.line(x1 - 12, c, x1, c, { color: T.DARK, w: 1 });
      if (d.foot) { s.rect(BX, BB - footH, BW, footH, { fill: T.PALE }); s.text(BX + 8, BB - footH, BW - 16, footH, d.foot, { size: T_BODY, anchor: "ctr" }); }
      source(s, d.source);
    },
  });

  // ---- 19 status report: RAG strip + 3 panels + issues/risks table ----
  def("19", {
    name: "定例報告(RAG の帯+3 パネル(実績/予定/決定依頼)+課題・リスクの表)",
    build(s, d) {
      const st = arr(d.statuses, 2, 5, "statuses");
      colsFill(BX, BW, st.length, 10).forEach(([x, w], i) => {
        const it = st[i], col = statusColor(it.status);
        s.rect(x, BY, w, 30, { fill: T.PALE }); dot(s, x + 14, BY + 15, 12, col);
        s.text(x + 26, BY, w - 30, 30, [{ t: it.label, size: T_SMALL, color: T.DARK }, { t: it.text, size: T_H2, bold: true }], { anchor: "ctr", margins: [0, 0, 0, 0] });
      });
      const y = BY + 40, panels = arr(d.panels, 2, 3, "panels"), cols = colsFill(BX, BW, panels.length, 20);
      const riskH = d.risks ? 112 : 0;
      cols.forEach(([x, w], i) => { const y0 = panelHead(s, x, y, w, panels[i].head, { size: T_H2 }); s.text(x, y0 + 6, w, BB - riskH - y0 - 14, bullets(panels[i].items), { size: T_BODY }); });
      if (d.risks) {
        const ry = BB - riskH, r = d.risks;
        band(s, BX, ry, BW, 20, r.head, { fill: T.DARK, align: "l", size: T_BODY });
        const hdrs = r.heads, widths = r.widths || [24, 330, 220, 220, BW - 24 - 330 - 220 - 220];
        let xx = BX;
        hdrs.forEach((lab, j) => { s.text(xx, ry + 22, widths[j] - 4, 14, lab, { size: T_NOTE, bold: true, color: T.DARK, margins: [2, 0, 2, 0], align: j === 0 ? "ctr" : "l" }); xx += widths[j]; });
        s.line(BX, ry + 37, BX + BW, ry + 37, { color: T.LIGHT });
        const rows = arr(r.rows, 1, 4, "risk rows");
        rowsFill(ry + 40, BB, rows.length, 2).forEach(([yr, rh], i) => {
          let xx = BX;
          rows[i].forEach((v, j) => { s.text(xx, yr, widths[j] - 4, rh, v, { size: T_SMALL, bold: j === 0, anchor: "ctr", margins: [2, 0, 2, 0], align: j === 0 ? "ctr" : "l" }); xx += widths[j]; });
          s.line(BX, yr + rh + 1, BX + BW, yr + rh + 1, { color: T.PALE });
        });
      }
      source(s, d.source);
    },
  });

  // ---- 20 survey: 100% stacked horizontal bar + findings + respondent table ----
  def("20", {
    name: "調査報告(100% 積み上げ横棒、灰階調)+主な発見+回答者属性",
    build(s, d) {
      const cw = 540;
      band(s, BX, BY, cw, 22, d.chartHead, { align: "l" });
      s.chart(BX, BY + 26, cw, BB - BY - 26, { categories: arr(d.categories, 2, 8, "categories"), series: arr(d.series, 2, 5, "series").map((x) => ({ name: x.name, values: x.values })) });
      const rx = BX + cw + 24, rw = BX + BW - rx;
      band(s, rx, BY, rw, 22, d.insightsHead, { align: "l" });
      s.text(rx, BY + 28, rw, 190, bullets(d.insights), { size: T_BODY });
      const ty = BY + 236;
      s.text(rx, ty, rw, 16, d.tableHead, { size: T_H2, bold: true, margins: [0, 0, 0, 0] });
      s.line(rx, ty + 18, rx + rw, ty + 18, { color: T.INK, w: 0.75 });
      const table = arr(d.table, 2, 6, "table"), nc = table[0].length, widths = [rw - 55 * (nc - 1), ...Array.from({ length: nc - 1 }, () => 55)];
      let yy = ty + 20;
      table.forEach((r, i) => {
        let xx = rx;
        r.forEach((v, j) => { s.text(xx, yy, widths[j], 17, v, { size: T_SMALL, bold: i === 0, color: i === 0 ? T.DARK : T.INK, align: j === 0 ? "l" : "r", margins: [0, 0, 0, 0], anchor: "ctr" }); xx += widths[j]; });
        s.line(rx, yy + 18, rx + rw, yy + 18, { color: T.PALE }); yy += 19;
      });
      source(s, d.source);
    },
  });
};
