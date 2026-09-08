"use strict";
// Layouts 21–25 (ported from scripts/make-golden.py).
const X = require("./xml");
const { T, BX, BY, BW, BB, BH, T_H1, T_H2, T_BODY, T_SMALL, T_NOTE, T_KPI, bullets, band, panelHead, vchevron, source, rowsFill, colsFill } = X;
const req = (cond, msg) => { if (!cond) throw Error(msg); };
const arr = (v, lo, hi, name) => { req(Array.isArray(v) && v.length >= lo && v.length <= hi, `${name} must have ${lo}–${hi} items`); return v; };
const PRIORITY = { "高": T.ACCENT, "中": T.DARK, "低": T.MID, high: T.ACCENT, mid: T.DARK, low: T.MID };

module.exports = function register(def) {
  // ---- 21 issue list: priority badge × issue × understanding × hypothesis × analysis × owner ----
  def("21", {
    name: "論点整理表(優先バッジ×論点×現状×仮説×分析×担当。優先「高」だけ橙)",
    build(s, d) {
      const heads = arr(d.heads, 3, 7, "heads"), n = heads.length;
      const widths = d.widths || [34, ...Array.from({ length: n - 1 }, () => (BW - 34) / (n - 1))];
      let x = BX;
      heads.forEach((lab, j) => { band(s, x, BY, widths[j] - 4, 22, lab, { align: j === 0 ? "ctr" : "l", size: T_BODY }); x += widths[j]; });
      const rows = arr(d.rows, 2, 8, "rows");
      rowsFill(BY + 28, BB, rows.length, 4).forEach(([y, rh], i) => {
        const r = rows[i]; let x = BX;
        const col = PRIORITY[r.priority]; req(col, "priority must be 高|中|低");
        s.rect(x + 3, y + rh / 2 - 11, 24, 22, { fill: col, text: { content: r.priority, size: T_SMALL, bold: true, color: T.WHITE } });
        x += widths[0];
        req(r.cells.length === n - 1, "row needs " + (n - 1) + " cells");
        r.cells.forEach((v, j) => { s.text(x, y, widths[j + 1] - 4, rh, v, { size: T_SMALL + 0.5, bold: j === 0, anchor: "ctr", margins: [4, 1, 4, 1] }); x += widths[j + 1]; });
        s.line(BX + widths[0], y + rh + 2, BX + BW, y + rh + 2, { color: T.LIGHT });
      });
      source(s, d.source);
    },
  });

  // ---- 22 case study: profile card + issue → approach → effect + implication ----
  def("22", {
    name: "事例紹介 1 枚(概要カード+課題→アプローチ→効果+示唆)",
    build(s, d) {
      const pw = 190, p = d.profile;
      s.rect(BX, BY, pw, BH, { fill: T.PALE });
      s.text(BX + 10, BY + 8, pw - 20, 20, p.head, { size: T_H1, bold: true, margins: [0, 0, 0, 0] });
      s.line(BX + 10, BY + 30, BX + pw - 10, BY + 30, { color: T.INK, w: 1 });
      const rows = arr(p.rows, 3, 8, "profile rows");
      rowsFill(BY + 38, BB - 70, rows.length, 2).forEach(([yy, rh], i) => {
        s.text(BX + 10, yy, 56, rh, rows[i][0], { size: T_SMALL + 0.5, color: T.DARK, margins: [0, 0, 0, 0], anchor: "ctr" });
        s.text(BX + 66, yy, pw - 76, rh, rows[i][1], { size: T_BODY, margins: [0, 0, 0, 0], anchor: "ctr" });
        s.line(BX + 10, yy + rh + 1, BX + pw - 10, yy + rh + 1, { color: T.LIGHT });
      });
      s.rect(BX + 10, BB - 62, pw - 20, 54, { fill: T.WHITE });
      s.text(BX + 10, BB - 62, pw - 20, 54, [{ t: p.kpi.label, size: T_SMALL, color: T.DARK, align: "ctr" }, { t: p.kpi.value, size: 16, bold: true, color: T.ACCENT, align: "ctr" }], { anchor: "ctr" });
      const x0 = BX + pw + 16, panels = arr(d.panels, 2, 3, "panels"), cols = colsFill(x0, BX + BW - x0, panels.length, 26);
      const footH = d.implication ? 40 : 0;
      cols.forEach(([x, w], k) => {
        const y0 = panelHead(s, x, BY, w, panels[k].head);
        s.text(x, y0 + 6, w, BB - y0 - 6 - footH, bullets(panels[k].items, 12), { size: 12 });
        if (k < panels.length - 1) s.rect(x + w + 6, BY + 4, 14, 16, { fill: T.DARK, shape: "rightArrow" });
      });
      if (d.implication) { s.rect(x0, BB - 34, BX + BW - x0, 34, { fill: T.PALE }); s.text(x0 + 8, BB - 34, BX + BW - x0 - 16, 34, d.implication, { size: T_BODY, anchor: "ctr" }); }
      source(s, d.source);
    },
  });

  // ---- 23 project approach: phase chevrons + vertical chevron row heads ----
  def("23", {
    name: "プロジェクトアプローチ(フェーズ矢羽+縦矢羽の行見出し: 目的/活動/成果物/体制)",
    build(s, d) {
      const rw = 90, phases = arr(d.phases, 2, 5, "phases"), n = phases.length, gap = 8, x0 = BX + rw + 6, cw = (BX + BW - x0 - gap * (n - 1)) / n;
      phases.forEach((p, i) => {
        const x = x0 + i * (cw + gap);
        s.rect(x, BY, cw, 40, { fill: T.DARK, shape: i ? "chevron" : "pentagon", text: { content: [{ t: p.head, size: T_H1, bold: true, color: T.WHITE }, { t: p.time, size: T_SMALL, color: T.WHITE }], size: T_H1, bold: true, color: T.WHITE, margins: [14, 0, 8, 0] } });
      });
      const rows = arr(d.rows, 2, 5, "rows");
      rowsFill(BY + 48, BB, rows.length, 6).forEach(([y, rh], k) => {
        const r = rows[k];
        vchevron(s, BX, y, rw, rh, r.label, { size: T_BODY });
        req(r.cells.length === n, "row needs " + n + " cells");
        r.cells.forEach((c, i) => {
          const x = x0 + i * (cw + gap);
          if (Array.isArray(c)) s.text(x, y, cw, rh, bullets(c), { size: T_BODY, anchor: "ctr" });
          else s.text(x, y, cw, rh, c, { size: T_BODY, anchor: "ctr", margins: [6, 2, 6, 2] });
        });
        s.line(x0, y + rh + 3, BX + BW, y + rh + 3, { color: T.LIGHT, w: 0.75 });
      });
      source(s, d.source);
    },
  });

  // ---- 24 deliverables: 4 mock thumbnails + name/due/pages/contents + submission timeline ----
  def("24", {
    name: "成果物想定イメージ(4 つの模型+名称/提出/分量/内容+提出スケジュールの時間軸)",
    build(s, d) {
      const items = arr(d.items, 2, 5, "items"), fh = 128, schedH = d.schedule ? 90 : 0;
      colsFill(BX, BW, items.length, 16).forEach(([x, w], k) => {
        const it = items[k], fx = x + (w - 160) / 2, fy = BY;
        s.rect(fx, fy, 160, fh, { fill: T.WHITE, line: T.LIGHT, lw: 1 });
        s.rect(fx + 10, fy + 10, 100, 4, { fill: T.INK }); s.rect(fx + 10, fy + 18, 130, 3, { fill: T.LIGHT });
        if (it.kind === "chart") [40, 52, 30, 66, 58].forEach((hh, i) => s.rect(fx + 14 + i * 26, fy + 108 - hh, 16, hh, { fill: i === 3 ? T.DARK : T.MID }));
        else if (it.kind === "matrix") { for (let r = 0; r < 4; r++) for (let c = 0; c < 3; c++) s.rect(fx + 12 + c * 46, fy + 30 + r * 20, 42, 16, { fill: r === 0 ? T.DARK : T.PALE }); }
        else if (it.kind === "flow") { for (let i = 0; i < 4; i++) { s.rect(fx + 10 + i * 36, fy + 40, 30, 22, { fill: T.PALE, line: T.DARK, lw: 0.5 }); if (i < 3) s.line(fx + 40 + i * 36, fy + 51, fx + 46 + i * 36, fy + 51, { color: T.DARK, w: 0.75, arrow: true }); } for (let i = 0; i < 3; i++) s.rect(fx + 10, fy + 72 + i * 12, 130, 4, { fill: T.LIGHT }); }
        else [[0, 40], [20, 60], [50, 50], [70, 60], [90, 40]].forEach(([a, b], i) => s.rect(fx + 12 + a, fy + 32 + i * 16, b, 8, { fill: i === 0 || i === 3 ? T.DARK : T.MID, shape: "pentagon" }));
        s.text(x, BY + fh + 10, w, 22, it.name, { size: T_H2, bold: true, align: "ctr", margins: [0, 0, 0, 0] });
        s.line(x, BY + fh + 34, x + w, BY + fh + 34, { color: T.INK, w: 1.5 });
        s.text(x, BY + fh + 40, w, 30, [{ t: "提出: " + it.due, size: T_SMALL, color: T.DARK }, { t: "分量: " + it.pages, size: T_SMALL, color: T.DARK }], { margins: [0, 0, 0, 0] });
        s.text(x, BY + fh + 74, w, BB - BY - fh - 74 - schedH, bullets(it.items, 12), { size: 12 });
      });
      if (d.schedule) {
        const ty = BB - 80;
        band(s, BX, ty, BW, 20, d.schedule.head, { align: "l", size: T_BODY });
        const ly = ty + 50, tx0 = BX + 60, tx1 = BX + BW - 60;
        s.line(tx0, ly, tx1, ly, { color: T.DARK, w: 1.25 });
        const marks = arr(d.schedule.marks, 2, 8, "marks");
        marks.forEach((m, i) => {
          const f = m.f ?? i / (marks.length - 1), cx = tx0 + f * (tx1 - tx0);
          s.rect(cx - 5, ly - 5, 10, 10, { fill: m.where ? T.DARK : T.MID, shape: "diamond" });
          s.text(cx - 50, ly - 24, 100, 14, m.label, { size: T_SMALL, bold: true, align: "ctr", margins: [0, 0, 0, 0] });
          s.text(cx - 56, ly + 6, 112, 24, [{ t: m.what, size: 8.5 }, { t: m.where || "", size: 8, color: T.MUTED }], { align: "ctr", margins: [0, 0, 0, 0] });
        });
      }
      source(s, d.source);
    },
  });

  // ---- 25 purpose and goals: background → purpose (dark) → goals with KPI numbers ----
  def("25", {
    name: "目的とゴール(背景→目的(濃地)→ゴール 3 数値)",
    build(s, d) {
      const lw = 300, mw = 200, gap = 20;
      let y0 = panelHead(s, BX, BY, lw, d.background.head);
      s.text(BX, y0 + 8, lw, BB - y0 - 8, bullets(d.background.items), { size: T_BODY });
      const mx = BX + lw + gap, p = d.purpose;
      s.rect(mx, BY, mw, BH, { fill: T.DARK });
      s.text(mx + 12, BY + 16, mw - 24, 40, p.head, { size: T_H1, bold: true, color: T.WHITE, align: "ctr" });
      s.text(mx + 12, BY + 60, mw - 24, 120, p.text, { size: 15, color: T.WHITE, align: "ctr", anchor: "ctr" });
      s.line(mx + 30, BY + 190, mx + mw - 30, BY + 190, { color: T.WHITE, w: 0.75 });
      s.text(mx + 12, BY + 200, mw - 24, BH - 210, p.items.map((t) => ({ t, level: 0, color: T.WHITE, size: T_SMALL + 0.5 })), { color: T.WHITE });
      s.rect(mx + mw + 4, BY + BH / 2 - 10, 14, 20, { fill: T.DARK, shape: "rightArrow" });
      const gx = mx + mw + gap + 6, gw = BX + BW - gx;
      y0 = panelHead(s, gx, BY, gw, d.goals.head);
      const goals = arr(d.goals.items, 2, 4, "goals");
      rowsFill(y0 + 10, BB, goals.length, 8).forEach(([y, rh], i) => {
        const g = goals[i];
        s.rect(gx, y, gw, rh, { fill: T.PALE });
        s.text(gx + 10, y + 4, 90, 14, g.tag, { size: 8, color: T.MUTED, margins: [0, 0, 0, 0] });
        s.text(gx + 10, y + 16, 110, rh - 20, g.label, { size: T_H2, bold: true, anchor: "ctr", margins: [0, 0, 0, 0] });
        s.text(gx + 120, y, 90, rh, g.value, { size: T_KPI, bold: true, color: T.ACCENT, align: "ctr", anchor: "ctr", margins: [0, 0, 0, 0] });
        s.text(gx + 214, y, gw - 224, rh, g.desc, { size: T_SMALL + 0.5, color: T.DARK, anchor: "ctr", margins: [0, 0, 0, 0] });
      });
      source(s, d.source);
    },
  });
};
