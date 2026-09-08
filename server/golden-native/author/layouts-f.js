"use strict";
// Layouts 26–32 (ported from scripts/make-golden.py; consulting-pptx-skill catalogue types).
const X = require("./xml");
const { T, BX, BY, BW, BB, BH, T_H1, T_H2, T_BODY, T_SMALL, T_NOTE, bullets, band, rowhead, dot, source, rowsFill } = X;
const req = (cond, msg) => { if (!cond) throw Error(msg); };
const arr = (v, lo, hi, name) => { req(Array.isArray(v) && v.length >= lo && v.length <= hi, `${name} must have ${lo}–${hi} items`); return v; };
const TONE = { accent: T.ACCENT, dark: T.DARK, mid: T.MID };
const tone = (t) => { const c = TONE[t]; req(c, "tone must be accent|dark|mid"); return c; };
const SEV = { "高": T.ACCENT, "中": T.DARK, "低": T.MID };

module.exports = function register(def) {
  // ---- 26 decision page: ask + premises | numbered decisions with impact/reference/due ----
  def("26", {
    name: "意思決定ページ(依頼の大きな1文+前提 | 決定事項 3 件: 影響・参照・期限)",
    build(s, d) {
      const lw = 360;
      s.text(BX, BY, lw, 16, d.ask.label, { size: T_SMALL, bold: true, color: T.MUTED, margins: [0, 0, 0, 0] });
      s.text(BX, BY + 18, lw, 90, d.ask.text, { size: 20, bold: true });
      const y0 = BY + 116;
      s.text(BX, y0, lw, 16, d.premise.head, { size: T_SMALL, bold: true, color: T.DARK, margins: [0, 0, 0, 0] });
      s.line(BX, y0 + 18, BX + lw, y0 + 18, { color: T.INK, w: 0.75 });
      s.text(BX, y0 + 22, lw, BB - y0 - 22, bullets(d.premise.items), { size: T_BODY });
      const rx = BX + lw + 28, rw = BX + BW - rx, items = arr(d.decisions, 2, 4, "decisions");
      rowsFill(BY, BB, items.length, 8).forEach(([y, rh], i) => {
        const it = items[i];
        s.rect(rx, y, 30, 30, { fill: T.ACCENT, text: { content: "?", size: 15, bold: true, color: T.WHITE } });
        s.text(rx + 40, y, rw - 40, 40, it.ask, { size: 14, bold: true });
        s.text(rx + 40, y + 44, rw - 40, rh - 82, it.impact, { size: T_BODY, color: T.DARK });
        s.text(rx + 40, y + rh - 34, rw - 40, 16, it.ref, { size: T_SMALL, color: T.MUTED });
        s.text(rx + 40, y + rh - 16, rw - 40, 16, "期限: " + it.due, { size: T_SMALL, bold: true, color: T.DARK });
        s.line(rx, y + rh + 5, rx + rw, y + rh + 5, { color: T.LIGHT });
      });
      if (d.after) { s.rect(BX, BB - 34, lw, 34, { fill: T.PALE }); s.text(BX + 8, BB - 34, lw - 16, 34, d.after, { size: T_SMALL + 0.5, anchor: "ctr" }); }
      source(s, d.source);
    },
  });

  // ---- 27 risk table: severity badge × risk × early-warning sign × action ----
  def("27", {
    name: "リスク表(深刻度バッジ×リスク×兆候×打ち手。高=橙・中=濃灰・低=中灰)",
    build(s, d) {
      const heads = d.heads, widths = [56, 220, 260, BW - 56 - 220 - 260];
      let x = BX;
      heads.forEach((lab, j) => { band(s, x, BY, widths[j] - 4, 22, lab, { align: j === 0 ? "ctr" : "l" }); x += widths[j]; });
      const rows = arr(d.rows, 2, 6, "rows");
      rowsFill(BY + 28, BB, rows.length, 6).forEach(([y, rh], i) => {
        const r = rows[i], col = SEV[r.severity]; req(col, "severity must be 高|中|低");
        let x = BX;
        s.rect(x + 8, y + rh / 2 - 11, 40, 22, { fill: col, text: { content: r.severity, size: T_SMALL, bold: true, color: T.WHITE } }); x += 56;
        s.text(x, y, 216, rh, r.risk, { size: T_H2, bold: true, anchor: "ctr", margins: [4, 2, 4, 2] }); x += 220;
        s.text(x, y, 256, rh, r.sign, { size: T_BODY, anchor: "ctr", margins: [4, 2, 4, 2] }); x += 260;
        s.text(x, y, widths[3] - 4, rh, r.action, { size: T_BODY, anchor: "ctr", margins: [4, 2, 4, 2] });
        s.line(BX + 56, y + rh + 3, BX + BW, y + rh + 3, { color: T.LIGHT });
      });
      source(s, d.source);
    },
  });

  // ---- 28 scenario table: scenario × premise × result × probability + decision band ----
  def("28", {
    name: "シナリオ表(楽観/基本/悲観×前提×結果(帯で塗り分け)+発生確率+投資判断の前提の帯)",
    build(s, d) {
      const widths = [130, 460, 140, BW - 130 - 460 - 140];
      let x = BX;
      d.heads.forEach((lab, j) => { band(s, x, BY, widths[j] - 4, 24, lab, { align: j === 0 ? "ctr" : "l" }); x += widths[j]; });
      const rows = arr(d.rows, 2, 4, "rows"), bandH = d.band ? 46 : 0;
      rowsFill(BY + 30, BB - bandH, rows.length, 10).forEach(([y, rh], i) => {
        const r = rows[i], col = tone(r.tone);
        let x = BX;
        s.rect(x, y, 126, rh, { fill: col, text: { content: r.name, size: T_H2, bold: true, color: T.WHITE } }); x += 130;
        s.text(x, y, 456, rh, r.premise, { size: T_BODY, anchor: "ctr", margins: [6, 2, 6, 2] }); x += 460;
        s.text(x, y, 136, rh, r.result, { size: 20, bold: true, color: col, align: "ctr", anchor: "ctr" }); x += 140;
        s.text(x, y, widths[3] - 4, rh, r.probability, { size: T_H2, align: "ctr", anchor: "ctr" });
      });
      if (d.band) { const y = BB - 40; s.rect(BX, y, BW, 40, { fill: T.PALE }); s.text(BX + 8, y, BW - 16, 40, d.band, { size: T_BODY, anchor: "ctr" }); }
      source(s, d.source);
    },
  });

  // ---- 29 status heatmap: metric × period, 5 shades + "what to look at next" column ----
  def("29", {
    name: "状態ヒートマップ(指標×月の 5 段階濃淡+「だから、次に見るべき点」の右カラム)",
    build(s, d) {
      const hw = 140, months = arr(d.periods, 2, 12, "periods"), cw = (BW - hw - 260) / months.length;
      band(s, BX, BY, hw - 4, 22, d.heads[0], { align: "l" });
      months.forEach((m, i) => band(s, BX + hw + cw * i, BY, cw - 4, 22, m, { fill: T.LIGHT, color: T.INK, size: T_BODY }));
      band(s, BX + hw + cw * months.length + 8, BY, 252, 22, d.heads[1], { align: "l" });
      const levels = [T.PALE, T.LIGHT, T.MID, T.GRAY3, T.DARK], labels = d.levelLabels || ["横ばい", "やや改善", "改善", "改善", "大きく改善"];
      const rows = arr(d.rows, 2, 6, "rows");
      rowsFill(BY + 28, BB, rows.length, 6).forEach(([y, rh], i) => {
        const r = rows[i];
        rowhead(s, BX, y, hw - 4, rh, r.label, { size: T_BODY });
        req(r.values.length === months.length, "values must match periods");
        r.values.forEach((v, j) => {
          req(Number.isInteger(v) && v >= 0 && v <= 4, "heat levels are integers 0–4");
          s.rect(BX + hw + cw * j, y, cw - 4, rh, { fill: levels[v] });
          s.text(BX + hw + cw * j, y, cw - 4, rh, labels[v], { size: T_NOTE, color: v >= 3 ? T.WHITE : T.INK, align: "ctr", anchor: "ctr", margins: [0, 0, 0, 0] });
        });
      });
      const cx = BX + hw + cw * months.length + 8;
      s.text(cx, BY + 28, 252, BB - BY - 28, bullets(d.insights), { size: T_BODY });
      if (d.legend) s.text(BX, BB - 12, 400, 12, d.legend, { size: 8, color: T.MUTED, margins: [0, 0, 0, 0] });
      source(s, d.source);
    },
  });

  // ---- 30 proportional circles: two values as area-scaled circles + breakdown table ----
  def("30", {
    name: "比例円の対比(2 値を面積比の円で対比+内訳の箇条書き+区分別の内訳表)",
    build(s, d) {
      band(s, BX, BY, 300, 20, d.head, { align: "l" });
      const items = arr(d.items, 2, 2, "items"), cx0 = BX + 160, gap = 220, maxd = 130, top0 = BY + 40;
      const scale = maxd / Math.sqrt(Math.max(...items.map((it) => it.value)));
      items.forEach((it, i) => {
        const cx = cx0 + i * gap, dd = Math.sqrt(it.value) * scale, col = tone(it.tone);
        dot(s, cx, top0 + maxd / 2, dd, col);
        s.text(cx - 90, BY + 24, 180, 16, it.head, { size: T_H2, bold: true, align: "ctr", margins: [0, 0, 0, 0] });
        s.text(cx - 60, top0 + maxd / 2 - 16, 120, 32, it.label, { size: 22, bold: true, color: col !== T.MID ? T.WHITE : T.INK, align: "ctr", anchor: "ctr", margins: [0, 0, 0, 0] });
        s.text(cx - 90, top0 + maxd + 6, 180, 14, d.unit, { size: T_SMALL, color: T.MUTED, align: "ctr", margins: [0, 0, 0, 0] });
      });
      s.rect(cx0 + gap / 2 - 10, top0 + maxd / 2 - 8, 20, 16, { fill: T.DARK, shape: "rightArrow" });
      const rx = cx0 + gap * 2 - 20, rw = BX + BW - rx;
      band(s, rx, BY, rw, 20, d.insightsHead, { align: "l" });
      s.text(rx, BY + 28, rw, top0 + maxd - BY, bullets(d.insights), { size: T_BODY });
      const ay = top0 + maxd + 32;
      band(s, BX, ay, BW, 22, d.tableHead, { align: "l" });
      const cols = d.tableCols, widths = [220, 200, 220, BW - 220 - 200 - 220];
      let x = BX;
      cols.forEach((lab, j) => { s.text(x, ay + 24, widths[j] - 4, 16, lab, { size: T_SMALL, bold: true, color: T.DARK, margins: [0, 0, 0, 0] }); x += widths[j]; });
      s.line(BX, ay + 42, BX + BW, ay + 42, { color: T.INK, w: 0.75 });
      const rows = arr(d.table, 2, 4, "table");
      rowsFill(ay + 46, BB, rows.length, 4).forEach(([y, rh], i) => {
        const r = rows[i];
        s.text(BX, y, 216, rh, r[0], { size: T_BODY, bold: true, anchor: "ctr", margins: [0, 0, 0, 0] });
        s.text(BX + 220, y, 196, rh, r[1], { size: T_BODY, anchor: "ctr", margins: [0, 0, 0, 0] });
        s.text(BX + 420, y, 216, rh, r[2], { size: T_BODY, bold: true, color: T.ACCENT, anchor: "ctr", margins: [0, 0, 0, 0] });
        s.text(BX + 636, y, BW - 636, rh, r[3], { size: T_SMALL + 0.5, color: T.DARK, anchor: "ctr", margins: [0, 0, 0, 0] });
        s.line(BX, y + rh + 2, BX + BW, y + rh + 2, { color: T.LIGHT });
      });
      source(s, d.source);
    },
  });

  // ---- 31 ranked bars with target line and annotation ----
  def("31", {
    name: "順位棒+注記(降順の横棒、維持/対象で色分け+目標線+読み取り)",
    build(s, d) {
      const cw = 560;
      band(s, BX, BY, cw, 20, d.head, { align: "l" });
      const bars = arr(d.bars, 2, 10, "bars"), maxw = cw - 200, max = d.max ?? 100;
      rowsFill(BY + 30, BB - 20, bars.length, 6).forEach(([y, rh], i) => {
        const b = bars[i];
        req(Number.isFinite(b.value) && b.value >= 0 && b.value <= max, "bar value out of range: " + b.name);
        s.text(BX, y, 60, rh, b.name, { size: T_BODY, bold: true, anchor: "ctr", margins: [0, 0, 0, 0] });
        const bw = maxw * b.value / max;
        s.rect(BX + 64, y + rh * 0.2, bw, rh * 0.6, { fill: b.keep ? T.ACCENT : T.MID });
        s.text(BX + 64 + bw + 6, y, 60, rh, `${b.value}${d.unit || ""}`, { size: T_BODY, bold: true, anchor: "ctr", margins: [0, 0, 0, 0] });
        if (!b.keep && d.targetTag) s.text(BX + 64 + maxw + 70, y, 100, rh, d.targetTag, { size: T_SMALL, color: T.DARK, anchor: "ctr", margins: [0, 0, 0, 0] });
      });
      if (d.target) {
        const tx = BX + 64 + maxw * d.target.value / max;
        s.line(tx, BY + 24, tx, BB - 20, { color: T.DARK, w: 1, dash: true });
        s.text(tx - 40, BY + 24, 80, 14, d.target.label, { size: 8, color: T.DARK, align: "ctr", margins: [0, 0, 0, 0] });
      }
      if (d.legend) s.text(BX, BB - 16, cw, 12, d.legend, { size: 8, color: T.MUTED, margins: [0, 0, 0, 0] });
      const rx = BX + cw + 24, rw = BX + BW - rx;
      band(s, rx, BY, rw, 20, d.insightsHead, { align: "l" });
      s.text(rx, BY + 28, rw, BB - BY - 28, bullets(d.insights), { size: T_BODY });
      source(s, d.source);
    },
  });

  // ---- 32 scenario lines: native line chart (3 scenarios) + per-scenario implication ----
  def("32", {
    name: "シナリオ線(ネイティブ折れ線 3 シナリオ+シナリオごとの年平均変化率+含意)",
    build(s, d) {
      const cw = 560;
      band(s, BX, BY, cw, 20, d.chartHead, { align: "l" });
      const series = arr(d.series, 2, 4, "series");
      s.chart(BX, BY + 26, cw, BB - BY - 26, { categories: d.categories, series: series.map((x) => ({ name: x.name, values: x.values })), domain: require("./layouts-b").usableDomain(d.domain, series.flatMap((x) => x.values)) });
      const rx = BX + cw + 24, rw = BX + BW - rx;
      band(s, rx, BY, rw, 20, d.insightsHead, { align: "l" });
      const rows = arr(d.rows, 2, 4, "rows");
      rowsFill(BY + 28, BB, rows.length, 8).forEach(([y, rh], i) => {
        const r = rows[i], col = tone(r.tone);
        dot(s, rx + 7, y + 10, 10, col);
        s.text(rx + 18, y, rw - 90, 18, r.name, { size: T_H2, bold: true, margins: [0, 0, 0, 0] });
        s.text(rx + rw - 72, y, 72, 18, r.rate, { size: T_BODY, bold: true, color: col !== T.DARK ? col : T.INK, align: "r", margins: [0, 0, 0, 0] });
        s.text(rx + 18, y + 20, rw - 18, rh - 20, r.note, { size: T_SMALL + 0.5, color: T.DARK, margins: [0, 0, 0, 0] });
      });
      source(s, d.source);
    },
  });
};
