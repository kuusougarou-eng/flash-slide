"use strict";
// Layouts 03–08 (ported from scripts/make-golden.py).
const X = require("./xml");
const { T, BX, BY, BW, BB, BH, T_H1, T_H2, T_BODY, T_SMALL, T_NOTE, bullets, band, panelHead, rowhead, harvey, source, rowsFill, colsFill } = X;
const req = (cond, msg) => { if (!cond) throw Error(msg); };
const arr = (v, lo, hi, name) => { req(Array.isArray(v) && v.length >= lo && v.length <= hi, `${name} must have ${lo}–${hi} items`); return v; };

/** small text grid (header row + rows) used in the right column of several layouts */
function miniTable(s, rx, ty, rw, head, rows, widths, opts = {}) {
  s.text(rx, ty, rw, 16, head, { size: T_H2, bold: true, margins: [0, 0, 0, 0] });
  s.line(rx, ty + 18, rx + rw, ty + 18, { color: T.INK, w: 0.75 });
  let yy = ty + 20;
  rows.forEach((r, i) => {
    let xx = rx;
    r.forEach((v, j) => {
      const numeric = opts.numeric ? opts.numeric(j) : j > 0;
      const bold = opts.bold ? opts.bold(i, j) : i === 0;
      const m = opts.margins ? opts.margins(j) : [0, 0, 0, 0];
      s.text(xx, yy, widths[j], 17, v, { size: T_SMALL, bold, color: i === 0 ? T.DARK : T.INK, align: numeric ? "r" : "l", margins: m, anchor: "ctr" });
      xx += widths[j];
    });
    s.line(rx, yy + 18, rx + rw, yy + 18, { color: T.PALE }); yy += 19;
  });
  return yy;
}
function autoDomain(values, pad = 0.15) {
  const lo = Math.min(...values), hi = Math.max(...values), span = Math.max(hi - lo, 1);
  const step = [0.5, 1, 2, 5, 10, 20, 50, 100, 200, 500, 1000].find((u) => span / u <= 6) || 1000;
  return { min: Math.floor((lo - span * pad) / step) * step, max: Math.ceil((hi + span * pad) / step) * step, majorUnit: step };
}
/** Keep a model-supplied axis only when it is readable (≤ 8 ticks) and contains every value with headroom. */
function usableDomain(domain, values) {
  if (!domain || !Number.isFinite(domain.min) || !Number.isFinite(domain.max) || domain.min >= domain.max) return autoDomain(values);
  const lo = Math.min(...values), hi = Math.max(...values), span = domain.max - domain.min;
  const unit = Number.isFinite(domain.majorUnit) && domain.majorUnit > 0 ? domain.majorUnit : null;
  if (lo < domain.min || hi > domain.max || (unit && span / unit > 8) || (hi - lo) / span < 0.25 || lo - domain.min < span * 0.04 || domain.max - hi < span * 0.04) return autoDomain(values);
  return unit ? domain : { ...domain, majorUnit: undefined };
}
module.exports = function register(def) {
  register.usableDomain = usableDomain;
  // ---- 03 trend: native line chart (2 series) + insights + mini table + so what ----
  def("03", {
    name: "推移+意味合い(ネイティブ折れ線 2 系列+注釈、右に箇条書き・小表・So what)",
    build(s, d) {
      const cw = 540;
      band(s, BX, BY, cw, 22, d.chartHead, { align: "l" });
      const series = arr(d.series, 1, 3, "series");
      const n = d.categories.length, domain = usableDomain(d.domain, series.flatMap((x) => x.values));
      s.chart(BX, BY + 26, cw, BB - BY - 26, { categories: d.categories, series: series.map((x, i) => ({ name: x.name, values: x.values, labelPoints: n >= 2 ? [n - 2, n - 1] : [n - 1] })), domain });
      if (d.annotation) {
        s.text(BX + 250, BY + 36, 230, 40, d.annotation.text, { size: T_SMALL, color: T.DARK });
        if (d.annotation.line) s.line(BX + 360, BY + 76, BX + 392, BY + 130, { color: T.DARK });
      }
      const rx = BX + cw + 24, rw = BX + BW - rx;
      band(s, rx, BY, rw, 22, d.insightsHead, { align: "l" });
      s.text(rx, BY + 28, rw, 150, bullets(d.insights));
      const ty = BY + 196;
      miniTable(s, rx, ty, rw, d.tableHead, d.table, [rw - 210, 70, 70, 70], { bold: (i, j) => i === 0 || j === 3 });
      s.rect(rx, BB - 44, rw, 44, { fill: T.PALE });
      s.text(rx + 6, BB - 44, rw - 12, 44, d.sowhat, { size: T_BODY, anchor: "ctr" });
      source(s, d.source);
    },
  });

  // ---- 04 waterfall: stacked column with transparent base + reading + recovery table ----
  def("04", {
    name: "要因分解(ウォーターフォール、最大要因だけ橙)+読み取り+回収見込みの小表",
    build(s, d) {
      const cw = 560;
      band(s, BX, BY, cw, 22, d.chartHead, { align: "l" });
      const cats = arr(d.categories, 3, 9, "categories"), vals = d.values;
      req(vals.length === cats.length, "values must match categories");
      const base = [], up = [], down = [], total = []; let run = 0;
      vals.forEach((v, i) => {
        if (i === 0 || i === vals.length - 1) { base.push(0); total.push(v); up.push(0); down.push(0); run = v; }
        else if (v < 0) { base.push(+(run + v).toFixed(6)); down.push(-v); up.push(0); total.push(0); run = +(run + v).toFixed(6); }
        else { base.push(run); up.push(v); down.push(0); total.push(0); run = +(run + v).toFixed(6); }
      });
      const inner = vals.slice(1, -1); let accent = -1, best = 0;
      inner.forEach((v, i) => { if (v < 0 && -v > best) { best = -v; accent = i + 1; } });
      const domain = d.domain || autoDomain([...base, ...vals.map((v, i) => base[i] + up[i] + down[i]), ...total.filter(Boolean)], 0.1);
      const chh = BB - BY - 26;
      s.chart(BX, BY + 26, cw, chh, { categories: cats, domain: { min: domain.min, max: domain.max }, series: [{ name: "base", values: base }, { name: "total", values: total }, { name: "減少", values: down, points: accent >= 0 ? [{ idx: accent, fill: T.ACCENT }] : [] }, { name: "増加", values: up }] });
      const slot = (cw - 50) / cats.length;
      vals.forEach((v, i) => {
        const lab = i === 0 || i === vals.length - 1 ? v.toFixed(1) : (v >= 0 ? "+" : "-") + Math.abs(v).toFixed(1);
        const top = Math.max(base[i] + up[i] + down[i], total[i]);
        const yy = BY + 26 + chh - 28 - (top - domain.min) / (domain.max - domain.min) * (chh - 40);
        s.text(BX + 30 + slot * i, yy - 15, slot, 14, lab, { size: T_SMALL, bold: true, align: "ctr", margins: [0, 0, 0, 0] });
      });
      const rx = BX + cw + 24, rw = BX + BW - rx;
      band(s, rx, BY, rw, 22, d.insightsHead, { align: "l" });
      s.text(rx, BY + 28, rw, 170, bullets(d.insights));
      const ty = BY + 218;
      const yy = miniTable(s, rx, ty, rw, d.tableHead, d.table, [74, 40, 40, rw - 154], { numeric: (j) => j === 1 || j === 2, margins: (j) => (j === 3 ? [6, 0, 0, 0] : j === 1 || j === 2 ? [0, 0, 4, 0] : [0, 0, 0, 0]) });
      if (d.footnote) s.text(rx, yy + 2, rw, 12, d.footnote, { size: 8, color: T.MUTED, margins: [0, 0, 0, 0] });
      source(s, d.source);
    },
  });

  // ---- 05 three panels: issue / hypothesis / approach rows + procedure band ----
  def("05", {
    name: "3パネル(論点/仮説/検証、行で揃えた3列+番号)+担当・期限、進め方の帯",
    build(s, d) {
      const cols = colsFill(BX, BW, 3, 20), heads = d.heads;
      let y0 = BY;
      cols.forEach(([x, w], i) => { y0 = panelHead(s, x, BY, w, heads[i]); });
      const items = arr(d.items, 2, 4, "items"), footH = d.band ? 70 : 0;
      rowsFill(y0 + 10, BB - footH - 8, items.length, 8).forEach(([y, rh], k) => {
        const it = items[k];
        let [x, w] = cols[0];
        s.rect(x, y, 32, 32, { fill: T.DARK, text: { content: String(k + 1), size: 14, bold: true, color: T.WHITE } });
        s.text(x + 40, y, w - 40, rh, [{ t: it.question, size: 13, bold: true }, { t: "担当: " + it.owner, size: T_SMALL + 0.5, color: T.DARK, sb: 6 }, { t: "結論の期限: " + it.due, size: T_SMALL + 0.5, color: T.DARK }], { anchor: "t" });
        [x, w] = cols[1];
        s.text(x, y, w, rh, [{ t: it.hypothesis, size: 12 }, { t: "検証後の使い方: " + it.use, size: T_SMALL + 0.5, color: T.DARK, sb: 6 }], { anchor: "t" });
        [x, w] = cols[2];
        s.text(x, y, w, rh, bullets(it.steps, 11.5), { size: 11.5, anchor: "t" });
        if (k < items.length - 1) s.line(BX, y + rh + 4, BX + BW, y + rh + 4, { color: T.LIGHT });
      });
      if (d.band) {
        const y = BB - footH;
        band(s, BX, y, BW, 22, d.band.head, { align: "l" });
        s.text(BX, y + 26, BW, footH - 26, bullets(d.band.items, T_BODY), { size: T_BODY });
      }
      source(s, d.source);
    },
  });

  // ---- 06 options comparison: option × criteria (harvey balls) + total row ----
  def("06", {
    name: "選択肢の比較(案×評価軸、ハーベイボール+数値、推奨列の帯・総合評価行)",
    build(s, d) {
      const hw = 190, options = arr(d.options, 2, 4, "options"), cw3 = (BW - hw) / options.length, rec = d.recommended;
      options.forEach((lab, j) => band(s, BX + hw + cw3 * j, BY, cw3 - 4, 24, lab, { fill: j === rec ? T.ACCENT : T.LIGHT, color: j === rec ? T.WHITE : T.INK }));
      const crit = arr(d.criteria, 3, 8, "criteria"), totH = 34, legH = 14;
      rowsFill(BY + 28, BB - totH - legH - 6, crit.length, 4).forEach(([y, rh], i) => {
        const c = crit[i];
        rowhead(s, BX, y, hw - 4, rh, c.label, { size: T_BODY });
        c.cells.forEach((cell, j) => {
          const cx0 = BX + hw + cw3 * j;
          if (j === rec) s.rect(cx0, y, cw3 - 4, rh, { fill: T.PALE });
          if (c.balls) { harvey(s, cx0 + 16, y + rh / 2, 14, c.balls[j]); s.text(cx0 + 28, y, cw3 - 34, rh, cell, { size: T_BODY, anchor: "ctr" }); }
          else s.text(cx0 + 6, y, cw3 - 14, rh, cell, { size: T_H2, anchor: "ctr" });
        });
        s.line(BX + hw, y + rh + 2, BX + BW, y + rh + 2, { color: T.LIGHT });
      });
      const y = BB - totH - legH - 2;
      rowhead(s, BX, y, hw - 4, totH, d.total.label, { fill: T.DARK, color: T.WHITE, size: T_BODY });
      d.total.cells.forEach((c, j) => {
        const cx0 = BX + hw + cw3 * j;
        s.rect(cx0, y, cw3 - 4, totH, { fill: j === rec ? T.ACCENT : T.PALE, text: { content: c, size: T_BODY, bold: true, color: j === rec ? T.WHITE : T.INK, align: "l", margins: [6, 1, 6, 1] } });
      });
      if (crit.some((c) => c.balls)) s.text(BX + 500, BB - legH, 396, legH, d.legend || "● 良い   ◕ やや良い   ◑ 中   ◔ やや悪い   ○ 悪い", { size: 8, color: T.MUTED, align: "r", margins: [0, 0, 0, 0] });
      source(s, d.source);
    },
  });

  // ---- 07 priority 2×2: bubble scatter (size = impact) + judgement + quadrant table ----
  def("07", {
    name: "優先順位 2×2(バブル散布、大きさ=効果)+判断の箇条書き+象限集計",
    build(s, d) {
      const px = BX + 36, py = BY + 6, pw = 560, ph = BB - BY - 26;
      s.rect(px, py, pw, ph, { line: T.LIGHT });
      s.line(px + pw / 2, py, px + pw / 2, py + ph, { color: T.LIGHT, dash: true });
      s.line(px, py + ph / 2, px + pw, py + ph / 2, { color: T.LIGHT, dash: true });
      s.text(px, py + ph + 4, pw, 14, d.axes.x, { size: T_SMALL, color: T.MUTED, align: "ctr", margins: [0, 0, 0, 0] });
      s.text(px - 30 - ph / 2 + 8, py + ph / 2 - 8, ph, 16, d.axes.y, { size: T_SMALL, color: T.MUTED, align: "ctr", anchor: "ctr", margins: [0, 0, 0, 0], rotation: 270 });
      const q = d.quadrants;
      [[q[0], px + 6, py + 4], [q[1], px + pw / 2 + 6, py + 4], [q[2], px + 6, py + ph - 18], [q[3], px + pw / 2 + 6, py + ph - 18]].forEach(([lab, xx, yy]) =>
        s.text(xx, yy, 200, 14, lab, { size: T_SMALL, bold: true, color: T.MUTED, margins: [0, 0, 0, 0] }));
      const pts = arr(d.points, 2, 16, "points"), maxImp = Math.max(...pts.map((p) => p.impact), 0.0001), unit = d.impactUnit ?? "";
      for (const p of pts) {
        const dd = 10 + 27 * (p.impact / maxImp), cx = px + p.x * pw, cy = py + (1 - p.y) * ph;
        s.rect(cx - dd / 2, cy - dd / 2, dd, dd, { fill: p.y > 0.5 && p.x < 0.5 ? T.ACCENT : p.y > 0.5 ? T.DARK : T.MID, shape: "oval" });
        const lab = p.impact >= 0.1 * maxImp - 1e-9 ? `${p.label} (${p.impact >= 0 ? "+" : "−"}${Math.abs(p.impact).toFixed(1)}${unit})` : p.label;
        if (p.x > 0.8) s.text(cx - dd / 2 - 132, cy - 8, 130, 16, lab, { size: T_SMALL, margins: [0, 0, 0, 0], anchor: "ctr", align: "r" });
        else s.text(cx + dd / 2 + 2, cy - 8, 130, 16, lab, { size: T_SMALL, margins: [0, 0, 0, 0], anchor: "ctr" });
      }
      const rx = px + pw + 24, rw = BX + BW - rx;
      band(s, rx, BY, rw, 22, d.insightsHead, { align: "l" });
      s.text(rx, BY + 28, rw, 200, bullets(d.insights));
      const ty = BY + 236;
      const yy = miniTable(s, rx, ty, rw, d.tableHead, d.table, [rw - 156, 46, 55, 55]);
      if (d.note) s.text(rx, yy + 4, rw, 12, d.note, { size: 8, color: T.MUTED, margins: [0, 0, 0, 0] });
      source(s, d.source);
    },
  });

  // ---- 08 n×m sentence grid: department × issue/cause/action/effect ----
  def("08", {
    name: "n×m 文章格子(行=部門、列=課題/原因/打ち手/効果。効果列は薄地+太字数値)",
    build(s, d) {
      const hw = 96, heads = arr(d.cols, 2, 5, "cols"), n = heads.length;
      const widths = d.widths || Array.from({ length: n }, () => (BW - hw) / n);
      req(widths.length === n && Math.abs(widths.reduce((a, b) => a + b, 0) - (BW - hw)) < 0.01, "widths must fill the body width");
      let x = BX + hw;
      heads.forEach((lab, j) => { band(s, x, BY, widths[j] - 4, 22, lab, { align: "l" }); x += widths[j]; });
      const rows = arr(d.rows, 2, 6, "rows"), effectCol = d.effectCol ?? n - 1;
      rowsFill(BY + 28, BB, rows.length, 6).forEach(([y, rh], i) => {
        const r = rows[i];
        rowhead(s, BX, y, hw - 4, rh, r.head, { size: T_H2, align: "ctr" });
        let x = BX + hw;
        r.cells.forEach((c, j) => {
          const w = widths[j];
          if (j === effectCol) { s.rect(x, y, w - 4, rh, { fill: T.PALE }); s.text(x, y, w - 4, rh, c, { size: T_BODY, align: "ctr", anchor: "ctr" }); }
          else s.text(x, y, w - 4, rh, c, { size: T_BODY, anchor: "ctr", margins: [6, 2, 6, 2] });
          x += w;
        });
        s.line(BX + hw, y + rh + 3, BX + BW, y + rh + 3, { color: T.LIGHT });
      });
      source(s, d.source);
    },
  });
};
