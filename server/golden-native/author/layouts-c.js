"use strict";
// Layouts 10–15 (ported from scripts/make-golden.py).
const X = require("./xml");
const { T, BX, BY, BW, BB, BH, T_H1, T_H2, T_BODY, T_SMALL, T_NOTE, bullets, band, panelHead, rowhead, source, rowsFill, colsFill, pad2 } = X;
const req = (cond, msg) => { if (!cond) throw Error(msg); };
const arr = (v, lo, hi, name) => { req(Array.isArray(v) && v.length >= lo && v.length <= hi, `${name} must have ${lo}–${hi} items`); return v; };

module.exports = function register(def) {
  // ---- 10 horizontal enumeration: 2–4 panels + metric boxes + application band ----
  def("10", {
    name: "横列挙(短文): 4 パネル(黒下線)+箇条書き+各パネルの確認指標の箱+適用の帯",
    build(s, d) {
      const panels = arr(d.panels, 2, 4, "panels"), cols = colsFill(BX, BW, panels.length, 16);
      const footH = d.band ? 118 : 0, mbH = 54;
      cols.forEach(([x, w], k) => {
        const p = panels[k];
        const y0 = panelHead(s, x, BY, w, `${k + 1}. ${p.head}`);
        const hasMetric = !!p.metric;
        s.text(x, y0 + 8, w, BB - footH - y0 - 16 - (hasMetric ? mbH + 8 : 0), bullets(p.items, 12.5), { size: 12.5 });
        if (hasMetric) {
          const my = BB - footH - 8 - mbH;
          s.rect(x, my, w, mbH, { fill: T.PALE });
          s.text(x + 8, my, w - 16, mbH, [{ t: p.metric.label, size: T_SMALL, color: T.MUTED }, { t: p.metric.value, size: T_H2, bold: true }, { t: p.metric.target, size: T_SMALL, color: T.DARK }], { anchor: "ctr", margins: [0, 0, 0, 0] });
        }
      });
      if (d.band) {
        const y = BB - footH;
        band(s, BX, y, BW, 22, d.band.head, { fill: T.DARK, align: "l" });
        s.text(BX, y + 26, BW, footH - 26, bullets(d.band.items, 11.5), { size: 11.5 });
      }
      source(s, d.source);
    },
  });

  // ---- 11 process steps: chevrons + owner/input/process/output/KPI rows ----
  def("11", {
    name: "ステップ(矢羽): 3〜5 段の矢羽+段ごとの担当/入力/処理/出力/KPI の行",
    build(s, d) {
      const steps = arr(d.steps, 2, 6, "steps"), n = steps.length, gap = 6, sw = (BW - 110 - gap * (n - 1)) / n, x0 = BX + 110;
      steps.forEach((st, i) => {
        const x = x0 + i * (sw + gap);
        s.rect(x, BY, sw, 40, { fill: i === d.accentStep ? T.ACCENT : T.DARK, shape: i ? "chevron" : "pentagon", text: { content: [{ t: pad2(i + 1), size: 9, color: T.WHITE }, { t: st.head, size: T_H1, bold: true, color: T.WHITE }], size: T_H1, bold: true, color: T.WHITE, margins: [14, 0, 8, 0] } });
      });
      const rows = [{ label: d.ownerLabel || "担当 / 所要", cells: steps.map((st) => `${st.owner}\n${st.time}`), style: "owner" }, ...arr(d.rows, 1, 5, "rows")];
      rowsFill(BY + 48, BB, rows.length, 4).forEach(([y, rh], k) => {
        const r = rows[k];
        rowhead(s, BX, y, 104, rh, r.label, { size: T_BODY, align: "ctr" });
        r.cells.forEach((c, i) => {
          const x = x0 + i * (sw + gap);
          if (r.style === "kpi") { s.rect(x, y, sw, rh, { fill: T.PALE }); s.text(x, y, sw, rh, c, { size: T_BODY, bold: true, align: "ctr", anchor: "ctr" }); }
          else if (r.style === "owner") s.text(x, y, sw, rh, c, { size: T_BODY, color: T.DARK, align: "ctr", anchor: "ctr" });
          else s.text(x, y, sw, rh, c, { size: T_BODY, anchor: "ctr", margins: [6, 2, 6, 2] });
        });
        s.line(x0, y + rh + 2, BX + BW, y + rh + 2, { color: T.LIGHT });
      });
      source(s, d.source);
    },
  });

  // ---- 12 swimlane flow: lanes × columns, boxes/diamonds, routed arrows ----
  def("12", {
    name: "業務フロー(スイムレーン): 3〜5 レーン+箱+判断の菱形+矢印、所要と例外の注記",
    build(s, d) {
      const lanes = arr(d.lanes, 2, 5, "lanes"), lw = 80, lx = BX + lw, lwid = BW - lw;
      const lrows = rowsFill(BY, BB, lanes.length, 0);
      lrows.forEach(([y, h], i) => { rowhead(s, BX, y + 1, lw - 6, h - 2, lanes[i], { size: T_BODY, align: "ctr", fill: T.LIGHT }); s.line(lx, y + h, BX + BW, y + h, { color: T.LIGHT }); });
      s.rect(lx, BY, lwid, BH, { line: T.LIGHT });
      const bw = 104, bh = 36, ly = (i) => lrows[i][0] + lrows[i][1] / 2;
      const ncol = Math.max(...d.nodes.map((nd) => nd.col)) + 1, spacing = Math.min(130, (lwid - 140) / Math.max(1, ncol - 1));
      const xs = (c) => lx + 70 + spacing * c;
      const geo = {};
      for (const nd of arr(d.nodes, 2, 14, "nodes")) {
        const diamond = nd.shape === "diamond", w = nd.w || (diamond ? 96 : bw), h = nd.h || (diamond ? 52 : bh);
        const fill = nd.style === "pale" ? T.PALE : nd.style === "dark" ? T.DARK : T.WHITE;
        const cx = xs(nd.col), cy = ly(nd.lane);
        s.rect(cx - w / 2, cy - h / 2, w, h, { fill, line: fill === T.WHITE ? T.DARK : null, lw: 0.75, shape: diamond ? "diamond" : "rect", text: { content: nd.label, size: T_SMALL + 0.5, color: fill === T.DARK ? T.WHITE : T.INK, margins: [2, 0, 2, 0] } });
        geo[nd.id] = { cx, cy, w, h, lane: nd.lane, col: nd.col };
      }
      const arrow = (x1, y1, x2, y2) => s.line(x1, y1, x2, y2, { color: T.DARK, w: 1, arrow: true });
      const plain = (x1, y1, x2, y2) => s.line(x1, y1, x2, y2, { color: T.DARK, w: 1 });
      for (const e of d.edges || []) {
        const a = geo[e.from], b = geo[e.to];
        req(a && b, "edge references unknown node: " + e.from + "→" + e.to);
        let sx, sy, dir;
        const route = e.route || "h";
        if (route === "h") { const right = b.cx > a.cx; sx = a.cx + (right ? a.w / 2 : -a.w / 2); sy = a.cy; arrow(sx, sy, b.cx + (right ? -b.w / 2 : b.w / 2), b.cy); dir = "h"; }
        else if (route === "v") { const down = b.cy > a.cy; sx = a.cx; sy = a.cy + (down ? a.h / 2 : -a.h / 2); arrow(sx, sy, b.cx, b.cy + (down ? -b.h / 2 : b.h / 2)); dir = down ? "down" : "up"; }
        else if (route === "vh") { const down = b.cy > a.cy; sx = a.cx; sy = a.cy + (down ? a.h / 2 : -a.h / 2); plain(sx, sy, sx, b.cy); arrow(sx, b.cy, b.cx + (b.cx > a.cx ? -b.w / 2 : b.w / 2), b.cy); dir = down ? "down" : "up"; }
        else if (route === "hv") { const right = b.cx > a.cx; sx = a.cx + (right ? a.w / 2 : -a.w / 2); sy = a.cy; plain(sx, sy, b.cx, sy); const down = b.cy > a.cy; arrow(b.cx, sy, b.cx, b.cy + (down ? -b.h / 2 : b.h / 2)); dir = "h"; }
        else if (route === "hc") { const right = b.cx > a.cx; sx = a.cx + (right ? a.w / 2 : -a.w / 2); sy = a.cy; plain(sx, sy, b.cx, sy); dir = "h"; }
        else throw Error("Unknown edge route " + route);
        if (e.label) {
          const dx = e.labelDx ?? (dir === "h" ? 2 : 4), dy = e.labelDy ?? (dir === "down" ? 4 : -16);
          s.text(sx + dx, sy + dy, e.labelW || 60, 12, e.label, { size: 8, color: T.DARK, margins: [0, 0, 0, 0] });
        }
      }
      if (d.note) s.text(lx + 6, BB - 16, lwid - 12, 14, d.note, { size: 8, color: T.MUTED, margins: [0, 0, 0, 0] });
      source(s, d.source);
    },
  });

  // ---- 13 roadmap gantt: period groups, months, tasks with owners, milestones, today line ----
  def("13", {
    name: "ロードマップ(ガント: 年度帯・月・担当・節目・今日線)+凡例",
    build(s, d) {
      const lw = 170, ow = 96, gx = BX + lw, gw = BW - lw - ow;
      const periods = arr(d.periods, 2, 24, "periods"), pw = gw / periods.length;
      const groups = []; periods.forEach((p, i) => { const g = groups.at(-1); if (g && g.label === p.group) g.count++; else groups.push({ label: p.group, start: i, count: 1 }); });
      // 時間軸は年度=濃灰の帯、月=淡灰のマルチカラムで二段の見出しにする(1段だけの素の文字より版面が締まる)
      for (const g of groups) band(s, gx + pw * g.start, BY, pw * g.count - 2, 16, g.label, { fill: T.DARK, color: T.WHITE, size: T_SMALL });
      periods.forEach((p, i) => band(s, gx + pw * i, BY + 16, pw - 2, 18, p.label, { fill: T.LIGHT, color: T.INK, size: T_SMALL }));
      band(s, gx + gw, BY + 16, ow - 2, 18, d.ownerHead || "担当", { fill: T.LIGHT, color: T.INK, size: T_SMALL });
      s.line(BX, BY + 34, BX + BW, BY + 34, { color: T.DARK, w: 1 });
      const tasks = arr(d.tasks, 1, 16, "tasks"), noteH = d.note ? 16 : 0, msH = d.milestones?.length ? 24 : 0;
      rowsFill(BY + 36, BB - noteH - msH - 4, tasks.length, 0).forEach(([y, rh], i) => {
        const t = tasks[i];
        req(Number.isInteger(t.start) && Number.isInteger(t.end) && t.start >= 0 && t.end >= t.start && t.end < periods.length, "task interval out of range: " + t.label);
        s.text(BX, y, lw - 4, rh, t.label, { size: T_SMALL + 0.5, anchor: "ctr", margins: [4, 0, 0, 0] });
        s.rect(gx + pw * t.start + 2, y + 5, pw * (t.end - t.start + 1) - 4, rh - 10, { fill: t.critical ? T.ACCENT : T.MID, shape: "pentagon" });
        s.text(gx + gw, y, ow, rh, t.owner, { size: T_NOTE, color: T.DARK, anchor: "ctr", margins: [4, 0, 0, 0] });
        s.line(BX, y + rh, BX + BW, y + rh, { color: T.PALE });
      });
      const y = BB - noteH - msH - 2;
      if (msH) {
        s.text(BX, y, lw - 4, msH, d.milestoneHead || "節目", { size: T_SMALL + 0.5, bold: true, anchor: "ctr", margins: [4, 0, 0, 0] });
        for (const m of d.milestones) {
          const cx = gx + pw * m.at + pw / 2;
          s.rect(cx - 5, y + msH / 2 - 5, 10, 10, { fill: T.DARK, shape: "diamond" });
          s.text(cx + 6, y, 80, msH, m.label, { size: T_NOTE, anchor: "ctr", margins: [0, 0, 0, 0] });
        }
      }
      for (let i = 1; i < periods.length; i++) s.line(gx + pw * i, BY + 34, gx + pw * i, y + msH, { color: T.PALE });
      if (d.today) {
        const tx = gx + pw * d.today.at;
        s.line(tx, BY + 34, tx, y + msH, { color: T.DARK, w: 1, dash: true });
        s.rect(tx - 22, BY + 20, 44, 12, { fill: T.DARK, text: { content: d.today.label, size: 7.5, bold: true, color: T.WHITE } });
      }
      if (d.note) s.text(BX, BB - noteH + 2, BW, noteH, d.note, { size: T_NOTE, color: T.DARK, margins: [0, 0, 0, 0] });
      source(s, d.source);
    },
  });

  // ---- 14 organisation: committee → PMO (+external) → workstreams with KPI/tasks; meetings table ----
  def("14", {
    name: "推進体制: 体制図(委員会→PMO→分科会)+分科会ごとの KPI/担務+会議体の表+原則の帯",
    build(s, d) {
      const box = (x, y, w, h, t, sub, dark) => s.rect(x, y, w, h, { fill: dark ? T.DARK : T.PALE, text: { content: [{ t, size: T_H2, bold: true, color: dark ? T.WHITE : T.INK }, { t: sub, size: T_SMALL, color: dark ? T.WHITE : T.DARK }], size: T_H2, bold: true, color: dark ? T.WHITE : T.INK } });
      const ox = BX, ow = 560, bw = 180, bh = 44, cx = ox + ow / 2;
      box(cx - bw / 2, BY, bw, bh, d.top.name, d.top.sub, true);
      s.line(cx, BY + bh, cx, BY + bh + 14, { color: T.DARK, w: 1 });
      const y1 = BY + bh + 14;
      box(cx - bw / 2, y1, bw, bh, d.pmo.name, d.pmo.sub, false);
      if (d.external) { box(cx + bw / 2 + 30, y1, 150, bh, d.external.name, d.external.sub, false); s.line(cx + bw / 2, y1 + bh / 2, cx + bw / 2 + 30, y1 + bh / 2, { color: T.MID, w: 1, dash: true }); }
      const y2 = y1 + bh + 22;
      s.line(cx, y1 + bh, cx, y2 - 8, { color: T.DARK, w: 1 });
      const ws = arr(d.workstreams, 2, 4, "workstreams"), wcols = colsFill(ox, ow, ws.length, 12);
      s.line(wcols[0][0] + wcols[0][1] / 2, y2 - 8, wcols.at(-1)[0] + wcols.at(-1)[1] / 2, y2 - 8, { color: T.DARK, w: 1 });
      wcols.forEach(([x, ww], i) => { s.line(x + ww / 2, y2 - 8, x + ww / 2, y2, { color: T.DARK, w: 1 }); box(x, y2, ww, bh, ws[i].name, ws[i].sub, false); });
      const ty = y2 + bh + 10;
      wcols.forEach(([x, ww], i) => {
        s.text(x, ty, ww, 16, "KPI: " + ws[i].kpi, { size: T_SMALL + 0.5, bold: true, margins: [2, 0, 2, 0] });
        s.line(x, ty + 17, x + ww, ty + 17, { color: T.INK, w: 0.75 });
        s.text(x, ty + 20, ww, 56, bullets(ws[i].tasks), { size: T_SMALL, margins: [2, 0, 2, 0] });
      });
      const by = ty + 84;
      s.rect(ox, by, ow, BB - by, { fill: T.PALE });
      s.text(ox + 8, by, ow - 16, BB - by, d.summary.map((t, i) => (i ? { t, sb: 4 } : { t })), { size: T_BODY, anchor: "ctr" });
      const rx = ox + ow + 24, rw = BX + BW - rx;
      band(s, rx, BY, rw, 22, d.meetingsHead, { align: "l" });
      let y = BY + 28;
      for (const m of arr(d.meetings, 2, 5, "meetings")) {
        s.text(rx, y, 68, 40, m.name, { size: T_SMALL + 0.5, bold: true, margins: [0, 0, 0, 0] });
        s.text(rx + 70, y, 56, 40, m.cadence, { size: T_SMALL, color: T.DARK, margins: [0, 0, 0, 0] });
        s.text(rx + 128, y, rw - 128, 40, m.role, { size: T_SMALL, margins: [0, 0, 0, 0] });
        s.line(rx, y + 40, rx + rw, y + 40, { color: T.PALE }); y += 44;
      }
      s.rect(rx, y + 6, rw, BB - y - 6, { fill: T.PALE });
      s.text(rx + 6, y + 6, rw - 12, BB - y - 6, [{ t: "**" + d.principles.head + "**", size: T_H2 }, ...d.principles.items.map((t) => ({ t, sb: 4 }))], { size: T_SMALL + 0.5, anchor: "ctr" });
      source(s, d.source);
    },
  });

  // ---- 15 As-Is / To-Be: aspect × current × arrow × future + key points band ----
  def("15", {
    name: "As-Is / To-Be(観点×現状×矢印×将来)+移行の要点",
    build(s, d) {
      const hw = 110, aw = 356, ar = 40, tw = BW - hw - aw - ar;
      band(s, BX, BY, hw - 4, 24, d.heads[0], { fill: T.DARK });
      band(s, BX + hw, BY, aw - 4, 24, d.heads[1], { fill: T.LIGHT, color: T.INK, align: "l" });
      band(s, BX + hw + aw + ar, BY, tw, 24, d.heads[2], { fill: T.ACCENT, align: "l" });
      const rows = arr(d.rows, 3, 6, "rows"), kpH = d.band ? 76 : 0;
      rowsFill(BY + 30, BB - kpH - 8, rows.length, 6).forEach(([y, rh], i) => {
        const r = rows[i];
        rowhead(s, BX, y, hw - 4, rh, r.aspect, { size: T_H2, align: "ctr" });
        s.text(BX + hw, y, aw - 4, rh, r.asis, { size: T_BODY, anchor: "ctr", margins: [6, 2, 6, 2] });
        s.rect(BX + hw + aw + 8, y + rh / 2 - 7, ar - 16, 14, { fill: T.DARK, shape: "rightArrow" });
        s.rect(BX + hw + aw + ar, y, tw, rh, { fill: T.PALE });
        s.text(BX + hw + aw + ar, y, tw, rh, r.tobe, { size: T_BODY, anchor: "ctr", margins: [6, 2, 6, 2] });
        s.line(BX + hw, y + rh + 3, BX + hw + aw - 4, y + rh + 3, { color: T.LIGHT });
      });
      if (d.band) {
        const y = BB - kpH;
        band(s, BX, y, BW, 20, d.band.head, { fill: T.DARK, align: "l", size: T_BODY });
        s.text(BX, y + 24, BW, kpH - 24, bullets(d.band.items), { size: T_BODY });
      }
      source(s, d.source);
    },
  });
};
