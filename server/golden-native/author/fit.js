"use strict";
// Static text fitting for non-golden data: estimate wrapped lines from a
// width-weighted character count and shrink the paragraph sizes step by step
// until the estimate fits the box. Never enlarges; never applied to golden data.
const { weightedLength } = require("../package");

function estimateHeight(content, base, w, m) {
  const items = Array.isArray(content) ? content : [content];
  let h = 0;
  for (const it of items) {
    const o = typeof it === "string" ? { t: it } : it;
    const size = o.size ?? base.size, ls = o.ls ?? base.ls ?? 1.15;
    const indent = o.level != null ? 12.6 + o.level * 14.2 : 0;
    const usable = Math.max(8, w - m[0] - m[2] - indent);
    for (const line of String(o.t).replace(/\*\*/g, "").split("\n")) {
      const n = weightedLength(line) * size * 1.0; // ~1em per CJK char, 0.55em per ASCII
      h += Math.max(1, Math.ceil(n / usable)) * size * ls + (o.sb || 0);
    }
  }
  return h + m[1] + m[3];
}
function scale(content, k) {
  const items = Array.isArray(content) ? content : [content];
  return items.map((it) => { const o = typeof it === "string" ? { t: it } : { ...it }; if (o.size != null) o.size = Math.max(7, +(o.size * k).toFixed(1)); return o; });
}
function fitter({ floor = 8, log } = {}) {
  return (content, base, w, h, m) => {
    if (estimateHeight(content, base, w, m) <= h + 0.5) return content;
    let k = 1, out = content, size = base.size;
    for (let i = 0; i < 12; i++) {
      k *= 0.92; size = base.size * k;
      if (size < floor) break;
      out = scale(content, k);
      if (estimateHeight(out, { ...base, size }, w, m) <= h + 0.5) { base.size = +size.toFixed(1); if (log) log({ k: +k.toFixed(2) }); return out; }
    }
    base.size = Math.max(floor, +size.toFixed(1));
    if (log) log({ k: +k.toFixed(2), overflow: true });
    return out;
  };
}
module.exports = { fitter, estimateHeight };
