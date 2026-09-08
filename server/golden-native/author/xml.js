"use strict";
// DrawingML emitter that reproduces python-pptx's output shape for the golden
// generator (scripts/make-golden.py). Coordinates are points; EMU = round(pt*12700)
// with Python's round-half-even so the XML matches the reference byte for byte.
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

const T = { ACCENT: "D04A02", INK: "1A1A1A", DARK: "404040", MUTED: "6E6E6E", MID: "A6A6A6", LIGHT: "D9D9D9", PALE: "F2F2F2", WHITE: "FFFFFF", GREEN: "2E7D32", AMBER: "E09A00", RED: "C62828", GRAY3: "8A8A8A" };
const FONT = "Arial";
const BX = 32, BY = 128, BW = 896, BB = 488, BH = BB - BY, SRC_Y = 494;
const T_H1 = 14, T_H2 = 12.5, T_BODY = 11.5, T_SMALL = 9.5, T_NOTE = 8.5, T_KPI = 30;

function pyRound(v) { // Python 3 round(): half to even
  const f = Math.floor(v), d = v - f;
  if (d > 0.5) return f + 1;
  if (d < 0.5) return f;
  return f % 2 === 0 ? f : f + 1;
}
const emu = (pt) => String(pyRound(pt * 12700));
const sz = (pt) => String(Math.round(pt * 100));
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const PRST = { rect: "rect", oval: "ellipse", pentagon: "homePlate", chevron: "chevron", diamond: "diamond", pie: "pie", rightArrow: "rightArrow" };
const NAMES = { rect: "Rectangle", oval: "Oval", pentagon: "Pentagon", chevron: "Chevron", diamond: "Diamond", pie: "Pie", rightArrow: "Right Arrow" };

function rPr(size, bold, color, name, italic) {
  return `<a:rPr sz="${sz(size)}" b="${bold ? 1 : 0}" i="${italic ? 1 : 0}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="${name}"/></a:rPr>`;
}
function runs(s, size, bold, color, name) {
  return String(s).split("**").map((seg, i) => seg ? `<a:r>${rPr(size, bold || i % 2 === 1, color, name)}<a:t>${esc(seg)}</a:t></a:r>` : "").join("");
}
function pPr(o) {
  // o: {align, ls, sb, level, char} ; level null => no bullet
  const bullet = o.level != null;
  const marL = bullet ? 160000 + o.level * 180000 : 0, indent = bullet ? -160000 : 0;
  let s = `<a:pPr algn="${o.align}" marL="${marL}" indent="${indent}">`;
  if (o.ls != null) s += `<a:lnSpc><a:spcPct val="${Math.round(o.ls * 100000)}"/></a:lnSpc>`;
  if (o.sb) s += `<a:spcBef><a:spcPts val="${Math.round(o.sb * 100)}"/></a:spcBef>`;
  s += bullet ? `<a:buChar char="${esc(o.char || (o.level === 0 ? "•" : "–"))}"/>` : "<a:buNone/>";
  return s + "</a:pPr>";
}
function paragraphs(content, base) {
  const items = Array.isArray(content) ? content : [content];
  return items.map((it) => {
    const o = typeof it === "string" ? { t: it } : it;
    const align = o.align || base.align, size = o.size ?? base.size, bold = o.bold ?? base.bold, color = o.color || base.color;
    return `<a:p>${pPr({ align, ls: o.ls ?? base.ls, sb: o.sb, level: o.level, char: o.char })}${runs(o.t, size, bold, color, base.name || FONT)}</a:p>`;
  }).join("");
}
const STYLE_SP = `<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>`;
const STYLE_CXN = `<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>`;

class Slide {
  constructor(opts = {}) {
    this.shapes = []; this.nextId = 5; this.charts = []; this.fit = null; this.warnings = [];
    if (opts.page) this.text(919, 513, 9, 9, String(opts.page), { size: 7.5, margins: [0, 0, 0, 0], align: "r" });
    this.fit = opts.fit || null;
  }
  id(kind) { const id = this.nextId++; return { id, name: `FS_${kind} ${id - 1}` }; }
  xfrm(x, y, w, h, rot) {
    return `<a:xfrm${rot ? ` rot="${Math.round(rot * 60000)}"` : ""}><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>`;
  }
  /** text box. content: string | array of string | {t, level, size, bold, color, align, sb, char, ls} */
  text(x, y, w, h, content, o = {}) {
    const base = { size: o.size ?? T_BODY, bold: !!o.bold, color: o.color || T.INK, align: o.align || "l", ls: o.ls ?? 1.15, name: o.name || FONT };
    const m = o.margins || [3, 2, 3, 2], anchor = o.anchor || "t";
    if (this.fit) content = this.fit(content, base, w, h, m);
    const { id, name } = this.id("TextBox");
    this.shapes.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${this.xfrm(x, y, w, h, o.rotation)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" anchor="${anchor}" lIns="${emu(m[0])}" tIns="${emu(m[1])}" rIns="${emu(m[2])}" bIns="${emu(m[3])}"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${paragraphs(content, base)}</p:txBody></p:sp>`);
    return { id, x, y, w, h };
  }
  /** autoshape. o: {fill, line, lw, shape, rotation, adj:[..], text:{content,size,bold,color,align,anchor,margins}} */
  rect(x, y, w, h, o = {}) {
    const kind = o.shape || "rect", { id, name } = this.id(NAMES[kind]);
    const av = o.adj ? `<a:avLst>${o.adj.map((v, i) => `<a:gd name="adj${i + 1}" fmla="val ${Math.round(v * 100000)}"/>`).join("")}</a:avLst>` : "<a:avLst/>";
    const fill = o.fill ? `<a:solidFill><a:srgbClr val="${o.fill}"/></a:solidFill>` : "<a:noFill/>";
    const line = o.line ? `<a:ln w="${emu(o.lw ?? 0.75)}"><a:solidFill><a:srgbClr val="${o.line}"/></a:solidFill></a:ln>` : "<a:ln><a:noFill/></a:ln>";
    // python-pptx add_shape keeps its default centred paragraph; rect() clears it to <a:p/>
    let body = `<a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>${o.defaultParagraph ? '<a:p><a:pPr algn="ctr"/></a:p>' : "<a:p/>"}`;
    if (o.text) {
      const t = o.text, m = t.margins || [4, 1, 4, 1];
      const base = { size: t.size, bold: !!t.bold, color: t.color || T.INK, align: t.align || "ctr", ls: 1.1, name: FONT };
      let content = t.content;
      if (this.fit) content = this.fit(content, base, w, h, m);
      body = `<a:bodyPr rtlCol="0" anchor="${t.anchor || "ctr"}" wrap="square" lIns="${emu(m[0])}" tIns="${emu(m[1])}" rIns="${emu(m[2])}" bIns="${emu(m[3])}"/><a:lstStyle/>${paragraphs(content, base)}`;
    }
    this.shapes.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${this.xfrm(x, y, w, h, o.rotation)}<a:prstGeom prst="${PRST[kind]}">${av}</a:prstGeom>${fill}${line}<a:effectLst/></p:spPr>${STYLE_SP}<p:txBody>${body}</p:txBody></p:sp>`);
    return { id, x, y, w, h };
  }
  line(x1, y1, x2, y2, o = {}) {
    const { id, name } = this.id("Connector");
    const flipH = x2 < x1, flipV = y2 < y1;
    // python-pptx rounds each endpoint to EMU before taking the extent
    const ex1 = pyRound(x1 * 12700), ex2 = pyRound(x2 * 12700), ey1 = pyRound(y1 * 12700), ey2 = pyRound(y2 * 12700);
    const x = Math.min(ex1, ex2), y = Math.min(ey1, ey2), w = Math.abs(ex2 - ex1), h = Math.abs(ey2 - ey1);
    const flips = `${flipH ? ' flipH="1"' : ""}${flipV ? ' flipV="1"' : ""}`;
    const ln = `<a:ln w="${emu(o.w ?? 0.75)}"><a:solidFill><a:srgbClr val="${o.color || T.LIGHT}"/></a:solidFill>${o.dash ? '<a:prstDash val="dash"/>' : ""}${o.arrow ? '<a:tailEnd type="triangle" w="med" len="med"/>' : ""}</a:ln>`;
    this.shapes.push(`<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm${flips}><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="${o.elbow ? "bentConnector3" : "line"}"><a:avLst/></a:prstGeom>${ln}</p:spPr>${STYLE_CXN}</p:cxnSp>`);
  }
  /** native chart placeholder; data is bound to the container's chart part in order */
  chart(x, y, w, h, data) {
    const { id, name } = this.id("Chart"), index = this.charts.length;
    this.charts.push(data);
    this.shapes.push(`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></p:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="__CHART_${index}__"/></a:graphicData></a:graphic></p:graphicFrame>`);
  }
  /** native table: cells[r][c] = {t, size, bold, color, fill, align, span, merged} ; widths/heights in pt */
  table(x, y, w, h, widths, heights, cells) {
    const { id, name } = this.id("Table");
    // python-pptx sets the frame height to the sum of the rounded row heights
    const cy = heights.reduce((n, rh) => n + pyRound(rh * 12700), 0);
    const tr = cells.map((row, i) => `<a:tr h="${emu(heights[i])}">` + row.map((c) => {
      if (c.merged) return `<a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/>${c.raw ? "<a:p/>" : `<a:p><a:pPr algn="${c.align || "l"}"/></a:p>`}</a:txBody>${tcPr(c, i)}</a:tc>`;
      const span = c.span > 1 ? ` gridSpan="${c.span}"` : "";
      const p = `<a:p><a:pPr algn="${c.align || "l"}"/>${runs(c.t, c.size ?? T_BODY, !!c.bold, c.color || T.INK, FONT)}</a:p>`;
      return `<a:tc${span}><a:txBody><a:bodyPr/><a:lstStyle/>${p}</a:txBody>${tcPr(c, i)}</a:tc>`;
    }).join("") + "</a:tr>").join("");
    this.shapes.push(`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="0" bandRow="0" lastRow="0" firstCol="0" lastCol="0" bandCol="0"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid>${widths.map((cw) => `<a:gridCol w="${emu(cw)}"/>`).join("")}</a:tblGrid>${tr}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`);
  }
  render({ title, lead, titleSize = 20, leadSize = 13 }) {
    const t = `<p:sp><p:nvSpPr><p:cNvPr id="3" name="FS_Title 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr>${this.xfrm(32, 30, 896, 58)}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="${sz(titleSize)}"/><a:t>${esc(title)}</a:t></a:r></a:p></p:txBody></p:sp>`;
    const l = `<p:sp><p:nvSpPr><p:cNvPr id="4" name="FS_Subtitle 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="16"/></p:nvPr></p:nvSpPr><p:spPr>${this.xfrm(32, 92, 896, 32)}</p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr sz="${sz(leadSize)}"><a:solidFill><a:srgbClr val="${T.DARK}"/></a:solidFill><a:latin typeface="${FONT}"/></a:rPr><a:t>${esc(lead)}</a:t></a:r></a:p></p:txBody></p:sp>`;
    return `<p:sld xmlns:a="${A}" xmlns:p="${P}" xmlns:r="${R}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${t}${l}${this.shapes.join("")}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
  }
}
function tcPr(c, rowIndex) {
  const fill = c.fill ? `<a:solidFill><a:srgbClr val="${c.fill}"/></a:solidFill>` : "<a:noFill/>";
  const bottom = rowIndex > 0 ? `<a:lnB w="6350"><a:solidFill><a:srgbClr val="${T.LIGHT}"/></a:solidFill></a:lnB>` : `<a:lnB w="0"><a:noFill/></a:lnB>`;
  if (c.raw) return `<a:tcPr><a:lnL w="0"><a:noFill/></a:lnL><a:lnR w="0"><a:noFill/></a:lnR><a:lnT w="0"><a:noFill/></a:lnT>${bottom}</a:tcPr>`;
  return `<a:tcPr marL="${emu(4)}" marR="${emu(4)}" marT="${emu(1)}" marB="${emu(1)}" anchor="ctr">${fill}<a:lnL w="0"><a:noFill/></a:lnL><a:lnR w="0"><a:noFill/></a:lnR><a:lnT w="0"><a:noFill/></a:lnT>${bottom}</a:tcPr>`;
}

// ---- composite helpers mirroring make-golden.py ----
function bullets(items, size = T_BODY) {
  return items.map((it) => {
    if (Array.isArray(it)) { const [t, lv] = it; return { t, level: lv, size: size - (lv ? 0.5 : 0) }; }
    if (it.startsWith("#")) return { t: it.slice(1), bold: true, size: T_H2, sb: 6 };
    return { t: it, level: 0, size };
  });
}
function band(s, x, y, w, h, label, o = {}) {
  return s.rect(x, y, w, h, { fill: o.fill || T.DARK, text: { content: label, size: o.size ?? T_H2, bold: o.bold ?? true, color: o.color || T.WHITE, align: o.align || "ctr" } });
}
function panelHead(s, x, y, w, label, o = {}) {
  s.text(x, y, w, 22, label, { size: o.size ?? T_H1, bold: true, align: "ctr", anchor: "b", margins: [0, 0, 0, 2] });
  s.line(x, y + 24, x + w, y + 24, { color: T.INK, w: o.lw ?? 1.5 });
  return y + 24;
}
function rowhead(s, x, y, w, h, label, o = {}) {
  return s.rect(x, y, w, h, { fill: o.fill || T.LIGHT, text: { content: label, size: o.size ?? T_BODY, bold: true, color: o.color || T.INK, align: o.align || "l", margins: [6, 1, 6, 1] } });
}
function vchevron(s, x, y, w, h, label, o = {}) {
  s.rect(x + (w - h) / 2, y + (h - w) / 2, h, w, { fill: o.fill || T.LIGHT, shape: "pentagon", rotation: 90 });
  s.text(x, y, w, h, label, { size: o.size ?? T_BODY, bold: true, align: "ctr", anchor: "ctr" });
}
function harvey(s, cx, cy, d, level) {
  const x = cx - d / 2, y = cy - d / 2;
  if (level >= 4) { s.rect(x, y, d, d, { fill: T.DARK, line: T.DARK, shape: "oval" }); return; }
  s.rect(x, y, d, d, { fill: T.WHITE, line: T.DARK, lw: 1, shape: "oval" });
  if (level <= 0) return;
  s.rect(x, y, d, d, { fill: T.DARK, shape: "pie", adj: [270 * 0.6, ((270 + 90 * level) % 360) * 0.6], defaultParagraph: true });
}
function dot(s, cx, cy, d, color) { return s.rect(cx - d / 2, cy - d / 2, d, d, { fill: color, shape: "oval" }); }
function source(s, t) { s.text(BX, SRC_Y, BW, 12, t, { size: 8, color: T.MUTED, margins: [0, 0, 0, 0] }); }
function rowsFill(y0, y1, n, gap = 4) { const h = (y1 - y0 - gap * (n - 1)) / n; return Array.from({ length: n }, (_, i) => [y0 + i * (h + gap), h]); }
function colsFill(x0, w, n, gap = 12) { const cw = (w - gap * (n - 1)) / n; return Array.from({ length: n }, (_, i) => [x0 + i * (cw + gap), cw]); }
const pad2 = (n) => String(n).padStart(2, "0");

module.exports = { Slide, T, FONT, BX, BY, BW, BB, BH, SRC_Y, T_H1, T_H2, T_BODY, T_SMALL, T_NOTE, T_KPI, bullets, band, panelHead, rowhead, vchevron, harvey, dot, source, rowsFill, colsFill, pad2, emu, pyRound };
