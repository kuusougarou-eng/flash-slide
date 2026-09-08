"use strict";
// Static compiler: layout id + semantic data -> one-slide PPTX on the golden master.
// The layouts are a port of scripts/make-golden.py; test/golden-author.test.js
// proves each layout reproduces the reference slide XML from the golden data.
const AdmZip = require("adm-zip");
const P = require("../package");
const { bindChart } = require("../charts");
const { Slide } = require("./xml");
const { LAYOUTS } = require("./layouts");
const { fitter } = require("./fit");

let compiler;
const getCompiler = () => compiler || (compiler = new P.GoldenCompiler());

function render(id, data, opts = {}) {
  const layout = LAYOUTS[id];
  if (!layout) throw Error("Unknown golden layout " + id);
  const s = new Slide({ fit: opts.fit ? fitter(opts.fit) : null, page: opts.page || "" });
  const head = layout.build(s, data) || {};
  const xml = s.render({ title: data.title, lead: data.lead, titleSize: head.titleSize, leadSize: head.leadSize});
  return { xml, charts: s.charts, warnings: s.warnings };
}

function build(id, data, opts = {}) {
  const g = getCompiler(), slide = g.get(id), rendered = render(id, data, opts);
  const zip = new AdmZip(g.reference(slide.id));
  const chartRels = P.relationships(zip, slide.part).filter((r) => r.getAttribute("Type").endsWith("/chart"));
  if (chartRels.length !== rendered.charts.length) throw Error(`Layout ${id} expects ${chartRels.length} native chart(s), got ${rendered.charts.length}`);
  let xml = rendered.xml;
  chartRels.forEach((rel, i) => {
    xml = xml.replace(`__CHART_${i}__`, rel.getAttribute("Id"));
    const part = P.target(slide.part, rel.getAttribute("Target"));
    const compiled = bindChart(P.read(zip, part), rendered.charts[i]);
    const workbookRel = P.relationships(zip, part).find((r) => r.getAttribute("Type").endsWith("/package"));
    if (!workbookRel) throw Error("Native chart has no editable workbook");
    zip.updateFile(part, Buffer.from(compiled.xml));
    zip.updateFile(P.target(part, workbookRel.getAttribute("Target")), compiled.workbook);
  });
  zip.updateFile(slide.part, Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
` + xml));
  for (const rel of P.relationships(zip, slide.part).filter((r) => r.getAttribute("Type").endsWith("/notesSlide"))) {
    const part = P.target(slide.part, rel.getAttribute("Target")), note = P.parse(P.read(zip, part));
    for (const t of P.els(note, "a", "t")) { while (t.firstChild) t.removeChild(t.firstChild); }
    zip.updateFile(part, Buffer.from(P.xml(note)));
  }
  return { buffer: zip.toBuffer(), xml, warnings: rendered.warnings, layout: slide.id };
}

module.exports = { render, build, LAYOUTS, getCompiler };
