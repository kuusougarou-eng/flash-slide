"use strict";
/**
 * レイアウトエンジンの単体テスト(Office.js 不要)。
 *   node test/layout.test.js
 * 各サンプル spec × スライドサイズ × プロファイル(既定/参照テンプレート)で、はみ出し・最小フォント・太字レンジを検証する。
 */
const assert = require("assert");
const SlideLayout = require("../public/layout.js");
const { SAMPLES } = require("../server/mock");
const { extractJson, shapeResponse } = require("../server/prompt");

const SIZES = [
  [960, 540],
  [720, 540],
  [1280, 720],
];
// 参照スライド解析結果の模擬(deck-B 相当)
const PROFILE_B = {
  title: { x: 48, y: 28, w: 864, h: 44, fontSize: 24, bold: true, color: "#1F3A5F", align: "left", reuse: true },
  lead: { x: 48, y: 84, w: 864, h: 40, fontSize: 16, color: "#595959", reuse: true },
  body: { x: 48, y: 136, w: 864, h: 326 },
  footnote: { x: 48, y: 470, w: 600, h: 20, fontSize: 10, reuse: true },
  fontName: "Meiryo UI",
  accent: "#1F3A5F",
};
let checks = 0;

function checkPrims(name, lay, W, H) {
  assert.ok(lay.prims.length > 2, `${name}: has prims`);
  const flat = [];
  for (const p of lay.prims) {
    if (p.kind === "table") {
      assert.ok(p.x >= -0.01 && p.y >= -0.01 && p.x + p.w <= W + 0.01 && p.y + p.h <= H + 0.01, `${name}: table within slide`);
      assert.strictEqual(p.cells.length, p.rowHeights.length, `${name}: table rows`);
      assert.ok(p.cells.every((r) => r.length === p.colWidths.length), `${name}: table cols`);
      assert.ok(p.fontSize >= 14, `${name}: table font >= 14`);
      Math.abs(p.rowHeights.reduce((a, b) => a + b, 0) - p.h) < 0.5 || assert.fail(`${name}: table row heights sum`);
      // slide-rules: ヘッダー塗りなし・太字・本文より大きい、最終行の下に罫線なし
      assert.ok(p.cells[0].every((c) => c.bold && (!c.fill || c.fill !== "#404040")), `${name}: header bold, no dark fill`);
      assert.ok(p.headFontSize > p.fontSize || p.fontSize === SlideLayout.FONT.head, `${name}: header font > body`);
      assert.ok(p.cells[p.cells.length - 1].every((c) => !c.borderBottom), `${name}: no rule under last row`);
      for (const f of p.fallback) flat.push(Object.assign({}, f, { role: "tablecell" }));
      checks++;
    } else flat.push(p);
  }
  for (const p0 of flat) {
    if (p0.kind !== "rect") continue;
    // 90°/270° 回転の図形は、見た目の枠(幅と高さを入れ替えて中心を固定)で判定する
    const rot = ((p0.rotation || 0) % 180 + 180) % 180;
    const p = rot === 90 ? Object.assign({}, p0, { x: p0.x + p0.w / 2 - p0.h / 2, y: p0.y + p0.h / 2 - p0.w / 2, w: p0.h, h: p0.w }) : p0;
    assert.ok(p.x >= -0.01 && p.y >= -0.01, `${name} ${W}x${H}: origin >=0 (${p.x},${p.y}) [${p.text.slice(0, 12)}]`);
    assert.ok(p.x + p.w <= W + 0.01, `${name} ${W}x${H}: right ${(p.x + p.w).toFixed(1)} <= ${W} [${p.text.slice(0, 12)}]`);
    assert.ok(p.y + p.h <= H + 0.01, `${name} ${W}x${H}: bottom ${(p.y + p.h).toFixed(1)} <= ${H} [${p.text.slice(0, 12)}]`);
    assert.ok(p.w > 0 && p.h > 0, `${name}: positive size`);
    if (p.kind === "image") {
      assert.ok(p.x >= -0.01 && p.y >= -0.01 && p.x + p.w <= W + 0.01 && p.y + p.h <= H + 0.01, `${name}: image within slide`);
      checks++;
      continue;
    }
    if (p.text) {
      const small = p.fontSize < SlideLayout.FONT.min;
      const allowed =
        (["footnote", "note", "title", "lead", "tablecell", "barlabel", "caption", "kicker", "num"].includes(p.role) && p.fontSize >= (["footnote", "note", "caption", "kicker"].includes(p.role) ? 9 : 11)) ||
        (p.role === "matrixcell" && p.fontSize >= SlideLayout.FONT.tableMin) ||
        (p.role === "orglabel" && p.fontSize >= SlideLayout.FONT.tableMin) || // 体制図の箱: ラベルを切らず箱幅に合わせて縮める
        (p.shrunk && p.fontSize >= SlideLayout.FONT.floor) || // 溢れ時の段階縮小(警告付き)
        (lay.fonts && lay.fonts.grown < 0 && p.fontSize >= lay.fonts.body - 0.01 && lay.fonts.body >= SlideLayout.FONT.floor); // 密なときのスライド全体の一括縮小(全部品で揃う)
      assert.ok(!small || allowed, `${name} ${W}x${H}: font ${p.fontSize} < ${SlideLayout.FONT.min} for "${p.text.slice(0, 20)}"`);
      for (const [s, l] of p.boldRanges) assert.ok(s >= 0 && l > 0 && s + l <= p.text.length, `${name}: bold range in bounds`);
      assert.ok(!p.text.includes("**"), `${name}: bold markers stripped`);
    }
    checks++;
  }
}

// 格子(table)はセルの組み合わせで描く(ネイティブ表プリミティブは出さない)。列見出しの下に太い強調色罫が 1 本
for (const key of ["insight", "flow"]) {
  const lay = SlideLayout.layout(SlideLayout.normalizeSpec(SAMPLES[key]), { width: 960, height: 540 });
  assert.strictEqual(lay.prims.filter((p) => p.kind === "table").length, 0, `${key}: matrix is drawn as cells, not a native table`);
  const thick = lay.prims.filter((p) => p.kind === "line" && p.weight === SlideLayout.STYLE.rule.thick && p.color === SlideLayout.DEFAULT_PALETTE.accent && p.y1 === p.y2);
  assert.ok(thick.length >= 1, `${key}: title rule (${thick.length})`);
  // 列見出しの細い罫は「パネルの外にある格子」だけ(パネル内では余白で区切る)
  const inPanel = JSON.stringify(SlideLayout.normalizeSpec(SAMPLES[key]).body).indexOf('"head"') < JSON.stringify(SlideLayout.normalizeSpec(SAMPLES[key]).body).indexOf('"type":"table"');
  const thinRules = lay.prims.filter((p) => p.kind === "line" && p.weight === 1.2 && p.color === SlideLayout.DEFAULT_PALETTE.line).length;
  const thickBody = lay.prims.filter((p) => p.kind === "line" && p.weight === SlideLayout.STYLE.rule.thick && p.body);
  if (inPanel) assert.strictEqual(thinRules, 0, `${key}: no column rule inside a panel`);
  else {
    // パネル先頭の格子は、列見出し行が帯(濃い塗り)か太い罫でパネル見出しの役を果たす
    const headBand = lay.prims.some((p) => p.kind === "rect" && p.body && !p.text && p.fill === SlideLayout.DEFAULT_PALETTE.fillDark && p.w > 200);
    assert.ok(thickBody.length >= 2 || headBand, `${key}: table header acts as the panel heading (thick rules=${thickBody.length}, band=${headBand})`);
    assert.strictEqual(new Set(thickBody.map((r) => Math.round(r.y1))).size, 1, `${key}: panel rules share one y`);
  }
  checks++;
}
// ntable(Figure)はネイティブ表プリミティブを 1 つ出す
{
  const nt = SlideLayout.normalizeSpec({ title: "t", body: { type: "ntable", corner: "項目", colHeaders: ["A", "B"], rows: [{ head: "r1", cells: ["1", "2"] }, { head: "r2", cells: ["3", "—"] }] } });
  const lay = SlideLayout.layout(nt, { width: 960, height: 540 });
  assert.strictEqual(lay.prims.filter((p) => p.kind === "table").length, 1, "ntable: exactly one native table prim");
  checks++;
}

for (const [name, raw] of Object.entries(SAMPLES)) {
  if (raw.sections) continue;
  const spec = SlideLayout.normalizeSpec(raw);
  assert.ok(spec.body && (spec.body.rows || spec.body.cols || spec.body.type), `${name}: body normalized`);
  for (const [W, H] of SIZES) {
    const lay = SlideLayout.layout(spec, { width: W, height: H });
    checkPrims(name, lay, W, H);
    // 品質ゲート: サンプルは 16:9 で縮小なし(18pt 以上)で収まること
    if (W === 960) assert.ok(!lay.prims.some((p) => p.shrunk), `${name}: no shrink at 960x540 (${lay.warnings.join("; ")})`);
    assert.deepStrictEqual(lay.warnings.filter((w) => !/overflow|shrunk/.test(w)), [], `${name} ${W}x${H}: warnings ${JSON.stringify(lay.warnings)}`);
  }
  // 参照テンプレート準拠
  const layP = SlideLayout.layout(spec, { width: 960, height: 540, profile: PROFILE_B });
  checkPrims(name + "+profile", layP, 960, 540);
  assert.strictEqual(layP.header.titleText, spec.title, "profile: title text passed to renderer");
  assert.ok(!layP.prims.some((p) => p.role === "title"), "profile reuse: no title prim drawn");
  for (const p0 of layP.prims) {
    if (!(p0.body && p0.kind === "rect")) continue;
    const rot = ((p0.rotation || 0) % 180 + 180) % 180;
    const p = rot === 90 ? { y: p0.y + p0.h / 2 - p0.w / 2, h: p0.w } : p0;
    assert.ok(p.y >= PROFILE_B.body.y - 0.01 && p.y + p.h <= PROFILE_B.body.y + PROFILE_B.body.h + 0.01, `${name}: body within profile body`);
  }
  for (const p of layP.prims) if (p.body && p.kind === "rect" && p.text && !p.icon) assert.strictEqual(p.fontName, "Meiryo UI", `${name}: template font applied`);
  // テンプレの本文サイズ(14pt)を継承する: 本文の文字は 14pt 以下、既定(18pt)には戻らない
  const layF = SlideLayout.layout(spec, { width: 960, height: 540, profile: Object.assign({}, PROFILE_B, { bodyFontSize: 14, headFontSize: 14 }) });
  // 版面充填で全体を一括拡大することはある(+6pt まで)が、部品ごとにバラバラに大きくはならない
  assert.ok(layF.fonts.body >= 14 && layF.fonts.body <= 20, `${name}: body font follows the template (${layF.fonts.body}pt)`);
  for (const p of layF.prims) if (p.body && p.kind === "rect" && p.text && !["footnote", "kpi", "num"].includes(p.role)) assert.ok(p.fontSize <= layF.fonts.body + 6.01, `${name}: no body text larger than the slide body + head step (${p.fontSize}pt "${p.text.slice(0, 12)}")`);
  assert.ok(!layF.prims.some((p) => p.shrunk), `${name}: template sizing leaves no shrink (${layF.warnings.join("; ")})`);
}

// 旧形式(boxes/columns/matrix/process/bullets)の変換
{
  const legacy = {
    title: "旧形式",
    lead: "",
    body: { type: "boxes", items: [{ head: "A", text: "a" }, { head: "B", text: "b" }, { head: "C", text: "c" }] },
  };
  const s = SlideLayout.normalizeSpec(legacy);
  assert.ok(s.body.cols && s.body.cols.length === 3 && s.body.cols[0].type === "cell", "boxes → cols of cells");
  const m = SlideLayout.normalizeSpec({ body: { type: "matrix", colHeaders: ["x"], rows: [{ head: "r", cells: ["1"] }] } });
  assert.strictEqual(m.body.type, "table");
  const p = SlideLayout.normalizeSpec({ body: { type: "process", steps: [{ head: "s1" }, { head: "s2" }] } });
  assert.ok(p.body.cols && p.body.cols[0].shape === "home" && p.body.cols[1].shape === "home", "process → band of home-plate cells");
  const p2 = SlideLayout.normalizeSpec({ body: { type: "chevrons", steps: [{ head: "s1", bullets: ["a"] }, { head: "s2", bullets: ["b", "c"] }] } });
  assert.ok(p2.body.rows && p2.body.rows.length === 2 && p2.body.rows[0].cols[0].shape === "home" && p2.body.rows[1].cols[1].items.length === 2, "chevrons with bodies → rows[band, bodies]");
  const p6 = SlideLayout.normalizeSpec({ body: { type: "chevrons", steps: Array.from({ length: 6 }, (_, i) => ({ head: "s" + i, text: "t" + i })) } });
  assert.ok(p6.body.type === "table" && p6.body.headShape === "chevron" && p6.body.rows.length === 6, "6+ steps → vertical matrix with chevron row heads");
  const c = SlideLayout.normalizeSpec({ body: { type: "columns", columns: [{ head: "h", bullets: ["a", "b"] }] } });
  assert.strictEqual(c.body.cols[0].type, "cell");
  assert.deepStrictEqual(c.body.cols[0].items, ["a", "b"]);
  const b = SlideLayout.normalizeSpec({ body: { type: "BULLETS", bullets: ["**A**: a", { title: "C", description: "c" }] } });
  assert.strictEqual(b.body.type, "cell");
  assert.deepStrictEqual(b.body.items, ["**A**: a", "**C**:c"]);
  const typedCols = SlideLayout.normalizeSpec({ body: { rows: [{ type: "cols", weights: [1, 1], cols: [{ type: "cell", head: "A", text: "a" }, { type: "cell", head: "B", text: "b" }] }, { type: "cell", items: ["x", "y"] }] } });
  assert.ok(typedCols.body.rows.length === 2 && typedCols.body.rows[0].cols && typedCols.body.rows[0].cols.length === 2, '{"type":"cols"} container is kept');
  const figLine = SlideLayout.normalizeSpec({ body: { type: "figure", line: { labels: ["a", "b"], series: [{ label: "x", values: [1, 2] }] } } });
  assert.strictEqual(figLine.body.type, "line", '{"type":"figure","line":{…}} unwrapped');
  const wrapped = SlideLayout.normalizeSpec({ body: { table: { headShape: "chevron", colHeaders: ["a"], rows: [{ head: "r", cells: ["1"] }] } } });
  assert.ok(wrapped.body.type === "table" && wrapped.body.headShape === "chevron", "{table:{…}} wrapper unwrapped");
  const wrapped2 = SlideLayout.normalizeSpec({ body: { rows: [{ cell: { head: "h", text: "t" } }, { bars: { items: [{ label: "a", value: 1 }] } }] } });
  assert.ok(wrapped2.body.rows[0].type === "cell" && wrapped2.body.rows[1].type === "bars", "nested wrappers unwrapped");
  const junk = SlideLayout.normalizeSpec({ body: { cols: [{ type: "cell", text: "PM 1 名 leadAlignedToTitleLikeNoDuplicateCheck?ignored:true} }]}]}]} }]}]} }} }]" }] } });
  assert.strictEqual(junk.body.cols[0].text, "PM 1 名", "garbage tail is stripped from LLM text");
  // 格子は縮小時も行数分の最低高さを保つ
  const tight = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "t", body: { rows: [{ type: "table", headShape: "chevron", colHeaders: ["a", "b"], rows: Array.from({ length: 5 }, (_, i) => ({ head: "s" + i, cells: ["x", "y"] })) }, { type: "cell", text: "長い本文".repeat(120) }] } }), { width: 960, height: 540 });
  const plates = tight.prims.filter((p) => p.rotation === 90 && p.shape === "homePlate");
  assert.ok(plates.length === 5 && plates.every((p) => p.w >= 24), "table rows keep a minimum height under pressure (" + plates.map((p) => Math.round(p.w)).join(",") + ")");
  const lb = SlideLayout.normalizeSpec({ body: { type: "label", text: "見出し" } });
  assert.ok(lb.body.type === "cell" && lb.body.head === "見出し", "label → cell{head}");
  const co = SlideLayout.normalizeSpec({ body: { type: "callout", text: "示唆" } });
  assert.ok(co.body.type === "cell" && co.body.fill === "tint", "callout → cell{fill:tint}");
  checks++;
}

// グリッド: weights と arrow の既定幅、icon 正規化、深さ制限
{
  const s = SlideLayout.normalizeSpec({
    title: "現状と課題：価格は 8 倍開いています",
    body: { cols: [{ type: "card", head: "a", text: "x", icon: "Warning" }, { type: "arrow" }, { type: "card", head: "b", text: "y" }], weights: [5, 0.6, 5] },
  });
  assert.strictEqual(s.title, "価格は 8 倍開いている", "title: label prefix removed and です/ます converted to 常体");
  assert.strictEqual(SlideLayout.cleanTitle("B 案が最適です。"), "B 案が最適");
  assert.strictEqual(SlideLayout.cleanTitle("6 割解消できます"), "6 割解消できる");
  assert.deepStrictEqual(s.body.weights, [1, 0.14, 1], "2 panels (+arrow) are forced to equal widths");
  assert.strictEqual(s.body.cols[0].icon, "warning");
  const lay = SlideLayout.layout(s, { width: 960, height: 540 });
  // アイコンは画像プリミティブ(サーバで Tabler Icons を描画)。因果の矢印は塗り三角
  const img = lay.prims.find((p) => p.kind === "image");
  assert.ok(img && img.name === "warning" && img.w > 0, "card icon emitted as image prim");
  assert.ok(img.x >= 0 && img.y >= 0 && img.x + img.w <= 960 && img.y + img.h <= 540, "image within slide");
  const tri = lay.prims.find((p) => p.shape === "triangle" && p.rotation === 90);
  assert.ok(tri, "arrow rendered as rotated filled triangle");
  assert.ok(!lay.prims.some((p) => p.fontName === "Segoe MDL2 Assets"), "no icon-font glyphs (avoids tofu)");
  // 塗りのある矩形に枠線が無い / 角丸を使わない(slide-rules)
  for (const p of lay.prims) if (p.kind === "rect" && p.fill && p.text === "") assert.ok(!p.line, "filled box has no border");
  assert.ok(!lay.prims.some((p) => p.shape === "roundRect"), "no rounded rectangles");
  checkPrims("grid", lay, 960, 540);
  // 深い入れ子は打ち切られる
  const deep = SlideLayout.normalizeSpec({ body: { rows: [{ rows: [{ rows: [{ rows: [{ rows: [{ type: "text", text: "deep" }] }] }] }] }] } });
  assert.ok(deep.body, "deep nesting handled");
  checks++;
}

// parseRuns / extractJson / shapeResponse
{
  const r = SlideLayout.parseRuns("**価格改定**:主力を**3%**値上げ\n次の行");
  assert.strictEqual(r.text, "価格改定:主力を3%値上げ\n次の行");
  assert.deepStrictEqual(r.boldRanges, [[0, 4], [8, 2]]);
  assert.deepStrictEqual(extractJson('```json\n{"a":1}\n```'), { a: 1 });
  assert.deepStrictEqual(extractJson('以下です:\n{"a":[1,2,],}'), { a: [1, 2] });
  assert.deepStrictEqual(extractJson('{\n  "type": "table",\n  - "colHeaders": ["a", "b"],\n  "rows": []\n}'), { type: "table", colHeaders: ["a", "b"], rows: [] }, "stray list dash before a key is repaired");
  const one = shapeResponse({ title: "t", body: { type: "text", text: "x" } }, SlideLayout.normalizeSpec, 2);
  assert.strictEqual(one.slides.length, 1);
  const two = shapeResponse({ slides: [{ title: "a" }, { title: "b" }, { title: "c" }] }, SlideLayout.normalizeSpec, 2);
  assert.strictEqual(two.slides.length, 2, "capped to maxSlides");
  const sec = shapeResponse(SAMPLES.sections, SlideLayout.normalizeSpec, 2);
  assert.strictEqual(sec.sections.length, 4);
  checks++;
}

// 長文でも枠外に出ない(縮小は PowerPoint 側の autofit に委ねる)
{
  const long = {
    title: "非常に長いタイトルで二行にまたがる可能性があるケースの検証をするための見出しテキストです",
    lead: "リード文も長め。".repeat(10),
    body: { rows: [{ type: "bullets", items: Array.from({ length: 7 }, (_, i) => "**要点" + i + "**:" + "補足説明テキスト".repeat(5)) }, { cols: [{ type: "card", head: "A", text: "本文".repeat(60) }, { type: "card", head: "B", text: "本文".repeat(60) }] }] },
  };
  const lay = SlideLayout.layout(SlideLayout.normalizeSpec(long), { width: 960, height: 540 });
  checkPrims("long", lay, 960, 540);
  checks++;
}

// v4: 強調バジェット(highlight 2 / tint 1)、モノトーン、ネスト箇条書き、兄弟列の行揃え、自動リフロー、列グループ、塗りセル
{
  const many = SlideLayout.normalizeSpec({
    title: "t",
    body: { cols: [
      { type: "cell", head: "a", text: "x", highlight: true },
      { type: "cell", head: "b", text: "y", highlight: true },
      { type: "cell", head: "c", text: "z", highlight: true },
      { type: "callout", text: "s1" }, { type: "callout", text: "s2" },
    ] },
  });
  const cells = [];
  (function walk(n) { if (n.rows || n.cols) return (n.rows || n.cols).forEach(walk); cells.push(n); })(many.body);
  // 強調は「結論(title/lead)に名前が出る対象」だけに許され、その上で 1 枚 2 箇所まで
  assert.ok(cells.filter((c) => c.highlight).length <= 2, "highlight budget <= 2");
  assert.strictEqual(cells.filter((c) => c.fill === "tint").length, 1, "tint budget = 1");
  assert.ok(many.body.rows && many.body.rows.length === 2, "5 leaves in cols → reflowed into 2 rows");
  // 格子の強調もバジェットに含める(highlightCol + 行 + セル)
  // 強調は title/lead に名前が出る対象(A社・r1)だけに許される
  const tb = SlideLayout.normalizeSpec({ title: "A社が最適で、r1 を優先する", body: { type: "table", colHeaders: ["A社", "B社"], highlightCol: 0, rows: [{ head: "r1", cells: ["1", "2"], highlight: true }, { head: "s1", cells: [{ text: "3", highlight: true }, "4"] }] } });
  assert.strictEqual(tb.body.highlightCol, 0);
  assert.strictEqual(tb.body.rows[0].highlight, true);
  assert.ok(!tb.body.rows[1].cells[0].highlight, "third highlight dropped");
  // モノトーン: accent:"none" なら強調色が prims に出ない
  const lay = SlideLayout.layout(SlideLayout.normalizeSpec(SAMPLES.flow), { width: 960, height: 540, palette: { accent: "none" } });
  assert.ok(!lay.prims.some((p) => (p.fill === "#FD5108") || (p.color === "#FD5108")), "mono mode: no accent color");
  checkPrims("mono", lay, 960, 540);
  // ネスト箇条書き: 点は PowerPoint 本来の箇条書き、第 2 階層はソフト改行で親の段落に入れる(子には点が付かない)
  const h = SlideLayout.layout(SlideLayout.normalizeSpec(SAMPLES.hierarchy), { width: 960, height: 540 });
  const nested = h.prims.find((p) => p.kind === "rect" && /\v/.test(p.text || ""));
  assert.ok(nested && nested.bullets, "nested items use native bullets and put children on soft line breaks");
  assert.ok(!/•/.test(nested.text), "no pseudo top-level bullet character in the text");
  // 子の記号「–」は、子が折り返すときだけ付ける(幅と文字サイズが決まってから決める)
  const short = SlideLayout.layout(SlideLayout.normalizeSpec({ panelCount: 1, title: "T", body: { type: "cell", items: ["**親 A**", ["子 1", "子 2"]] } }), { width: 960, height: 540 });
  const shortPrim = short.prims.find((p) => /\v/.test(p.text || ""));
  assert.ok(shortPrim && !/–/.test(shortPrim.text), "children that fit on one line need no marker");
  const longChild = "主要部材の単価が前年比 +14% で、仕入先 182 社に分散した結果、上位 10 社で購買額の 38% にとどまっている";
  const wrapCase = SlideLayout.layout(SlideLayout.normalizeSpec({ panelCount: 1, title: "T", body: { type: "cell", items: ["**親 A**", [longChild, longChild]] } }), { width: 960, height: 540 });
  const widePrim = wrapCase.prims.find((p) => /\v/.test(p.text || ""));
  assert.ok(widePrim && /\v　– /.test(widePrim.text), "children that wrap get a marker so the boundary stays visible");
  // 濃い見出し帯は「結論に名前が出ている軸」1 つだけ。行を濃くしたら列は薄くする
  const darkRow = SlideLayout.normalizeSpec({ title: "可視化基盤が最も参入余地が大きい", body: { type: "table", rowHeadFill: "dark", colHeaders: ["市場規模", "成長率"], rows: [{ head: "倉庫自動化", cells: ["3,200億円", "8.2%"] }, { head: "可視化基盤", cells: ["1,800億円", "17.1%"] }] } });
  // 濃くするのは結論に名前が出ている行だけ(全行を濃くするとどこも立たない)。列の帯は薄くなる
  assert.deepStrictEqual(darkRow.body.rows.map((r) => !!r.headDark), [false, true], "only the row named in the message is dark");
  assert.strictEqual(darkRow.body.headFill, "light", "the other axis falls back to a light band");
  // 単位はセルに散らさず列見出しへ
  const units = SlideLayout.normalizeSpec({ panelCount: 1, compositionVersion: 1, title: "t", body: { type: "table", colHeaders: ["市場規模", "成長率"], rows: [{ head: "A", cells: ["3,200億円", "8.2%"] }, { head: "B", cells: ["1,800億円", "17.1%"] }, { head: "C", cells: ["1,300億円", "6.0%"] }] } });
  assert.deepStrictEqual(units.body.colHeaders, ["市場規模(億円)", "成長率(%)"], "units move to the column header");
  assert.deepStrictEqual(units.body.rows[0].cells, ["3,200", "8.2"], "cells keep the number only");
  const darkNone = SlideLayout.normalizeSpec({ title: "3 つの観点で比較する", body: { type: "table", rowHeadFill: "dark", colHeaders: ["A", "B"], rows: [{ head: "甲", cells: ["1", "2"] }, { head: "乙", cells: ["3", "4"] }] } });
  assert.ok(darkNone.body.rows.every((r) => !r.headDark), "an axis not named in the message is not darkened");

  // 割った 2 枚目は番号が続く(01 に戻らない)。タスクペインが再正規化しても numberFrom が残る
  const sixSteps = SlideLayout.normalizeSpec({ panelCount: 1, compositionVersion: 1, title: "6 工程で移行する", body: { type: "table", headShape: "chevron", numbered: true, colHeaders: [], rows: ["現状把握", "設計", "参照系移行", "更新系移行", "並行稼働", "旧環境停止"].map((h, i) => ({ head: h, cells: ["工程 " + (i + 1) + " の作業内容。判定条件と成果物を併記する。"] })) } });
  const halves = SlideLayout.splitSpec(sixSteps);
  assert.strictEqual(halves[1].body.numberFrom, 4, "second half numbers from 4");
  assert.strictEqual(SlideLayout.normalizeSpec(JSON.parse(JSON.stringify(halves[1]))).body.numberFrom, 4, "numberFrom survives re-normalization");
  const secondHalf = SlideLayout.layout(SlideLayout.normalizeSpec(JSON.parse(JSON.stringify(halves[1]))), { width: 960, height: 540 });
  assert.ok(secondHalf.prims.some((p) => /^04/.test(p.text || "")), "rendered numbering continues at 04");
  // 箇条書きの塊を持つ 2 列の格子は、再正規化しても列見出しが本文へ戻らない(「論点: ・…」にならない)
  const twoCol = SlideLayout.normalizeSpec({ panelCount: 1, compositionVersion: 1, title: "在庫を立て直す", body: { type: "table", colHeaders: ["論点", "現状"], rows: [{ head: "在庫", cells: ["欠品 年4,200万\n・廃棄 年1,800万", "S&OP を週次で実施"] }, { head: "営業", cells: ["受注 +18%\n・新規 32社", "値引き決裁を基準化"] }] } });
  const twoColAgain = SlideLayout.normalizeSpec(JSON.parse(JSON.stringify(twoCol)));
  assert.deepStrictEqual(twoColAgain.body.colHeaders, ["論点", "現状"], "column headers survive re-normalization");
  assert.ok(!/^論点[:：]/.test(String(twoColAgain.body.rows[0].cells[0].text || twoColAgain.body.rows[0].cells[0])), "label is not folded back into the cell");
  // 図形で組む格子は版面を埋めるのが正しい姿。fill だけを理由に「収まらない」と判定しない
  const grid = SlideLayout.normalizeSpec({ panelCount: 1, compositionVersion: 1, title: "t", body: { type: "table", colHeaders: ["A", "B"], rows: [{ head: "甲", cells: ["1", "2"] }, { head: "乙", cells: ["3", "4"] }, { head: "丙", cells: ["5", "6"] }] } });
  assert.ok(SlideLayout.fitsReadably(grid).ok, "a small shape grid is not rejected for filling the body");

  // 矢羽は工程の進行にだけ。状態の列挙(課題 / 対応 / 予定)には使わない
  const status = SlideLayout.normalizeSpec({ panelCount: 1, compositionVersion: 1, title: "棚卸しは順調", body: { type: "table", headShape: "chevron", numbered: true, colHeaders: [], rows: [{ head: "棚卸し", cells: ["860 件"] }, { head: "課題", cells: ["仕様書 18 本欠落"] }, { head: "対応", cells: ["2 名追加"] }, { head: "来週の予定", cells: ["全件完了"] }] } });
  assert.notStrictEqual(status.body.headShape, "chevron", "status rows are not drawn as chevrons");
  const process = SlideLayout.normalizeSpec({ panelCount: 1, compositionVersion: 1, title: "4 段で移行する", body: { type: "table", headShape: "chevron", numbered: true, colHeaders: [], rows: [{ head: "現状把握", cells: ["棚卸し"] }, { head: "設計", cells: ["方式"] }, { head: "移行", cells: ["切替"] }, { head: "停止", cells: ["解約"] }] } });
  assert.strictEqual(process.body.headShape, "chevron", "process rows keep chevrons");

  // 列方向の矢羽: 行見出しが無く headShape:"chevron" なら、列見出しが左→右の矢羽になる
  const colChev = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "移行は 4 段階で進める", body: { type: "table", headShape: "chevron", colHeaders: ["現状把握", "参照系移行", "更新系移行", "旧環境停止"], rows: [{ cells: ["棚卸し", "BI 38 本", "バッチ 64 本", "解約"] }, { cells: ["可視化", "並行稼働", "リハーサル", "手順更新"] }] } }), { width: 960, height: 540 });
  const chevShapes = colChev.prims.filter((p) => p.shape === "homePlate");
  assert.strictEqual(chevShapes.length, 4, "column headers become chevrons (" + chevShapes.length + ")");
  assert.ok(chevShapes.every((p, i, a) => i === 0 || p.x > a[i - 1].x), "chevrons run left to right");
  assert.ok(chevShapes.every((p) => Math.abs(p.y - chevShapes[0].y) < 0.5), "chevrons share one row");

  // 兄弟列の行揃え: 4 列の見出し(1 段目)の高さ・下端が揃う
  const e = SlideLayout.layout(SlideLayout.normalizeSpec(SAMPLES.enumerate), { width: 960, height: 540 });
  const rules = e.prims.filter((p) => p.kind === "line" && p.weight === SlideLayout.STYLE.rule.thick && p.body);
  assert.strictEqual(new Set(rules.map((r) => Math.round(r.y1))).size, 1, "heading rules of sibling columns share one y");
  assert.strictEqual(rules.length, 4, "every panel heading keeps its rule");
  // 見出しテキストは罫線と同じ幅・同じ x(線分の上に中央揃えで載る)
  for (const rl of rules) {
    const tb = e.prims.find((p) => p.kind === "rect" && p.bold && Math.abs(p.x - rl.x1) < 0.01 && Math.abs(p.x + p.w - rl.x2) < 0.01 && p.y < rl.y1 && p.y + p.h >= rl.y1 - 3);
    assert.ok(tb && tb.align === "center", "heading text spans the rule width, centered");
  }
  // 黒下線はパネルの最上位だけ(塊の中で入れ子にしない)
  const panelNest = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "t", body: { rows: [
    { type: "cell", head: "推進体制" },
    { cols: [{ type: "cell", head: "業務チーム", items: ["要件定義", "受入"] }, { type: "cell", head: "IT チーム", items: ["設計", "開発"] }] },
  ] } }), { width: 960, height: 540 });
  const thickRules = panelNest.prims.filter((p) => p.kind === "line" && p.weight === SlideLayout.STYLE.rule.thick && p.body);
  assert.strictEqual(thickRules.length, 1, "only the panel heading is ruled (" + thickRules.length + ")");
  // パネルの中の格子は列見出しの罫を引かない
  const panelTable = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "t", body: { rows: [
    { type: "cell", head: "評価" },
    { type: "table", corner: "軸", colHeaders: ["A", "B"], rows: [{ head: "コスト", cells: ["1", "2"] }] },
  ] } }), { width: 960, height: 540 });
  assert.strictEqual(panelTable.prims.filter((p) => p.kind === "line" && p.weight === 1.2 && p.body).length, 0, "no column rule inside a panel");
  // 単体の格子(パネル外)は列見出しの罫を引く
  const bareTable = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "t", body: { type: "table", corner: "軸", colHeaders: ["A", "B"], rows: [{ head: "コスト", cells: ["1", "2"] }] } }), { width: 960, height: 540 });
  const gridRule = bareTable.prims.filter((p) => p.kind === "line" && p.weight === 1.2 && p.body).length;
  const gridBand = bareTable.prims.some((p) => p.kind === "rect" && p.body && !p.text && p.fill === "#404040" && p.w > 300);
  assert.ok(gridRule === 1 || gridBand, "standalone grid keeps its column rule or header band");
  // 表の列幅は内容量に比例する(短い列が広く、説明列が狭くならない)
  const wmat = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "t", body: { type: "table", corner: "ベンダー", colHeaders: ["5 年総額", "稼働まで", "適合度"], rows: [
    { head: "A 社", cells: ["7.2 億円", "18 か月", "中。カスタマイズが多く、標準機能では 6 割程度しか賄えない"] },
    { head: "B 社", cells: ["6.6 億円", "12 か月", "高。標準機能で 8 割を賄え、追加開発は限定的である"] },
  ] } }), { width: 960, height: 540 });
  const hdr = wmat.prims.filter((p) => p.role === "matrixcell" && p.bold && /総額|稼働|適合度/.test(p.text));
  const wide = hdr.find((p) => p.text === "適合度").w, narrow = hdr.find((p) => p.text === "稼働まで").w;
  assert.ok(wide > narrow * 1.3, "description column is wider than the short one (" + Math.round(wide) + " vs " + Math.round(narrow) + ")");
  for (const p of hdr) assert.ok(SlideLayout.estimateLines(p.text, p.w - 14, p.fontSize) <= 1, "column header fits on one line: " + p.text + " (" + Math.round(p.w) + "pt)");
  // 濃い塗りは短い 1 つだけ。長文・2 つ目は薄い塗りになる
  // 濃い塗り(見出し帯の反転)は title/lead に名前が出る対象 1 つだけ。それ以外は塗りなし
  const darks = SlideLayout.normalizeSpec({ title: "案Aを推す", body: { cols: [{ type: "cell", head: "案A", fill: "dark", text: "短い強調" }, { type: "cell", head: "案B", fill: "dark", text: "二つ目" }, { type: "cell", head: "案C", fill: "dark", text: "とても長い説明文が続く。".repeat(6) }] } });
  assert.deepStrictEqual(darks.body.cols.map((c) => c.fill || "none"), ["dark", "none", "none"], "dark fill limited to the mentioned cell");
  // 箱を含む rows では LLM の weights を無視して箱は中身の高さになる
  const boxedRows = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "t", body: { rows: [{ cols: [{ type: "cell", head: "a", shape: "step", text: "x" }, { type: "cell", head: "b", shape: "step", text: "y" }] }, { type: "cell", items: ["p", "q"] }], weights: [4, 1] } }), { width: 960, height: 540 });
  const stepBoxes = boxedRows.prims.filter((p) => p.kind === "rect" && !p.text && p.fill && p.h > 40);
  assert.ok(stepBoxes.every((p) => p.h < 150), "step boxes hug their content (" + stepBoxes.map((p) => Math.round(p.h)).join(",") + ")");
  // 桁違いの副系列は線にせず注記に落ちる
  const twoScale = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "t", body: { type: "line", labels: ["a", "b"], series: [{ label: "件数", values: [1200, 2160] }, { label: "人員", values: [9, 10] }] } }), { width: 960, height: 540 });
  assert.ok(twoScale.prims.some((p) => p.role === "note" && /人員/.test(p.text)), "small series becomes a note");
  // 因果の三角はアイコン的な正方形に収まる
  const tri = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "t", body: { cols: [{ type: "cell", text: "前提" }, { type: "arrow" }, { type: "cell", text: "帰結" }] } }), { width: 960, height: 540 });
  const t3 = tri.prims.find((p) => p.shape === "triangle");
  assert.ok(t3 && Math.abs(t3.w - t3.h) < 0.01 && t3.w <= 22, "arrow triangle is a small square (" + Math.round(t3.w) + "x" + Math.round(t3.h) + ")");
  // 格子の列見出しは中身の揃えに合わせる(数値列は右)
  const alignMat = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "t", body: { type: "table", corner: "項目", colHeaders: ["構成比", "背景"], rows: [{ head: "A", cells: ["62%", "定型質問に担当者が対応"] }, { head: "B", cells: ["21%", "品質が個人で異なる"] }] } }), { width: 960, height: 540 });
  const heads = alignMat.prims.filter((p) => p.role === "matrixcell" && p.bold && /構成比|背景/.test(p.text));
  assert.strictEqual(heads.find((p) => p.text === "構成比").align, "right", "numeric column header is right aligned");
  assert.strictEqual(heads.find((p) => p.text === "背景").align, "left", "text column header is left aligned");
  // 情報量が多い 2 列は縦積みに切り替わる(横長に読ませる)
  const long = "現場の運用実態を踏まえると、段階的な移行計画と並行稼働の判断基準を先に固める必要がある。".repeat(3);
  const denseSpec = SlideLayout.normalizeSpec({ title: "t", body: { cols: [{ type: "cell", head: "A", text: long }, { type: "cell", head: "B", text: long }, { type: "cell", head: "C", text: long }] } });
  const denseLay = SlideLayout.layout(denseSpec, { width: 960, height: 540 });
  const bodies = denseLay.prims.filter((p) => p.body && p.text && !p.bold && !p.role);
  assert.ok(bodies.every((p) => p.w > 700), "dense text columns are stacked full width (" + bodies.map((p) => Math.round(p.w)).join(",") + ")");
  // JSON の残骸は本文に出さない / 1 値ずつの stacked は横棒になる / 単位もどきは無視
  const debris = SlideLayout.normalizeSpec({ body: { cols: [{ type: "cell", text: 'weights":[4,0.3,5]}}}}]}},' }, { type: "cell", text: "本文" }] } });
  assert.strictEqual(debris.body.cols.length, 1, "json debris dropped");
  const st = SlideLayout.normalizeSpec({ body: { type: "stacked", parts: ["比率"], items: [{ label: "外部", values: [30] }, { label: "内部", values: [50] }, { label: "他", values: [20] }] } });
  assert.strictEqual(st.body.type, "bars", "single-value stacked becomes bars");
  const badUnit = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "t", body: { type: "line", unit: "2019年比", labels: ["a", "b"], series: [{ label: "x", values: [10, 20] }] } }), { width: 960, height: 540 });
  assert.ok(!badUnit.prims.some((p) => /2019年比/.test(p.text || "")), "bogus unit ignored");
  // スキーマの語だけの文字列は捨てる(本文に "type" や "arrow" と印字されない)
  const junkWords = SlideLayout.normalizeSpec({ body: { cols: [{ type: "cell", text: "本文" }, "type", "arrow", { type: "cell", text: "続き" }] } });
  assert.strictEqual(junkWords.body.cols.length, 2, "schema-word fragments dropped");
  // 折れ線: 0 から遠い系列は基線を詰めて縦方向を使い切る
  const ln = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "t", body: { type: "line", unit: "件", labels: ["2023", "2024", "2025", "2026"], series: [{ label: "件数", values: [1200, 1500, 1900, 2160] }] } }), { width: 960, height: 540 });
  const dots = ln.prims.filter((p) => p.shape === "ellipse").map((p) => p.y).sort((a, b) => a - b);
  assert.ok(dots.length === 4 && dots[3] - dots[0] > 150, "line uses the vertical space (" + Math.round(dots[3] - dots[0]) + "pt)");
  // データの無い figure は捨てる(空の軸だけが描かれない)
  const emptyFig = SlideLayout.normalizeSpec({ body: { cols: [{ type: "bars", unit: "", items: [] }, { type: "cell", text: "本文" }] } });
  assert.ok(!JSON.stringify(emptyFig.body).includes('"bars"'), "empty bars dropped");
  const emptyLine = SlideLayout.normalizeSpec({ body: { type: "line", labels: ["a"], series: [{ label: "x", values: [] }] } });
  assert.notStrictEqual(emptyLine.body.type, "line", "empty line dropped");
  const okBars = SlideLayout.normalizeSpec({ body: { type: "bars", items: [{ label: "a", value: 3 }] } });
  assert.strictEqual(okBars.body.type, "bars", "bars with data kept");
  // 同じ容器の本文セルは同じ文字サイズになる
  const mixed = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "t", body: { cols: [
    { rows: [{ type: "cell", head: "起点", items: ["営業員は 120 名 → 96 名と 2 割減った", "担当顧客数は 1 人あたり 1.4 倍に増えた"] }, { type: "cell", items: ["余力が接点の維持ではなく配分調整に吸われた", "訪問と提案の量が落ちる前提ができた"] }] },
    { rows: [{ type: "cell", head: "結果", items: ["新規訪問は月 480 件 → 290 件に減った", "追加受注は 15% 減になった", "コンペ勝率は 42% → 31% に落ちた"] }, { type: "cell", items: ["量だけでなく質と頻度も落ちている"] }] },
  ] } }), { width: 960, height: 540, profile: Object.assign({}, PROFILE_B, { bodyFontSize: 14, headFontSize: 14 }) });
  const bodyFs = new Set(mixed.prims.filter((p) => p.gid && p.text && !p.role).map((p) => p.fontSize));
  assert.ok(bodyFs.size <= 2, "cell bodies share sizes within their container (" + Array.from(bodyFs).join(",") + ")");
  // 観点が 1 つしかない格子は箇条書きに落ちる(格子の乱用を機械的に抑える)
  const thin = SlideLayout.normalizeSpec({ body: { type: "table", corner: "項目", colHeaders: ["内容"], rows: [{ head: "A", cells: ["あ"] }, { head: "B", cells: ["い"] }] } });
  assert.strictEqual(thin.body.type, "cell", "1-column table becomes a list");
  assert.deepStrictEqual(thin.body.items, ["**A**: あ", "**B**: い"]);
  const keep2 = SlideLayout.normalizeSpec({ body: { type: "table", corner: "項目", colHeaders: ["x", "y"], rows: [{ head: "A", cells: ["あ", "い"] }] } });
  assert.strictEqual(keep2.body.type, "table", "2-column table stays a table");
  const keepChev = SlideLayout.normalizeSpec({ body: { type: "table", headShape: "chevron", colHeaders: ["概要"], rows: [{ head: "現状把握", cells: ["…"] }] } });
  assert.strictEqual(keepChev.body.type, "table", "swimlane keeps its grid even with one column");
  // 横に並ぶ格子は同じ文字サイズになる
  const twoMat = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "t", body: { cols: [
    { type: "table", corner: "項目", colHeaders: ["変化", "補足"], rows: [{ head: "路線バス利用者数", cells: ["2019 年比 22% 減", "全体需要が縮小している"] }, { head: "65 歳以上の利用", cells: ["9% 増", "高齢者依存が高まる"] }, { head: "通学利用", cells: ["41% 減", "従来の主要需要が落ち込む"] }] },
    { type: "table", corner: "項目", colHeaders: ["状況"], rows: [{ head: "運転手不足", cells: ["87% が不足と回答"] }, { head: "赤字路線", cells: ["212 中 168 が赤字"] }] },
  ] } }), { width: 960, height: 540, profile: Object.assign({}, PROFILE_B, { bodyFontSize: 14, headFontSize: 14 }) });
  const mats = {};
  twoMat.prims.filter((p) => p.role === "matrixcell").forEach((p) => (mats[p.mid] = Math.max(mats[p.mid] || 0, p.fontSize)));
  assert.strictEqual(new Set(Object.values(mats)).size, 1, "side-by-side matrices share one font size (" + JSON.stringify(mats) + ")");
  // 横棒のラベルは折り返さず、実グリフ幅に余裕のある枠に収まる
  const rb = SlideLayout.layout(SlideLayout.normalizeSpec(SAMPLES.report), { width: 960, height: 540, profile: Object.assign({}, PROFILE_B, { bodyFontSize: 14, headFontSize: 14 }) });
  for (const p of rb.prims.filter((p) => p.role === "barlabel" && p.align === "right")) {
    assert.strictEqual(p.wrap, false, "bar labels do not wrap");
    assert.ok(SlideLayout.estimateLines(p.text, p.w - 16, p.fontSize) <= 1, `bar label fits on one line: ${p.text}`);
  }
  // 段数が違う列(左: 見出し+格子、右: 見出し+箇条書き+示唆)でも見出し罫の y が揃う
  const rp = SlideLayout.layout(SlideLayout.normalizeSpec(SAMPLES.report), { width: 960, height: 540 });
  const rpRules = rp.prims.filter((p) => p.kind === "line" && p.weight === SlideLayout.STYLE.rule.thick && p.body);
  assert.strictEqual(new Set(rpRules.map((r) => Math.round(r.y1))).size, 1, "heading rules aligned across columns with different row counts");
  // 2 パネル(格子 + 箇条書き)でも両方の見出しに罫がある
  const fl = SlideLayout.layout(SlideLayout.normalizeSpec(SAMPLES.flow), { width: 960, height: 540 });
  assert.strictEqual(fl.prims.filter((p) => p.kind === "line" && p.weight === SlideLayout.STYLE.rule.thick && p.body).length, 2, "both panel headings ruled");
  // 矢羽は小さいと長方形に落ちる
  const tiny = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "t", body: { cols: Array.from({ length: 9 }, (_, i) => ({ type: "cell", head: "s" + i, shape: "chevron" })) } }), { width: 960, height: 540 });
  assert.ok(tiny.prims.filter((p) => p.fill && p.text).every((p) => p.shape === "rect"), "narrow bands fall back to rectangles");
  // スイムレーン: 行見出しの矢羽(回転図形)が行数分あり、先頭は五角形
  const sw = SlideLayout.layout(SlideLayout.normalizeSpec(SAMPLES.swimlane), { width: 960, height: 540 });
  const chev = sw.prims.filter((p) => p.rotation === 90 && p.shape === "homePlate");
  assert.strictEqual(chev.length, 4, "4 rotated home plates");
  // 強調は title/lead に名前が出る段にだけ許される(サンプルの title は段名を含まないので全段が既定の濃い塗り)
  assert.ok(chev.every((p) => p.fill === "#404040" || p.fill === "#FD5108"), "stages use the dark fill or the accent only");
  // 列グループ・軸ラベル・塗りセル
  const g = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "t", body: { type: "table", corner: "", colHeaders: ["a", "b", "c", "d"], colGroups: [{ text: "G1", span: 2 }, { text: "G2", span: 2 }], rows: [{ head: "r", cells: ["1", "2", "3", "4"] }] } }), { width: 960, height: 540 });
  assert.ok(g.prims.some((p) => p.text === "G1") && g.prims.some((p) => p.text === "G2"), "column groups drawn");
  checkPrims("groups", g, 960, 540);
  // 行グループ(2 階層の行見出し): 3+1 の不揃いな span でも和が合えば描画され、版面内に収まる
  const rg = SlideLayout.layout(
    SlideLayout.normalizeSpec({
      title: "t",
      body: {
        type: "table",
        colHeaders: ["現状", "対策"],
        rowGroups: [{ text: "上期実績", span: 3 }, { text: "下期計画", span: 1 }],
        rows: [
          { head: "受注", cells: ["+18%", "継続"] },
          { head: "粗利率", cells: ["▲2.0pt", "是正"] },
          { head: "パイプライン", cells: ["92%", "強化"] },
          { head: "リスク", cells: ["反動減", "監視"] },
        ],
      },
    }),
    { width: 960, height: 540 }
  );
  assert.ok(rg.prims.some((p) => p.text === "上期実績") && rg.prims.some((p) => p.text === "下期計画"), "row groups drawn");
  checkPrims("rowGroups", rg, 960, 540);
  const q = SlideLayout.layout(SlideLayout.normalizeSpec(SAMPLES.quadrant), { width: 960, height: 540 });
  assert.ok(q.prims.some((p) => p.rotation === 270 && /重要度/.test(p.text)), "y axis label rotated");
  const gt = SlideLayout.layout(SlideLayout.normalizeSpec(SAMPLES.gantt), { width: 960, height: 540 });
  assert.ok(gt.prims.filter((p) => p.kind === "rect" && p.fill === "#404040" && !p.text).length >= 8, "gantt dark cells");
  // 塗りのある矩形に枠線が無い / 角丸なし(全サンプル)
  for (const [name, raw] of Object.entries(SAMPLES)) {
    if (raw.sections) continue;
    const l = SlideLayout.layout(SlideLayout.normalizeSpec(raw), { width: 960, height: 540 });
    for (const p of l.prims) if (p.kind === "rect" && p.fill) assert.ok(!p.line, `${name}: filled box has no border`);
    assert.ok(!l.prims.some((p) => p.shape === "roundRect"), `${name}: no rounded rectangles`);
  }
  checks++;
}

console.log(`layout tests OK (${checks} checks)`);
