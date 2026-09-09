"use strict";
/**
 * 品質ゲート: 実際の LLM 応答(1 回の推論)を固定した fixture に対して、
 * 「多様な入力量・種類でも読めるスライドになっているか」を機械的に検査する。
 *   node test/quality.test.js
 *
 * fixture は scripts/capture-specs.js が本物の LLM から取って保存する
 * (test/fixtures/specs/<name>.json = { name, genre, prompt, response })。
 * LLM は毎回同じ答えを返さないので、テスト自体は保存した応答に対して決定論的に動かす。
 * fixture を取り直したいときだけ capture-specs.js を実行する。
 */
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const SlideLayout = require("../public/layout.js");
const coverage = require("../server/coverage");

const DIR = path.join(__dirname, "fixtures", "specs");
const INPUTS = require("./fixtures/inputs.json");
const W = 960,
  H = 540;

// ---- ゲート ----------------------------------------------------------------
// 1 枚のスライドが満たすべき条件。閾値はここに集約する(緩めるときは理由をコメントに書く)
const GATE = {
  minBodyFont: 14, // 本文の最小サイズ。注記・格子セルは別枠(下の allow)
  minCellFont: 12, // 格子・表のセルはここまで
  fillMin: 0.3, // これ未満は疎すぎ(版面が余っている)
  fillMinShort: 0.22, // 200 字未満の入力は素材そのものが少ないので、埋まらないのは版面の責任ではない
  shortInput: 200,
  fillMax: 0.99, // これを超えると詰まりすぎ
  coverage: 0.75, // 入力にあった数値・固有名詞のうち、spec に残っている割合
  titleMax: 24, // title は 10〜18 字を指示。24 字を超えたら説明文になっている
  leadLines: 2, // lead は版面で 2 行以内
  maxSlides: 4, // 既定は 1〜2 枚。読める大きさで入らないときだけエンジンが増やす(server.js の SPLIT_CAP と同じ)
};

const failures = [];
function check(name, ok, message) {
  if (!ok) failures.push(name + ": " + message);
}

/** スライド 1 枚を検査する */
function gateSlide(name, spec) {
  const lay = SlideLayout.layout(spec, { width: W, height: H });

  // G1 版面の外に出ていない
  for (const p of lay.prims) {
    if (p.kind === "line") continue;
    if (p.x == null || p.w == null) continue;
    const rot = (((p.rotation || 0) % 180) + 180) % 180;
    const box = rot === 90 ? { x: p.x + p.w / 2 - p.h / 2, y: p.y + p.h / 2 - p.w / 2, w: p.h, h: p.w } : p;
    check(name, box.x >= -0.5 && box.y >= -0.5 && box.x + box.w <= W + 0.5 && box.y + box.h <= H + 0.5, `版面の外に出た [${String(p.text || p.shape || p.kind).slice(0, 16)}]`);
  }

  // G2 読める大きさか
  const noteRoles = ["footnote", "note", "caption", "kicker", "barlabel", "axis"];
  for (const p of lay.prims) {
    if (!p.text || !p.fontSize || p.kind !== "rect") continue;
    if (noteRoles.includes(p.role)) continue;
    const floor = ["matrixcell", "tablecell", "orglabel", "ganttlabel"].includes(p.role) ? GATE.minCellFont : GATE.minBodyFont;
    check(name, p.fontSize >= floor, `${p.fontSize}pt < ${floor}pt [${p.text.slice(0, 16)}]`);
  }

  // G3 内容が落ちていない・溢れていない
  const bad = lay.warnings.filter((w) => /content dropped|overflow|too small|too wide|failed/.test(w));
  check(name, bad.length === 0, "警告: " + bad.join(" / "));

  // G4 パネルの箇条書きは 1 シェイプ(1 行ずつシェイプに分けない。生成後の編集性)。
  //    格子のセルはセルごとに 1 シェイプなので数えない
  const bulletShapes = lay.prims.filter((p) => p.bullets && p.text && p.role !== "matrixcell" && p.role !== "tablecell");
  const panels = spec.panelCount || 1;
  check(name, bulletShapes.length <= panels, `箇条書きが ${bulletShapes.length} シェイプに分かれた(パネル ${panels})`);
  // ラベル列のように 1 行を 2 シェイプに割る描き方も禁止(同じ理由)
  check(name, !lay.prims.some((p) => p.role === "listlabel"), "1 行が 2 シェイプに割れている(listlabel)");

  // G8 伝達力: タイトルは簡潔、リードは 2 行以内、主役が名指しされていれば強調が付いている
  const title = String(spec.title || "").replace(/\s+/g, "");
  check(name, title.length <= GATE.titleMax, `title が ${title.length} 字(${GATE.titleMax} 字以内)`);
  const leadPrim = lay.prims.find((p) => p.role === "lead");
  if (leadPrim) {
    const lines = SlideLayout.estimateLines(leadPrim.text, leadPrim.w - 16, leadPrim.fontSize);
    check(name, lines <= GATE.leadLines, `lead が ${lines} 行(${GATE.leadLines} 行以内)`);
  }
  const mention = ((spec.title || "") + (spec.lead || "")).replace(/\s+/g, "");
  const tables = [];
  (function walk(n) {
    if (!n || typeof n !== "object") return;
    if (n.type === "table" || n.type === "ntable") tables.push(n);
    (n.rows || n.cols || []).forEach(walk);
  })(spec.body);
  for (const tb of tables) {
    const named = (tb.rows || []).filter((rw) => rw.head && mention.includes(String(rw.head).replace(/\*\*/g, "").replace(/\s+/g, "")));
    if (named.length) check(name, named.some((rw) => rw.headDark || rw.highlight), "結論に名前が出ている行が立っていない(強調ゼロ)");
    const cells = (tb.rows || []).flatMap((rw) => rw.cells || []).map((c) => (typeof c === "string" ? c : (c && c.text) || ""));
    const unitCells = cells.filter((c) => /^\s*[\d,.]+\s*(億円|万円|%|件|回|名|社|日|本)\s*$/.test(c));
    check(name, unitCells.length < Math.max(3, cells.length * 0.5), `単位付きの数値がセルに散っている(${unitCells.length} セル)`);
  }

  // G5 版面の使い方。ネイティブ表は版面を割り付けて埋めるのが正しい姿なので詰まりすぎの判定から外す
  const hasNativeTable = lay.prims.some((p) => p.kind === "table" || p.role === "matrixcell"); // 図形の格子も版面を埋めるのが正しい
  const fillMin = (spec._shortInput ? GATE.fillMinShort : GATE.fillMin);
  check(name, lay.fill >= fillMin, `版面が疎すぎる(fill=${lay.fill.toFixed(2)} < ${fillMin})`);
  check(name, hasNativeTable || lay.fill <= GATE.fillMax, `版面が詰まりすぎ(fill=${lay.fill.toFixed(2)})`);
  return lay;
}

// ---- 実行 ------------------------------------------------------------------
if (!fs.existsSync(DIR) || !fs.readdirSync(DIR).filter((f) => f.endsWith(".json")).length) {
  console.error("品質ゲートの fixture がありません。先に取得してください:\n  node scripts/capture-specs.js");
  process.exit(1);
}

let n = 0;
for (const input of INPUTS) {
  const file = path.join(DIR, input.name + ".json");
  if (!fs.existsSync(file)) {
    failures.push(input.name + ": fixture が無い(node scripts/capture-specs.js で取得する)");
    continue;
  }
  const fx = JSON.parse(fs.readFileSync(file, "utf8"));
  const res = fx.response || {};
  const name = input.name;

  // G6 枚数: sections でないなら 1〜2 枚
  if (res.sections) {
    check(name, res.sections.length >= 3 && res.sections.length <= 8, `章立てが ${res.sections.length} 章`);
  } else {
    // 製品と同じ経路: 読める大きさで収まらない spec はエンジンが割る(LLM には聞き直さない)
    const raw = res.slides || [];
    const slides = [];
    for (const s of raw) for (const part of SlideLayout.splitToFit(SlideLayout.normalizeSpec(s), GATE.maxSlides - slides.length)) slides.push(part);
    check(name, slides.length >= 1 && slides.length <= GATE.maxSlides, `${slides.length} 枚(上限 ${GATE.maxSlides})`);
    for (let i = 0; i < slides.length; i++) gateSlide(name + "#" + (i + 1), Object.assign(slides[i], { _shortInput: input.prompt.length < GATE.shortInput }));

    // G7 入力の事実がどれだけ残っているか(1 回の推論で落とさないこと)
    const want = coverage.atoms(input.prompt).concat(coverage.terms(input.prompt));
    if (want.length >= 8) {
      const text = JSON.stringify(slides);
      const lost = coverage.missing(coverage.atoms(input.prompt), text).concat(coverage.missingTerms(coverage.terms(input.prompt), text));
      const kept = (want.length - lost.length) / want.length;
      check(name, kept >= GATE.coverage, `事実の保持率 ${(kept * 100).toFixed(0)}% < ${GATE.coverage * 100}%(欠落 ${lost.slice(0, 5).join(" / ")})`);
    }
  }
  n++;
}

if (failures.length) {
  console.error(`quality gate FAILED (${failures.length} 件 / ${n} ケース)`);
  for (const f of failures) console.error("  - " + f);
  process.exit(1);
}
console.log(`quality tests OK (${n} ケース)`);
