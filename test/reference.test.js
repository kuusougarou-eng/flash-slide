"use strict";
/**
 * 参照スライド解析(analyzeReference)の単体テスト。Office.js 非依存。
 *   node test/reference.test.js
 * render.js から analyzeReference を切り出して評価する(ブラウザ用 IIFE のため eval で取り込む)。
 * 3 種のデザイン(プレースホルダ / コピペ運用 / 4:3 ダークバンド)+ 表紙 + 空 を検証する。
 */
const assert = require("assert");
const fs = require("fs");
const path = require("path");

const src = fs.readFileSync(path.join(__dirname, "..", "public", "render.js"), "utf8");
const SlideLayout = require("../public/layout");
const geoKey = (s) => [s.type, Math.round(s.left), Math.round(s.top), Math.round(s.width), Math.round(s.height)].join("|");
const slice = src.slice(src.indexOf("function analyzeReference"), src.indexOf("root.SlideRender"));
// 間接 eval でグローバルに関数宣言を展開(strict のファイルスコープを避ける)。geoKey は先に注入。
const indirectEval = eval;
global.geoKey = geoKey;
indirectEval(slice + "\nglobal.analyzeReference = analyzeReference;");
const analyzeReference = global.analyzeReference;

let checks = 0;

// --- A: 標準プレースホルダ(Title + Body、リード無し、フッター/ページ番号) ---
{
  const A = [
    { id: "t", type: "Placeholder", placeholder: "Title", left: 60, top: 30, width: 840, height: 80, hasText: true, text: "事業環境の整理", fontSize: 32, bold: true, color: "#000000", fontName: "Yu Gothic" },
    { id: "b", type: "Placeholder", placeholder: "Body", left: 60, top: 130, width: 840, height: 320, hasText: true, text: "国内市場は縮小\n寡占が進行", fontSize: 18 },
    { id: "f", type: "Placeholder", placeholder: "Footer", left: 40, top: 505, width: 300, height: 24, hasText: true, text: "Company Confidential", fontSize: 10 },
    { id: "n", type: "Placeholder", placeholder: "SlideNumber", left: 880, top: 505, width: 60, height: 24, hasText: true, text: "1", fontSize: 10 },
  ];
  const r = analyzeReference(A, 960, 540);
  assert.ok(r.profile, "A: clone mode");
  assert.strictEqual(r.profile.lead, null, "A: Body placeholder is NOT treated as lead");
  assert.strictEqual(r.profile.title.reuse, true, "A: title reused");
  assert.ok(r.profile.body.h > 300, "A: body region tall");
  assert.ok(r.plan.keepKeys[geoKey(A[2])], "A: footer kept");
  assert.ok(r.plan.keepKeys[geoKey(A[3])], "A: page number kept");
  checks += 5;
}

// --- B: コピペ運用(ネイビー Meiryo タイトル + リード + 罫線 + 出典 + ページ番号) ---
{
  const navy = "#1F3A5F";
  const B = [
    { id: "t", type: "TextBox", left: 48, top: 28, width: 864, height: 44, hasText: true, text: "主力製品の粗利率は 3 年で 6pt 低下", fontSize: 24, bold: true, color: navy, fontName: "Meiryo UI" },
    { id: "rule", type: "GeometricShape", left: 48, top: 76, width: 864, height: 2, hasText: false, fillType: "Solid", fillColor: navy },
    { id: "lead", type: "TextBox", left: 48, top: 84, width: 864, height: 40, hasText: true, text: "原材料高と値引き競争が主因。価格改定と SKU 整理で来期 2pt の回復を狙う。", fontSize: 16, color: "#595959", fontName: "Meiryo UI" },
    { id: "src", type: "TextBox", left: 48, top: 470, width: 600, height: 20, hasText: true, text: "出典: 社内管理会計データ(2026年3月期)", fontSize: 10, color: "#595959" },
    { id: "cr", type: "TextBox", left: 48, top: 506, width: 400, height: 20, hasText: true, text: "© 2026 Sample Consulting Inc.", fontSize: 9, color: "#595959" },
    { id: "pg", type: "TextBox", left: 880, top: 506, width: 32, height: 20, hasText: true, text: "1", fontSize: 9, color: "#595959" },
  ];
  const r = analyzeReference(B, 960, 540);
  assert.ok(r.profile, "B: clone mode");
  assert.ok(r.profile.lead, "B: lead detected");
  assert.strictEqual(r.profile.accent, navy.toUpperCase(), "B: accent = navy from rule/title");
  assert.strictEqual(r.profile.fontName, "Meiryo UI", "B: font Meiryo");
  assert.ok(r.profile.footnote, "B: source line detected as footnote");
  assert.ok(r.plan.keepKeys[geoKey(B[4])] && r.plan.keepKeys[geoKey(B[5])], "B: copyright + page kept");
  checks += 6;
}

// --- C: 4:3 ダークバンド、白抜きタイトル、リード無し、ロゴ、ページ番号 ---
{
  const green = "#0B7A4B";
  const C = [
    { id: "band", type: "GeometricShape", left: 0, top: 0, width: 720, height: 64, hasText: false, fillType: "Solid", fillColor: "#2B2B2B" },
    { id: "accent", type: "GeometricShape", left: 0, top: 64, width: 720, height: 4, hasText: false, fillType: "Solid", fillColor: green },
    { id: "t", type: "TextBox", left: 30, top: 8, width: 660, height: 48, hasText: true, text: "施策ロードマップ", fontSize: 22, bold: true, color: "#FFFFFF", fontName: "Meiryo" },
    { id: "body", type: "TextBox", left: 30, top: 90, width: 660, height: 380, hasText: true, text: "施策1\n施策2", fontSize: 16 },
    { id: "logo", type: "GeometricShape", left: 620, top: 496, width: 70, height: 24, hasText: true, text: "LOGO", fontSize: 10, fillType: "Solid", fillColor: green, color: "#FFFFFF" },
    { id: "pg", type: "TextBox", left: 30, top: 500, width: 120, height: 20, hasText: true, text: "1 / 2", fontSize: 9, color: "#888888" },
  ];
  const r = analyzeReference(C, 720, 540);
  assert.ok(r.profile, "C: clone mode");
  assert.strictEqual(r.profile.lead, null, "C: no lead");
  assert.strictEqual(r.profile.accent, green.toUpperCase(), "C: accent = green");
  assert.strictEqual(r.profile.fontName, "Meiryo", "C: font Meiryo");
  assert.ok(r.plan.keepKeys[geoKey(C[0])], "C: dark band kept");
  assert.ok(r.plan.keepKeys[geoKey(C[1])], "C: accent bar kept");
  assert.ok(r.plan.keepKeys[geoKey(C[4])], "C: logo kept");
  checks += 7;
}

// --- D: 自分で生成したスライドを参照にするケース(棒グラフ注記と出典行が両方ある) ---
{
  const D = [
    { id: "t", type: "GeometricShape", left: 36, top: 27, width: 888, height: 49, hasText: true, text: "結論のタイトル", fontSize: 26, bold: true, color: "#1A1A1A" },
    { id: "rule", type: "GeometricShape", left: 36, top: 79, width: 888, height: 2, hasText: false, fillType: "Solid", fillColor: "#FD5108" },
    { id: "lead", type: "GeometricShape", left: 36, top: 90, width: 888, height: 39, hasText: true, text: "リード文", fontSize: 18 },
    { id: "lbl", type: "GeometricShape", left: 36, top: 138, width: 450, height: 35, hasText: true, text: "見出し", fontSize: 20, bold: true },
    { id: "lblrule", type: "GeometricShape", left: 36, top: 174, width: 450, height: 2, hasText: false, fillType: "Solid", fillColor: "#404040" },
    { id: "bar", type: "GeometricShape", left: 200, top: 200, width: 300, height: 30, hasText: false, fillType: "Solid", fillColor: "#D9D9D9" },
    { id: "note", type: "GeometricShape", left: 36, top: 462, width: 450, height: 24, hasText: true, text: "n=1,488(未着手企業)", fontSize: 12 },
    { id: "src", type: "GeometricShape", left: 36, top: 495, width: 888, height: 18, hasText: true, text: "出典: 実態調査(2026年)", fontSize: 11 },
    { id: "callbar", type: "GeometricShape", left: 563, top: 398, width: 5, height: 91, hasText: false, fillType: "Solid", fillColor: "#FD5108" },
  ];
  const r = analyzeReference(D, 960, 540);
  assert.ok(r.profile, "D: clone mode");
  assert.strictEqual(r.plan.footnoteKey, geoKey(D[7]), "D: bottom-most small text is the footnote (not the chart note)");
  assert.ok(!r.plan.keepKeys[geoKey(D[4])], "D: body heading rule NOT kept");
  assert.ok(!r.plan.keepKeys[geoKey(D[5])], "D: body bar NOT kept");
  assert.ok(!r.plan.keepKeys[geoKey(D[6])], "D: chart note NOT kept");
  assert.ok(!r.plan.keepKeys[geoKey(D[8])], "D: callout bar NOT kept");
  assert.ok(r.plan.keepKeys[geoKey(D[1])], "D: title rule kept (header zone)");
  checks += 7;
}

// --- 表紙(CenterTitle)/ 空 は既定レイアウトへ ---
{
  const cover = [{ id: "ct", type: "Placeholder", placeholder: "CenterTitle", left: 100, top: 200, width: 760, height: 120, hasText: true, text: "提案書", fontSize: 40 }];
  assert.strictEqual(analyzeReference(cover, 960, 540).profile, null, "cover → layout");
  assert.strictEqual(analyzeReference([], 960, 540).profile, null, "empty → layout");
  // タイトルが下半分にしかない → 表紙扱い
  const lowTitle = [{ id: "t", type: "TextBox", left: 60, top: 300, width: 840, height: 80, hasText: true, text: "タイトル", fontSize: 32 }];
  assert.strictEqual(analyzeReference(lowTitle, 960, 540).profile, null, "low title → layout");
  checks += 3;
}

// --- D: 実務テンプレ(Title + Subtitle をリード文として使い、本文はシェイプ群)→ clone。Title + Subtitle だけ → 表紙 ---
{
  const D = [
    { id: "t", type: "Placeholder", placeholder: "Title", left: 32, top: 32, width: 896, height: 40, hasText: true, text: "プロジェクトアプローチ", fontSize: 32, bold: false, color: "#000000", fontName: "Georgia" },
    { id: "m", type: "Placeholder", placeholder: "Subtitle", left: 32, top: 77, width: 896, height: 62, hasText: true, text: "本プロジェクトでは以下のアプローチを取ります:", fontSize: 24, color: "#000000", fontName: "Arial" },
    { id: "h1", type: "GeometricShape", left: 166, top: 161, width: 240, height: 28, hasText: true, text: "概要", fontSize: 14, bold: true, color: "#FFFFFF", fillType: "Solid", fillColor: "#D04A02" },
    { id: "b1", type: "TextBox", left: 166, top: 202, width: 240, height: 70, hasText: true, text: "AI PoC の現状と課題を明確化する", fontSize: 14 },
    { id: "v1", type: "GeometricShape", left: 60, top: 174, width: 70, height: 126, hasText: true, text: "現状把握", fontSize: 14, bold: true, fillType: "Solid", fillColor: "#D9D9D9" },
    { id: "brandmark", type: "Placeholder", placeholder: "Footer", left: 32, top: 512, width: 60, height: 14, hasText: true, text: "ACME", fontSize: 9 },
    { id: "n", type: "Placeholder", placeholder: "SlideNumber", left: 919, top: 513, width: 9, height: 9, hasText: true, text: "1", fontSize: 7.5 },
  ];
  const r = analyzeReference(D, 960, 540);
  assert.ok(r.profile, "D: Subtitle used as lead → clone mode (not a cover)");
  assert.ok(r.profile.lead && r.profile.lead.y === 77, "D: Subtitle placeholder becomes the lead slot");
  assert.strictEqual(r.profile.accent, "#D04A02", "D: accent from header fill");
  assert.strictEqual(r.profile.bodyFontSize, 14, "D: body font size taken from the deck");
  assert.strictEqual(r.profile.headFontSize, 14, "D: heading font size taken from the deck");
  assert.ok(r.profile.body.y > 139 && r.profile.body.h > 300, "D: body below lead");
  const cover = analyzeReference(D.slice(0, 2).concat(D.slice(5)), 960, 540);
  assert.strictEqual(cover.profile, null, "D: Title + Subtitle only → cover");
  checks += 5;
}

// --- E: 全面ダーク背景。地を消してしまうとテンプレートが壊れ、白地に白文字で読めなくなる ---
{
  const E = [
    { id: "bg", type: "Rectangle", left: 0, top: 0, width: 960, height: 540, hasText: false, fillType: "Solid", fillColor: "#14161A" },
    { id: "t", type: "TextBox", left: 48, top: 40, width: 864, height: 44, hasText: true, text: "既存スライド: 需要構造の転換", fontSize: 28, bold: true, color: "#F5F5F5" },
    { id: "l", type: "TextBox", left: 48, top: 92, width: 864, height: 28, hasText: true, text: "主要顧客の購買行動が四半期で入れ替わっている", fontSize: 14, color: "#B0B6C0" },
    { id: "b", type: "TextBox", left: 48, top: 150, width: 864, height: 320, hasText: true, text: "本文ダミー\n・需要の分散", fontSize: 16, color: "#F5F5F5" },
    { id: "src", type: "TextBox", left: 48, top: 500, width: 400, height: 20, hasText: true, text: "出典: 社内データ", fontSize: 9 },
  ];
  const r = analyzeReference(E, 960, 540);
  assert.ok(r.profile, "E: 全面ダークでも参照として使える");
  assert.strictEqual(r.profile.darkBackground, true, "E: 暗い地として検出する");
  assert.ok(r.plan.keepKeys[geoKey(E[0])], "E: 全面の地を保持する");
  const lay = SlideLayout.layout(SlideLayout.normalizeSpec({ title: "t", body: { type: "cell", items: ["a", "b"] } }), { width: 960, height: 540, profile: r.profile, palette: { accent: "none" } });
  const body = lay.prims.filter((x) => x.body && x.text);
  assert.ok(body.length && body.every((x) => x.color !== "#1A1A1A"), "E: 暗い地では本文を黒で描かない");
  checks += 4;
}

// --- F: 左サイドバー。帯を消すとテンプレートの章ナビが失われ、本文の位置もずれる ---
{
  const F = [
    { id: "bar", type: "Rectangle", left: 0, top: 0, width: 190, height: 540, hasText: false, fillType: "Solid", fillColor: "#2B3A55" },
    { id: "nav", type: "TextBox", left: 16, top: 40, width: 158, height: 200, hasText: true, text: "第 2 章\n市場環境", fontSize: 14, bold: true, color: "#FFFFFF" },
    { id: "t", type: "TextBox", left: 220, top: 44, width: 700, height: 40, hasText: true, text: "既存スライド: 価格競争の実態", fontSize: 24, bold: true },
    { id: "l", type: "TextBox", left: 220, top: 92, width: 700, height: 26, hasText: true, text: "値引き原資が粗利を圧迫している", fontSize: 13, color: "#666666" },
    { id: "b", type: "TextBox", left: 220, top: 140, width: 700, height: 330, hasText: true, text: "本文ダミー\n・平均値引き率", fontSize: 14 },
    { id: "src", type: "TextBox", left: 220, top: 500, width: 400, height: 20, hasText: true, text: "出典: 販売管理システム", fontSize: 9 },
  ];
  const r = analyzeReference(F, 960, 540);
  assert.ok(r.profile, "F: サイドバー付きでも参照として使える");
  assert.strictEqual(r.profile.darkBackground, false, "F: 帯があるだけでは暗い地としない");
  assert.ok(r.plan.keepKeys[geoKey(F[0])], "F: 縦帯を保持する");
  assert.ok(r.plan.keepKeys[geoKey(F[1])], "F: 帯の上の章ラベルも保持する");
  assert.ok(r.profile.body.x >= 210 && r.profile.body.w <= 720, "F: 本文領域は帯の右から始まる");
  checks += 5;
}

console.log(`reference tests OK (${checks} checks)`);
