"use strict";
// One inference → {layout, data} → static compile on the golden master.
// The model never sees coordinates, shape ids or Office.js; it chooses one of the
// 32 golden layouts and fills its semantic data. The catalog is derived from the
// golden samples (abbreviated), so the schema shown is exactly what the compiler accepts.
const { performance } = require("perf_hooks");
const author = require("./index");
const { SAMPLES } = require("./samples");

const GUIDE = `
# 型の選び方(入力の性質 → layout)
- 結論と根拠と数値効果の一覧 → 01 / 課題と打ち手の二項対立 → 02 / 時系列の数値(2〜3系列) → 03 / 差分の要因分解(始点・要因・終点の数値) → 04
- 論点×仮説×検証の計画 → 05 / 複数案の比較評価 → 06 / 施策の優先順位付け(効果×難易度) → 07 / 部門や領域ごとの課題・原因・打ち手 → 08
- 施策や項目の長い説明の列挙 → 09 / 短い原則・方針が 3〜4 個 → 10 / 業務の段階(入力・処理・出力) → 11 / 部門をまたぐ業務フロー → 12
- 工程の期間と担当 → 13 / 組織体制と会議体 → 14 / 現状と将来の対比 → 15 / 多数の施策の一覧表(担当・時期・数値) → 16
- 進捗と KPI の月次報告 → 17 / 問いの分解と検証状況 → 18 / 週次の定例報告 → 19 / アンケート結果 → 20 / 論点の管理表 → 21
- 他社事例 → 22 / プロジェクトの進め方(フェーズ) → 23 / 成果物の一覧 → 24 / 目的とゴール → 25 / 決定してほしい事項 → 26
- リスク一覧 → 27 / シナリオ別の結果 → 28 / 指標×期間の状態 → 29 / 2 つの量の対比 → 30 / 順位付きの数値比較 → 31 / シナリオ別の推移 → 32
# 書き方
- title は結論を数値付きの常体で 70 字以内。lead はタイトルを繰り返さない根拠か整理の仕方を 100 字以内。
- 入力の事実・数値・固有名詞・条件だけを使う。入力にない数値・日付・担当・効果・出典を作らない。出典が無ければ source は "" にする。
- 例にある文字列は形式の見本であり、内容は入力から書き換える。配列の件数は入力に合わせて増減してよい(例の 0.5〜1.5 倍が目安、各型の許容範囲内)。
- "**…**" で部分太字。"#…" で始まる文字列は太字の小見出し。[文, 1] は第 2 階層の箇条書き。
- 数値は数値型(values, value, impact, x, y, at, start, end)で、文字列の数値表記(label, result, effect)は単位付きの文字列。
- 入力に足りない要素は空文字や空配列にせず、その要素を持たない型を選ぶ。どの型にも当てはまらない場合は最も近い型を選ぶ。
`;

function abbreviate(v, depth = 0) {
  if (Array.isArray(v)) { const n = v.length > 3 ? 2 : v.length; return v.slice(0, n).map((x) => abbreviate(x, depth + 1)).concat(v.length > n ? ["…"] : []); }
  if (v && typeof v === "object") return Object.fromEntries(Object.entries(v).map(([k, x]) => [k, abbreviate(x, depth + 1)]));
  if (typeof v === "string") return v.length > 26 ? v.slice(0, 24) + "…" : v;
  return v;
}
function catalog(ids) {
  return ids.map((id) => `## layout ${id}: ${author.LAYOUTS[id].name}\n${JSON.stringify(abbreviate(SAMPLES[id]))}`).join("\n");
}
function messages(input, { layout } = {}) {
  const ids = layout ? [layout] : Object.keys(author.LAYOUTS).sort();
  const system = `素材(メモ・議事録・報告)を、コンサル資料の 1 枚に構造化する。出力は JSON のみ: {"layout":"NN","data":{…}}。data の形は選んだ layout の例に厳密に合わせる(キー名を変えない・足さない)。${layout ? `layout は必ず "${layout}" にする。` : ""}
${GUIDE}
# 型カタログ(id: 名前、data の例)
${catalog(ids)}`;
  return [{ role: "system", content: system }, { role: "user", content: String(input) }];
}
/** Geometry-dependent or unknowable fields the model must not decide. */
function sanitize(id, d) {
  if (d.annotation && typeof d.annotation === "object") delete d.annotation.line; // the pointer line is calibrated to the golden data only
  if (d.today && (typeof d.today !== "object" || !String(d.today.label || "").trim())) d.today = null;
  if (Array.isArray(d.milestones)) d.milestones = d.milestones.filter((m) => m && typeof m.label === "string" && m.label.trim());
  for (const k of ["widths"]) if (k in d && id !== "08" && id !== "16" && id !== "21") delete d[k];
}
function compile(plan, opts = {}) {
  if (!plan || typeof plan !== "object") throw Error("Model returned no plan");
  if (plan.error) throw Error(String(plan.error));
  const id = String(plan.layout || "").padStart(2, "0");
  if (!author.LAYOUTS[id]) throw Error("Unknown layout " + plan.layout);
  const data = plan.data;
  if (!data || typeof data !== "object" || typeof data.title !== "string" || !data.title.trim()) throw Error("data.title is required");
  if (typeof data.lead !== "string") data.lead = "";
  if (typeof data.source !== "string") data.source = "";
  sanitize(id, data);
  const warnings = [];
  const started = performance.now();
  const built = author.build(id, data, { fit: { floor: 8, log: (w) => warnings.push(w) }, page: opts.page || "" });
  return { buffer: built.buffer, compileMs: +(performance.now() - started).toFixed(3), layout: id, warnings };
}
async function generate(input, { model, chat, layout, page } = {}) {
  if (typeof input !== "string" || !input.trim()) throw Error("Input is empty");
  if (layout) { layout = String(layout).padStart(2, "0"); if (!author.LAYOUTS[layout]) throw Error("Unknown layout " + layout); }
  const started = performance.now();
  const response = await (chat || require("../../llm").chat)({ messages: messages(input, { layout }), model, maxTokens: 6000, temperature: 0.2, jsonMode: true });
  const inferenceMs = performance.now() - started;
  const plan = require("../../prompt").extractJson(response.content);
  const compiled = compile(plan, { page });
  return { ...compiled, plan, inferenceCalls: 1, inferenceMs: Math.round(inferenceMs), model: response.model, usage: response.usage };
}
module.exports = { messages, compile, generate, catalog };
