"use strict";
/**
 * LLM へ渡すシステムプロンプトと、応答 JSON の抽出・正規化。
 * 方針: 最初の推論で 1/2/3 パネルと主図を選択。図の自由な入れ子や自動強調は禁止。
 * 色・座標・pt は LLM に出させない(layout.js の STYLE トークンが決める)。
 * デザイン規律は consulting-pptx-skill の slide-rules を単スライド向けに要約し、プロンプトとエンジンの両方で担保。
 */
const icons = require("./icons");

function iconList() {
  return icons.available().join(", ");
}

function schemaDoc() {
  return `
# 最初に決めること
入力の意味関係から panelCount を 1 / 2 / 3 のいずれかに決める。
パネルとは横方向に等分した領域。2・3 パネルはそれぞれ中央のサブタイトルと黒下線を持つ。
パネル分割は図の種類ではない。1 枚に独立した図のデザインは 1 種類だけ。
複数パネルの各領域は基本テキストだけ。格子・グラフ・ステップ・ガントを半分のパネルの下に押し込まない。

# JSON の形式（JSON だけを返す）
1 パネル: {"panelCount":1,"title","lead"?,"footnote"?,"body":図またはcell,"note"?:短い補足文字列}
2・3 パネル: {"panelCount":2|3,"title","lead"?,"footnote"?,"panels":[{"head","text"? | "items":[文字列または子箇条書きの配列]},…]}
複数枚が必要なら {"slides":[上の形式,…]}。最大 2 枚。それでも収まらないときは {"sections":[{"title","summary","slides":1},…]}（3〜6 章）。
- title: 主語と述語を持つ結論の常体文。24〜40 字程度。見出しラベルの前置きや内容と無関係な個数を入れない。
- lead: title と重ならない根拠を短い 1 文で。不要なら空文字。
- footnote: 入力にある出典・注記だけ。
- head: パネルのサブタイトル。2〜3 パネルの head は同じ粒度に揃える。
- text / items: 事実・数値・条件を残す。子の箇条書きは ["内部要因 50%",["商品の陳腐化 15%",…]] のように書ける。
- note: 1 パネルの主図の下に置く、主図と幅を揃えた無地のテキストボックス 1 個だけ。短い補足・条件に限る。
  図だけで伝わるなら省略する。title/lead/本文を反復する示唆帯は禁止。2・3 パネルに note は付けない。
- rows / cols の自由な入れ子は出力しない。座標・色・文字サイズ・塗り・highlight・太字マークも出力しない。
  優先する打ち手は文章で述べる。右側・先頭・最重要というだけで黒い箱やハイライトを指定しない。

# 1 パネルで選べる主図
1. 文章: {"type":"cell","head"?,"text"? | "items":[…]}
2. マトリクス: {"type":"table","corner"?,"colHeaders":[…],"rows":[{"head"?,"cells":[文字列,…]},…],"numbered"?,"headShape"?,"colGroups"?,"axes"?}
   n×m のセルを配置し、行・列名には背景色を付ける。列数 m<=2 は列名の帯を、行数 n<=2 は行名の帯を作らない。
   行・列名が区別に必要な場合は本文セルに意味を含める。情報を削除しない。
   順序のある行名は headShape:"chevron" でホームベース型の五角形にし、番号を付ける。
   完全に並列なら番号は不要（必要なら A/B/C）。2×2 の座標分類は axes:{x,y} を使い、意味のある軸名は残す。
   概要・タスク・アウトプットなど共通の観点がある工程は、1 つのマトリクスとして表す。
3. ステップ（マトリクスの一方向の形）: {"type":"sequence","steps":[{"head","text"},…]}
   2〜8 段。エンジンが 01/02/03 の番号付き五角形の見出しを作る。
   短文なら左→右、説明が長ければ上→下。方向は実際の文字量と幅からエンジンが決める。
   切替条件などは該当ステップの text に含める。下にもう一つ別のデザインを足さない。
4. ガント: {"type":"gantt","groups":[{"text","span"}],"periods":[…],"tasks":[{"label","start","end","note"?}],"milestones"?:[{"label","at"}]}
   groups は年度と月、月と週など上位の列名。span は periods の列数で、groups 全体の span 合計は periods.length と一致させる。start/end/at は periods のラベルまたは 0 起点の位置。
   periods は最初の工程の開始期間から最後の工程の終了期間まで、時間順に途切れなく並べる(途中の月から始めない。入力が 4 月〜翌 4 月なら 13 か月)。
   同じ月名が複数年度に出るときは年を含むラベルか位置で識別する。セルの代わりに五角形の期間帯を描く。
   役割はタスクの note に含め、体制の短い補足は全幅 note 一つにまとめる。下に体制図や複数の箱を足さない。
5. 数値図: {"type":"line","unit","labels":[…],"series":[{"label","values":[…]}]}
   / {"type":"bars"|"column","unit","items":[{"label","value"}]}
   / {"type":"stacked","parts":[…],"items":[{"label","values":[…]}]}
   / {"type":"kpi","value","label"}
   推移・量の差・構成比が主題のときに選ぶ。数値が文章中にあるだけで図に変換しない。
   年次・月次などの系列(3 点以上)、複数項目の構成比・量の比較が入力にあるときは、それを主図にする(箇条書きに数値を並べない)。
6. 体制図: {"type":"org","root":{"label","sub"?,"children":[同形式,…]}}
   接続は報告関係。人数・役割・開催頻度は sub。深さ 4、18 箱まで。動的な位置と接続線はエンジンで計算。
7. 密な表: {"type":"ntable","corner"?,"colHeaders":[…],"rows":[{"head","cells":[…]}]}
   8 行以上、または 36 セル以上は PowerPoint の Table オブジェクトを使う。文字の表も対象。
   最大 12 行・6 列。超えるときはスライドを分割する。

# 意味からの選択
- 原因と打ち手、現状と今後などの二項対立は 2 パネルのテキスト。片側を黒く塗らない。
- 3 つの並列論点は 3 パネル。ただし長い説明が続く列挙なら 1 パネルの縦マトリクスを選ぶ。
- 比較の共通項目がある場合は 1 パネルのマトリクス。別のパネル内に格子を置かない。
- ステップ、ガント、体制図、数値図はそれぞれ単独で版面を使う。
- 項目数を増やして版面を埋めない。余白があれば文字サイズと縦の配置をエンジンが調整する。
- 1 スライド 1 メッセージ。入力の事実・数値・固有名詞・条件を落とさず、入力にないものを作らない。全体の変化率（例：前年比 30% 減）と要因内訳の構成比は別の情報なので両方残す。出力前に入力の各数値・条件が残っているか照合する。
- 細部が多い場合は本文 1 枚 + 詳細 1 枚。同じ内容を繰り返さない。
- 現在の実装にない図種を cell の入れ子で擬似的に描かない。適切な既存の 1 図か文章にする。
- 文体は常体。同じ階層の箇条書きは語尾を揃える。出力言語は入力と同じ。

# 例: 原因と打ち手
{"panelCount":2,"title":"売上減少への対応は内部要因の改善を優先する","lead":"売上は前年比 30% 減となった","panels":[{"head":"減少の要因","items":["内部要因 50%：商品の陳腐化 15%、品質問題 10%、ラインナップ不足 5%、営業人員の減少 10%、営業効率の低下 10%","外部要因 30%：市場縮小 10%、規制強化 5%、新規参入 10%、価格競争 5%","その他 20%：季節要因 10%、一時的要因 10%"]},{"head":"優先する打ち手","text":"内部要因の改善、特に商品力への投資を優先する"}]}

# 例: ステップだけで伝える
{"panelCount":1,"title":"移行は検証条件を満たした段階から順に進める","body":{"type":"sequence","steps":[{"head":"現状把握","text":"1,200 テーブルを棚卸しし依存関係を可視化する"},{"head":"参照系移行","text":"BI ダッシュボード 38 本を移し、2 週間並行稼働する。稼働率 99.5% 以上、再処理は週 3 件以下を 2 週連続で満たすことを確認する"},{"head":"更新系移行","text":"基幹連携バッチ 64 本を移し、切替リハーサルを 2 回行う"},{"head":"旧環境停止","text":"年 1,800 万円のライセンスを解約し、運用手順を更新する"}]}}
`;
}

function buildMessages({ prompt, context, hint, maxSlides }) {
  const sys =
    "あなたは戦略コンサルティングファームのスライド作成エンジンです。与えられた素材から、1 回の推論で「1 スライド = 1 メッセージ」のスライド(原則 1 枚、最大 2 枚)の構造化データを JSON で返します。" +
    "モノトーン基調・強調色 1 色のシンプルなデザインで、余白と整列、左→右の読み順、情報の階層を重視します。要素間の関係性(列挙・因果・順序・比較・数値・階層)を見極め、パネル数を決め、1 枚に主となる図は 1 つだけ選択します。\n" +
    schemaDoc();

  let user = "【今日の日付】" + new Date().toISOString().slice(0, 10) + "(期間や「来月」の計算にだけ使う。本文・出典・footnote に日付そのものを書かない)\n\n";
  if (context && context.trim()) user += "【素材(コンテキスト)】\n" + context.trim() + "\n\n";
  user += "【指示】\n" + (prompt || "").trim();
  if (hint && hint.trim() && hint !== "auto") user += "\n\n【レイアウトの希望(弱い制約。内容に合わなければ従わなくてよい)】\n" + hint.trim();
  user += `\n\n【枚数】原則 1 枚。独立した 2 つのメッセージがあり 1 枚に収まらないときだけ 2 枚(slides)。それでも収まらない量なら sections を返す。最大 ${maxSlides || 2} 枚。`;
  user += "\n\nJSON のみを出力:";
  return [
    { role: "system", content: sys },
    { role: "user", content: user },
  ];
}

function extractJson(text) {
  if (!text) throw new Error("empty LLM response");
  let t = String(text).trim();
  const fence = t.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fence) t = fence[1].trim();
  try {
    return JSON.parse(t);
  } catch (_) {}
  const start = t.indexOf("{");
  const end = t.lastIndexOf("}");
  if (start >= 0 && end > start) {
    const inner = t.slice(start, end + 1);
    try {
      return JSON.parse(inner);
    } catch (_) {}
    // 軽微な崩れの修復: 行頭の箇条書きダッシュ(Gemini が "- \"key\": …" と出す)、末尾カンマ
    const repaired = inner.replace(/\n\s*-\s+(?=")/g, "\n").replace(/,\s*([}\]])/g, "$1");
    return JSON.parse(repaired);
  }
  throw new Error("no JSON object found in LLM response: " + t.slice(0, 200));
}

function shapeResponse(raw, normalizeSpec, maxSlides) {
  if (raw && Array.isArray(raw.sections) && raw.sections.length && !(Array.isArray(raw.slides) && raw.slides.length)) {
    return {
      sections: raw.sections.slice(0, 8).map((s) => ({ title: String(s.title || "").trim(), summary: String(s.summary || s.message || "").trim(), slides: Math.max(1, Math.min(2, Number(s.slides) || 1)) })),
    };
  }
  let list = Array.isArray(raw && raw.slides) ? raw.slides : [raw];
  list = list.filter((s) => s && typeof s === "object").slice(0, maxSlides || 2);
  if (!list.length) throw new Error("LLM 応答にスライドが含まれていません");
  return { slides: list.map(normalizeSpec) };
}

module.exports = { buildMessages, extractJson, shapeResponse, schemaDoc };
