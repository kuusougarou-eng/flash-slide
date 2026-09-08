"use strict";
/**
 * LLM へ渡すシステムプロンプトと、応答 JSON の抽出・正規化。
 * 方針: 最初の推論で 1/2/3 パネルと主図を選択。図の自由な入れ子や自動強調は禁止。
 * 色・座標・pt は LLM に出させない(layout.js の STYLE トークンが決める)。
 * デザイン規律は consulting-pptx-skill の slide-rules を単スライド向けに要約し、プロンプトとエンジンの両方で担保。
 */
const icons = require("./icons");
const { designGuide } = require("./design-guide");

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
- text / items: 事実・数値・条件を残す。長い文章は **冒頭の短いラベル** で話題を示すか、文中の重要な判断・条件だけを **太字** にする。全文・全項目を太字にしない。親の論点と具体例・内訳は 2 階層の箇条書きに分けてよい。細切れにしすぎず、短い内容は一階層のままでよい。子の箇条書きは文字列の先頭に空白とハイフンを並べず JSON の配列で表す。例えば ["内部要因 50%",["商品の陳腐化 15%",…]] のように書ける。
  各項目が「短いラベル + その説明」の形になるときは、文字列に混ぜず {"label":"受注","value":"前年比 **+18%**。新規 32 社の獲得が寄与"} の形で分けてよい。表示側が幅を測り、余裕があればラベルを左の列に揃えて説明だけを折り返す。ラベルは 10 字以内の名詞にする。
  **文字数を理由に情報を削らない**。1 項目が版面で折り返す長さでも構わない(折り返しは表示側が処理する)。文章を削って詰め込む代わりに、体言止めで冗長な助詞・語尾を削り(「〜を実施する」→「〜を実施」)、主従の粒度が違うなら親の結論(体言止めの短い一文)と子の内訳・条件・数値の 2 階層に分ける。子は表示側で小さく補足として組む前提でよいので、内訳を丸めたり削ったりしない。
- note: 1 パネルの主図の下に置く、主図と幅を揃えた無地のテキストボックス 1 個だけ。短い補足・条件に限る。
  既定では note を省略する。note は本文にない条件・注記・出典に限る。結論・対応方針・要約は lead に一度だけ書く。言葉を変えた言い換えでも title/lead/本文を繰り返してはいけない。2・3 パネルに note は付けない。
- rows / cols の自由な入れ子は出力しない。座標・色・文字サイズ・塗り・highlight は出力しない。本文の部分的な **太字** は使ってよい。
  優先する打ち手は文章で述べる。右側・先頭・最重要というだけで黒い箱やハイライトを指定しない。

# 1 パネルで選べる主図
1. 文章: {"type":"cell","head"?,"text"? | "items":[…]}
2. マトリクス: {"type":"table","corner"?,"colHeaders":[…],"rows":[{"head"?,"cells":[文字列,…]},…],"numbered"?,"headShape"?,"colGroups"?,"axes"?}
   n×m のセルを配置し、行・列名には背景色を付ける。列数 m<=2 は列名の帯を、行数 n<=2 は行名の帯を作らない。
   行・列名が区別に必要な場合は本文セルに意味を含める。情報を削除しない。
   セルの文字列が複数の事実・数値を含み長くなるときは、読点で繋いだ一文に詰め込まず改行(\n)で分けた箇条書きにする。
   1 行目に短い太字見出し(記号なし)、2 行目以降を "・" で始める箇条書きにすると、見出し+箇条書きの階層として描画される。
   例: "**在庫管理領域**\n・欠品による機会損失 年間4,200万円\n・過剰在庫による廃棄損 年間1,800万円"。
   見出しが不要なら全行を "・" で始めれば見出しなしの箇条書きになる。
   行見出しは既定で長方形。単なる項目の列挙には五角形を使わない。工程の進行方向そのものを強く示す必要があるときだけ headShape:"chevron" を選べる。番号はエンジンが付けるので head に番号を書かない。
   完全に並列なら番号は不要（必要なら A/B/C）。2×2 の座標分類は axes:{x,y} を使い、意味のある軸名は残す。
   概要・タスク・アウトプットなど共通の観点がある工程は、1 つのマトリクスとして表す。
3. ステップ（マトリクスの一方向の形）: {"type":"sequence","steps":[{"head","text","icon"?},…]}
   各工程が短文（目安40文字以下）で、意味に合う絵がある場合は steps の各要素に icon を必ず出力する。例: {"head":"調査","text":"現状と課題を整理する","icon":"search"}、設計なら pencil、実行なら rocket。短文の工程図では各工程の意味に合う icon を全工程に選び、大きなピクトグラムで視覚的に理解できるようにする。単なる飾りや同じアイコンの反復は避け、適切な絵がないなら全工程で省略する。文章を削ってアイコンに置き換えない。利用可能な icon: ${iconList()}
   2〜8 段。エンジンが 01/02/03 の番号を付ける。head に番号・Step 1・第1段階等を重ねない。横並びでは矢羽根（五角形）で進行方向を示す。長文の縦配置は長方形が標準。順序のない論点をステップにしない。
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
8. ピラミッド: {"type":"pyramid","layers":[{"head","text"},…]}
   2〜5 層を上から順に指定する。左の階層形状と右の説明を同じ高さで対応させる。座標や形状の幅は指定しない。

# 意味からの選択
- 原因と打ち手、現状と今後などの二項対立は 2 パネルのテキスト。片側を黒く塗らない。
- 3 つの並列論点は 3 パネル。ただし長い説明が続く列挙なら 1 パネルの縦マトリクスを選ぶ。
- 比較の共通項目がある場合は 1 パネルのマトリクス。別のパネル内に格子を置かない。
- ステップ、ガント、体制図、数値図はそれぞれ単独で版面を使う。
- 項目数を増やして版面を埋めない。余白があれば文字サイズと縦の配置をエンジンが調整する。
- 1 スライド 1 メッセージ。入力の事実・数値・固有名詞・条件を落とさず、入力にないものを作らない。全体の変化率（例：前年比 30% 減）と要因内訳の構成比は別の情報なので両方残す。出力前に入力の各数値・条件が残っているか照合する。
- **入力にある数値は 1 つ残らず載せる**。内訳を合計に丸めない(「初期費用 4.2 億円、運用費 年 6,000 万円」を「5 年総額 7.2 億円」だけにしない)。
  1 枚に載りきらないなら slides を 2 枚にする。載せる場所は本文でも note でも 2 枚目でもよいが、消してはいけない。
- 細部が多い場合は本文 1 枚 + 詳細 1 枚。同じ内容を繰り返さない。
- 現在の実装にない図種を cell の入れ子で擬似的に描かない。適切な既存の 1 図か文章にする。
- 文体は常体。同じ階層の箇条書きは語尾を揃える。出力言語は入力と同じ。

# 例: 原因と打ち手
{"panelCount":2,"title":"売上減少への対応は内部要因の改善を優先する","lead":"売上は前年比 30% 減となった","panels":[{"head":"減少の要因","items":["**内部要因 50%**：商品の陳腐化 15%、品質問題 10%、ラインナップ不足 5%、営業人員の減少 10%、営業効率の低下 10%","**外部要因 30%**：市場縮小 10%、規制強化 5%、新規参入 10%、価格競争 5%","**その他 20%**：季節要因 10%、一時的要因 10%"]},{"head":"優先する打ち手","text":"内部要因の改善、特に商品力への投資を優先する"}]}

# 例: ステップだけで伝える
{"panelCount":1,"title":"移行は検証条件を満たした段階から順に進める","body":{"type":"sequence","steps":[{"head":"現状把握","text":"1,200 テーブルを棚卸しし依存関係を可視化する"},{"head":"参照系移行","text":"BI ダッシュボード 38 本を移し、2 週間並行稼働する。稼働率 99.5% 以上、再処理は週 3 件以下を 2 週連続で満たすことを確認する"},{"head":"更新系移行","text":"基幹連携バッチ 64 本を移し、切替リハーサルを 2 回行う"},{"head":"旧環境停止","text":"年 1,800 万円のライセンスを解約し、運用手順を更新する"}]}}
`;
}

function buildMessages({ prompt, context, hint, maxSlides, variant = 0 }) {
  const sys =
    "あなたは戦略コンサルティングファームのスライド作成エンジンです。与えられた素材から、1 回の推論で「1 スライド = 1 メッセージ」のスライド(原則 1 枚、最大 2 枚)の構造化データを JSON で返します。" +
    "モノトーン基調・強調色 1 色のシンプルなデザインで、余白と整列、左→右の読み順、情報の階層を重視します。要素間の関係性(列挙・因果・順序・比較・数値・階層)を見極め、パネル数を決め、1 枚に主となる図は 1 つだけ選択します。\n" +
    schemaDoc() + designGuide();

  let user = "【今日の日付】" + new Date().toISOString().slice(0, 10) + "(期間や「来月」の計算にだけ使う。本文・出典・footnote に日付そのものを書かない)\n\n";
  if (context && context.trim()) user += "【素材(コンテキスト)】\n" + context.trim() + "\n\n";
  user += "【指示】\n" + (prompt || "").trim();
  if (hint && hint.trim() && hint !== "auto") user += "\n\n【レイアウトの希望(弱い制約。内容に合わなければ従わなくてよい)】\n" + hint.trim();
  const alternatives = [
    "内容の意味に最も合う構成を選ぶ。2 パネルや五角形を既定にしない。",
    "比較・整理の別案。panelCount:1 の table（または密な ntable）を主図とする。行と列の観点を素材から選び、五角形は使わない。",
    "文章の階層で読む別案。panelCount:1、body.type:cell。冒頭の太字ラベルと親子の箇条書きで整理する。子項目は短い内訳なら同じ一行にまとめ、全てを一項目一行にして縦に膨らませない。パネル分割・格子・五角形は使わない。",
    "関係を見せる別案。数値なら適切なグラフ、時間ならガント、本当に順序があるなら sequence。図にできる素材がなければ意味のある 3 論点を panels で整理する。2 パネル・単純な文章列挙・比較表は選ばない。"
  ];
  const v = Math.max(0, Math.min(3, Math.trunc(Number(variant) || 0)));
  user += "\n\n【今回の案の構成方針】\n" + alternatives[v] + "\n別案でも事実・条件を保持し、構成を変えるための事象・数値・順序は捏造しない。内容に成立しない形式を無理に使わない。";
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
