"use strict";
/**
 * 入力の取りこぼし検出。
 *
 * 「コンテキストを増やすことは許されないが、与えられた内容はなるべく漏れなく書く」という要件を
 * 機械的に確かめるために、入力から落としてはいけない原子(数値・割合・期間・英字の固有名詞)を抜き、
 * 生成された spec に現れるかを見る。落ちていればサーバが 1 回だけ聞き直す。
 */

/** 全角→半角、桁区切りと空白を落として比較しやすくする */
function fold(s) {
  return String(s || "")
    .replace(/[Ａ-Ｚａ-ｚ０-９％．，－]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0))
    .replace(/[,，\s]/g, "")
    .toUpperCase();
}

const UNIT = "%|％|割|倍|件|名|社|人|歳|億円|万円|円|路線|本|テーブル|か月|ヶ月|カ月|年度|年|月|週間|週|日|回|時間|pt|ポイント";

/** 落としてはいけない原子を抜く */
function atoms(prompt) {
  const t = String(prompt || "");
  const out = new Set();
  const re = new RegExp(String.raw`\d[\d,，.]*\s*(?:${UNIT}|\/\s*\d+)?`, "g");
  let m;
  while ((m = re.exec(t))) {
    const raw = m[0].trim();
    const digits = fold(raw).replace(/[^0-9]/g, "");
    if (!digits) continue;
    // 単位の無い裸の数字は、桁が小さいと箇条書きの番号と紛れるので 3 桁以上だけ拾う
    if (/^\d[\d,，.]*$/.test(raw) && digits.length < 3) continue;
    out.add(raw);
  }
  for (const w of t.match(/\b[A-Z][A-Za-z]{1,9}\b/g) || []) out.add(w);
  for (const w of t.match(/[A-Z]\s?社/g) || []) out.add(w);
  return [...out];
}

/** spec(JSON 文字列でよい)に現れない原子を返す */
function missing(list, specText) {
  const hay = fold(specText);
  const digitsHay = hay.replace(/[^0-9]/g, "");
  return (list || []).filter((a) => {
    const f = fold(a);
    if (hay.includes(f)) return false;
    // 図は値と単位を別に持つ(values:[1500], unit:"件")ので、数字だけの一致も可とする
    const d = f.replace(/[^0-9]/g, "");
    return !(d.length >= 3 && digitsHay.includes(d));
  });
}

/**
 * 日本語の内容語(固有名詞・専門語)。数値と違って言い換えが起こりうるので、
 * 聞き直しの判定には使わず「どれだけ落ちたか」を見る指標として使う。
 */
const STOP = new Set(
  ("必要 実施 対応 内容 場合 以上 以下 未満 状況 状態 課題 問題 結果 理由 目的 方法 検討 確認 実行 推進 改善 強化 削減 増加 減少 拡大 縮小 全体 一部 今後 現在 現状 今回 以下 上記 中心 中央 可能 影響 効果 費用 期間 体制 提案 資料 情報 数値 項目 観点 前提 想定 各種 その他 一時 全社 社内 業務 事業 会社 部門 担当 顧客 市場 収支 運用 管理 導入 開発 設計 移行 停止 予定 目安 傾向 比率 割合 単価 総額 合計 平均").split()
);

/** 指示文(「〜を 1 枚にまとめてください」など)は内容ではないので、内容語の対象から外す */
const INSTRUCTION = /(まとめて|作って|作成して|してください|して。|示して|整理して|説明して|1\s*枚|一枚|スライド|以下は|以下の|次の内容)/;
const META = new Set(
  ("スライド メモ 資料 報告 一覧 作成 整理 記載 以下 内容 要約 タイトル レポート 本文 見出し 補足 注記 出典 フレーム アプローチ".split(" "))
);

function contentLines(prompt) {
  const lines = String(prompt || "").split(/\n/);
  const kept = lines.filter((ln, i) => {
    const t = ln.trim();
    if (!t) return false;
    // 先頭の指示行(内容を含まない依頼文)だけを外す。箇条書き・「ラベル:」の行は内容として残す
    if (i <= 1 && INSTRUCTION.test(t) && !/[:：]/.test(t.replace(/[:：]\s*$/, ""))) return false;
    return true;
  });
  return (kept.length ? kept : lines).join("\n");
}

function terms(prompt) {
  const t = contentLines(prompt);
  const out = new Set();
  // カタカナ語(3 字以上)。「デマンド交通」のように直後の漢字までを 1 語として拾う
  for (const w of t.match(/[ァ-ヴー]{3,}(?:[一-龥]{1,4})?/g) || []) out.add(w);
  // 漢字の複合語(2〜6 字)。直後のカタカナまで含める(「兼業モデル」)
  for (const w of t.match(/[一-龥]{2,6}(?:[ァ-ヴー]{2,})?/g) || []) out.add(w);
  return [...out].filter((w) => w.length >= 3 && !STOP.has(w) && !META.has(w));
}

/** 落ちた内容語。部分一致(語の一部が本文に残っていれば載っているとみなす)で緩く判定する */
function missingTerms(list, specText) {
  const hay = fold(specText);
  return (list || []).filter((w) => {
    const f = fold(w);
    if (hay.includes(f)) return false;
    // 「兼業モデル」が「兼業」だけで載っている場合も、意味は残っているので落ちたとはみなさない
    for (let n = f.length - 1; n >= 3; n--) if (hay.includes(f.slice(0, n))) return false;
    return true;
  });
}

module.exports = { atoms, missing, terms, missingTerms, fold };
