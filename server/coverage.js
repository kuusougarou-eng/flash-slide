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

module.exports = { atoms, missing, fold };
