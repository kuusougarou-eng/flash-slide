"use strict";
// Distilled from the development split's design decisions, not from held-out inputs.
// No golden coordinates or complete slides are sent to the fast model.
function designGuide(){return `
# 構成を選んだ後の編集判断
- 短い事務連絡で「対象・変更点・例外・開始条件」のような説明なら、cell の部分太字ラベルで一続きに読む。項目が四つあるだけで表にしない。
- 二項対立には等価な二パネル。各本文に複数の話題があるなら短い親ラベルを **太字** にし、その根拠を子の配列にする。単一の短文を無理に二階層にしない。
- 共通の観点を持つ案の比較は table。行数が8以上またはセル数が36以上なら ntable を維持する。意味のある行名・列名・単位を残す。
- 工程は sequence。短文なら意味の合う icon と短い説明、長文なら各工程に確認条件をまとめる。順序を持たない三論点を工程に変えない。
- 時間の推移には line、値の大小には bars、期間には gantt、報告関係には org。数値とその対象の対応を入れ替えない。
- 上下の概念階層と各層の説明は pyramid。報告関係の木構造とは区別する。
- 主図を選び終えたら、余った面積を埋めるために別の要約・別の図を付け足さない。
`}
module.exports={designGuide};
