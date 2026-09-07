# 部品合成型レイアウトエンジン(cell × grid × table)— 実装計画書

- 作成日: 2026-09-06
- ステータス: **実装済み(2026-09-06)。** Phase 0〜5 を完了。As-Is は [../architecture/generation-pipeline.md](../architecture/generation-pipeline.md) を参照。
- 実装時の決定(ユーザー指示 2026-09-06): **`table` は PPT ネイティブ表ではなく「セルを組み合わせた格子」**として図形で描く(コンサル資料は表をセルで組む慣習)。ネイティブ表は `ntable` として figure に分類。
- 前提ドキュメント:
  - [../architecture/generation-pipeline.md](../architecture/generation-pipeline.md)(As-Is)
  - [2026-09-06-swimlane-node-type-proposal.md](2026-09-06-swimlane-node-type-proposal.md)(本計画で **置き換える**。§5 参照)
  - [2026-09-06-taskpane-ui-refresh-implementation-plan.md](2026-09-06-taskpane-ui-refresh-implementation-plan.md)(独立。依存なし)

---

## 0. 思想(ブレさせないための 4 行)

1. **「パターン」は静的に定義できない。** コンサル流の 16 型 + 階層構造は *分類* ではなく *部品の語彙* として使う。組み合わせは LLM の柔らかさで吸収する。
2. **推論は 1 回、描画は 1 発。** 生成→画像化→修正のループで品質を稼がない。品質は「エンジンが構造的に崩れない設計になっている」ことで担保する(検証ループではなく構成的制約)。
3. **コンサル資料の大半はマトリクスである。** `table` を最も表現力のある第一級部品にし、スイムレーン・ロードマップ・RACI・ガント・4 象限・フローマトリクスはすべて `table` の属性で出す。
4. **エンジンが持つのは「型」ではなく「トークン」。** 文字サイズ・強調・パディング・揃え・色役割を 1 箇所のデザイントークンに集約し、LLM は意味(強調する/しない、関係の向き)だけを出す。色・座標・pt は LLM に一切出させない。

## 1. 現状の何が「静的パターン」なのか(As-Is 診断)

| 箇所 | 実態 | 問題 |
|---|---|---|
| `public/layout.js` `LEAF.*`(14 関数、約 500 行) | `label`/`card`/`bullets`/`text`/`callout`/`chevrons`/`icon` がそれぞれ独自に「見出し+下罫」「薄い地」「矢羽+下段」を描く | **型 = パターン**。矢羽は横向き固定、callout は地色固定。新しい見た目 = 新しい type = 新しい関数 = プロンプト肥大、という増殖ループ |
| `server/prompt.js` 「構成の選び方(基本形 4 つ)」+ フルサンプル 3 本(約 1,800 トークン) | ①表 ②前提→帰結 ③対比 ④分析→解説 | LLM が 4 形に収束する。ユーザー指摘「右下 callout の量産」「同じ構成ばかり」の直接原因 |
| swimlane 提案 | 縦矢羽 + 3 列 = 新 type `roadmap` | 「段階 × 観点」は `table` の行見出しが矢羽形になっただけ。type を足す必要がない |
| 強調の上限・色の規律 | プロンプトで「1〜2 箇所」と *お願い* している | コードで強制していない。LLM が 5 箇所に highlight を付ければそのまま 5 箇所が塗られる |
| 階層構造(インデント/ボックス/ツリー) | `bullets` は 1 階層のみ。ネストは rows/cols の入れ子でしか表せない | 「インタビュー概要」型(非対称 3 階層)が組めない |

結論: **部品は多いのに直交していない**。部品数を減らし、属性で表現の幅を広げる。

## 2. To-Be: 4 つの葉 + 2 つの容器

```
spec  = { title, lead?, footnote?, body: Node }
Node  = Grid | Cell | Table | Figure | Arrow
Grid  = { rows:[Node…], weights? } | { cols:[Node…], weights? }        … 入れ子 3 段まで(現行どおり)
Cell  = { type:"cell", head?, text? | items?:[…], shape?, fill?, highlight?, icon?, align? }
Table = { type:"table", corner?, colHeaders:[…], rows:[{head, cells:[…], highlight?}],
          headShape?:"chevron", colGroups?:[{text, span}], highlightCol?, axes?:{x, y} }
Figure= { type:"figure", kind:"bars"|"column"|"line"|"stacked"|"kpi", … 現行フィールド }
Arrow = { type:"arrow", direction? }
```

### 2.1 `cell` — テキストを持つ唯一の部品

現行 7 型(`label`/`card`/`bullets`/`text`/`callout`/`chevrons` の 1 段/`icon`)を **1 つのレンダラ**に畳む。見た目は属性で決まる。

| 属性 | 値 | エンジンの描画規則(デザイントークンから決定) |
|---|---|---|
| `head` | 文字列 | 太字 20pt(入らなければ 18pt×2 行)。`fill` が none/light のとき下に太い罫(区切る対象=本文があるときだけ) |
| `text` / `items` | 文字列 / 配列 | 18pt 基本、14〜22pt で fit。`items` が 2 件以上で箇条書き。**`items` のネスト配列は第 2 階層**(インデント型階層) |
| `shape` | `rect`(既定) / `chevron` / `home` | 矢羽・五角形。中央揃え・塗り必須(既定 dark)・下罫なし |
| `fill` | `none`(既定) / `light` / `dark` / `tint` | none=枠も塗りもなし、light=#F2F2F2、dark=#404040+白文字、tint=強調色の 90% 薄地(旧 callout)。**塗りと枠線の同時指定は不可**(トークンで禁止) |
| `highlight` | bool | 見出し・罫・矢羽塗りを強調色に。**1 枚あたり 2 個まで**(§2.5) |
| `align` | left(既定)/center/right | 数値だけのセルは自動で right |
| `icon` | Tabler 名 | 見出し左に画像。見出し無しなら中央に単体(旧 `icon` 型) |
| `level` | 0/1 | 親 grid 内でのインデント(ボックス型階層の子要素) |

旧型は **糖衣構文(alias)** として `normNode` で desugar する。LLM・モック・既存テストは無改修で動く。

```
label{text}            → cell{head:text}
card{head,text|bullets}→ cell{head,text|items}
bullets{items}         → cell{items}
text{text}             → cell{text}
callout{text}          → cell{text, fill:"tint"}
icon{name,label}       → cell{icon:name, text:label, align:"center"}
chevrons{steps}        → rows[ cols[cell{head, shape:"chevron"}…], cols[cell{items}…] ]   (本文なしなら 1 段)
```

### 2.2 `table` — マトリクスを第一級に

現行 `table` に以下を足す。すべて **ネイティブ表 1 回の `addTable` + 少数の重ね描き**で済み、速度は変わらない。

| 拡張 | 表せる型 | 実装 |
|---|---|---|
| `headShape:"chevron"` | ロードマップ / スイムレーン / フローマトリクス(段階 × 観点) | 第 1 列を空欄で表を描き、`rowHeights` から各行の位置を取って `homePlate`/`chevron` 図形を重ねる(`colHeaders` 側に付けるなら横向き) |
| `colGroups:[{text,span}]` | 2 階層の列見出し(階層マトリクス) | 表の上に 1 段のグループ見出し帯(テキスト+群幅の細罫)を描く。セル結合 API に依存しない |
| セル値 `{text, fill?, highlight?}` | ガント(■ の塗りセル)、RACI(R/A/C/I の強調)、評価表(◎ のみ強調) | `specificCellProperties.fill` を使う(現行の highlight 実装を一般化) |
| `axes:{x, y}` | 4 象限(2×2 表 + 軸ラベル) | 表の左と下に軸ラベルと矢印線を描く |
| 行見出し無し(`rows[].head` 空 / `corner` 空) | 純粋な列挙表・会議体表 | 第 1 列幅を 0 にする |

数値セルの右揃え・非該当 "—" の薄色・最終行の罫なし・ヘッダー無地+強調色罫は現行のまま(デザイン規律の正典を維持)。

### 2.3 `grid` アルゴリズムの追加規則(すべて決定論・ミリ秒)

1. **兄弟列の行揃え**: `cols[rows[…], rows[…]]` で各列の rows 数が同じなら、行インデックスごとに最大の自然高さで揃える(列挙型パターン 3、対比型で見出し行の下端が揃う)。
2. **自動リフロー**: `cols` 直下の葉が 5 個以上なら 2 行に折る(現行 `boxes>4` の規則を一般化)。`cols[cell{shape:chevron}×n]` は n≥6 で縦フロー(`rows`)に倒す。
3. **等幅の既定**: 2 パネル(+arrow)の左右等幅強制は現行どおり。
4. **版面充填**: `FILL_RATIO_MIN`(55%)未満は垂直中央。現行どおり。

### 2.4 デザイントークン(`STYLE` 1 箇所に集約)

現行の `FONT`/`MARGIN`/`PAD`/`GAP`/`RULE_*`/`DEFAULT_PALETTE` を 1 オブジェクトにまとめ、`cell`/`table`/`figure` の全描画がここだけを参照する。

```
STYLE = {
  font: { title:26, head:20, body:18, bodyMin:14, bodyMax:22, tableMin:14, foot:11, lineHeight:1.25 },
  space:{ margin:36, pad:8, gap:20, bulletIndent:18, levelIndent:18 },
  rule: { thick:1.8, thin:0.75 },
  color:{ text, textMuted, textOnDark, line, lineLight, fillLight, fillDark, accent, accentTint },
  align:{ head:"left", chevron:"center", number:"right", tableHead:"center" },
  budget:{ highlight:2, tintCells:1 },
}
```

- **強調色オフ**: `palette.accent === "none"` で完全モノトーン。highlight は太字+`fillDark` で表現(タスクペインのスウォッチに「なし」を追加。UI 刷新計画 Phase 3 と接続)。
- 参照解析(`profile.fontName`/`accent`)は現行どおり STYLE を上書きする。

### 2.5 エンジン側で強制する規律(プロンプトの「お願い」からコードへ)

| 規律 | 現行 | To-Be |
|---|---|---|
| 強調 1〜2 箇所 | プロンプトのみ | `normalizeSpec` で highlight を先頭 2 個に切り詰め(表内・図内を含めて合算) |
| tint(旧 callout)は 1 枚 1 つ | プロンプトのみ | 2 つ目以降は `fill:"none"` に落とす |
| 塗り+枠線の併用禁止・角丸禁止 | 各 LEAF の実装任せ | `cell` レンダラの単一経路で構造的に不可能にする |
| 葉 12 個まで / 深さ 3 まで | あり | 維持 |
| タイトル常体・ラベル前置き除去 | あり(`cleanTitle`) | 維持 |

## 3. プロンプトの作り替え(「型」ではなく「関係 → 部品」)

### 3.1 型カタログの置き換え

「構成の選び方(基本形 4 つ)」を削除し、コンサル流手法論の **関係性フローチャート**を部品対応表に翻訳して渡す。これは *組み合わせ例* であって型ではない、と明記する。

```
要素間の関係           → 部品の組み方(例。内容に合わせて組み合わせてよい)
関係なし(列挙)         → cols[cell×2〜4]。説明が長ければ rows[cell×n] / cols[rows[cell{head}, cell{items}]…]
複数 → 1 つ(背景・合流) → cols[ rows[cell×n], arrow, cell ]
1 つ → 複数(拡散)       → cols[ cell, arrow, rows[cell×n] ]
順序(フロー)           → cols[cell{shape:"chevron"}×3〜5](要点は下段の cols)。5 段超は rows
比較・分類・評価         → table(行=項目、列=観点)。対象 2 つで長文なら cols[cell, cell]
段階 × 観点             → table{headShape:"chevron"}(ロードマップ・スイムレーン)
役割・会議体・スケジュール → table(セルに ■/R/A/C/I など。強調は highlight)
2 軸で分類             → table 2×2 + axes
数値                   → figure(推移=line、差=bars/column、構成比=stacked、1 つ=kpi)
前提 → 帰結             → cols[事実側(table/figure/cell), arrow, 帰結側(cell)]
階層(非対称)           → cell.items のネスト(2 段)/ rows[cell{head}, cols[cell…]]
```

### 3.2 few-shot の削減

フルサンプル 3 本 → **フル 1 本(前提→帰結+表)+ body だけの断片 4 本(各 1〜3 行)**。
狙いは (a) 4 形への収束を止める、(b) システムプロンプトを約 40% 短縮する、の 2 つ。JSON の妥当性は `jsonMode` + `normNode` で担保しているので、サンプル本数を減らしても壊れない(§6 の計測で確認)。

### 3.3 LLM に出させないもの(再確認)

色コード・座標・pt・weights の細かい数値(2 パネルは強制等幅、3 列以上のみ尊重)。出させるのは *意味*(head/text/items/関係の向き/highlight/shape)だけ。

## 4. 速度(1 プロンプト → 描画まで)への効き方

体感の大半は LLM の出力トークン待ち(現行 ≈2.2s @ 軽量モデル)。本計画で触るのは以下。

| 施策 | 効果 | 備考 |
|---|---|---|
| 語彙を 14 型 → 4 型 + 属性に圧縮 | 出力 JSON の安定化(誤った type 名の減少)、**再試行ゼロ** | 出力トークン数はほぼ同じ(内容量が支配的) |
| システムプロンプト約 40% 短縮 | TTFT(初動)わずかに改善、キャッシュ効率向上 | 入力側は元々速い |
| `maxTokens` 3000 の妥当性を計測して見直す | 上限自体は速度に無関係。過剰出力の検知用 | §6 の計測で決める |
| エンジン: `cell` 単一経路 | 変化なし(元々ミリ秒) | 目的は品質の構造的担保 |
| Office.js: `table` は 1 コール + 重ね図形 ≤ 行数 | 変化なし | `sync` 回数は現行の 3〜4 回/枚を維持 |

**やらないこと**: 2 段推論、画像化して LLM に見せる自己修正、ストリーミング分割描画(レイアウトは全体依存なので効かない)。

## 5. swimlane 提案との関係

[2026-09-06-swimlane-node-type-proposal.md](2026-09-06-swimlane-node-type-proposal.md) の決定事項のうち、

- (1)「新 Node 型を追加する」 → **撤回**。`table{headShape:"chevron"}` で表現する(§2.2)。
- (2)「列ヘッダーの色塗り不採用」 → **維持**。
- (3)「矢羽ラベルは塗り+白文字でよい」 → **維持**(`cell{shape:chevron}` の既定が dark 塗り)。

本計画の Phase 2 完了時に、同提案のステータスを「本計画に統合(superseded)」へ更新する。

## 6. 実装フェーズ(依存順)

各フェーズで `npm test` グリーン → 実機で `scripts/ui.ps1 shot` の目視、を通してから次へ進む。

### Phase 0 — 計測の土台(半日)

- [ ] `/api/generate` のログに `usage.completion_tokens` / `prompt_tokens` / ms を出す(`server.js`)。ベースライン 10 プロンプト分を `debug/bench-baseline.json` に保存
- [ ] 実機スクショの比較セットを作る: モック 8 種 × clone/layout モード(`DEV_TRIGGER` + `makeref`)

**受け入れ**: 改修前の数値と画像が揃っている(改修後の比較対象)。

### Phase 1 — `cell` 単一レンダラ + desugar(1〜2 日)

対象: `public/layout.js`

- [ ] `STYLE` トークンを 1 箇所に集約(既存定数はエイリアスとして残す)
- [ ] `LEAF.cell` を新設(§2.1 の属性表どおり)。`naturalHeight` の `cell` 分岐を追加
- [ ] `normNode` に desugar を追加: `label`/`card`/`bullets`/`text`/`callout`/`icon`/`chevrons` → `cell`/`grid`
- [ ] `items` のネスト配列(第 2 階層)対応。描画は「先頭に `　–` 相当のインデント付きテキスト」で決定論的に行う(Office.js の `indentLevel` には依存しない。**要検証**: `ParagraphFormat.indentLevel` が使える API バージョンなら差し替え可)
- [ ] 旧 `LEAF.label/card/bullets/text/callout/icon/chevrons` を削除(desugar 後に到達しないことをテストで確認)
- [ ] 強調バジェット(highlight ≤2、tint ≤1)を `normalizeSpec` で強制

**受け入れ**: `npm test` が無改修で通る(モック 8 種の prims が版面内・18pt 以上・塗り+枠なし・角丸なし)。Phase 0 のスクショと目視で差がない、または意図した差だけ(矢羽の下罫消失など)。

### Phase 2 — `table` 拡張(1 日)

対象: `public/layout.js`, `public/render.js`

- [ ] `headShape:"chevron"`(行見出し矢羽の重ね描き)。`render.js` は既存の `rect` prim(shape: homePlate/chevron)で描けるので変更不要
- [ ] `colGroups`(2 階層列見出し帯)
- [ ] セル値 `{text, fill, highlight}` と `axes`(4 象限)
- [ ] 行見出し無し表(第 1 列幅 0)
- [ ] `test/layout.test.js` にスイムレーン・ガント・RACI・4 象限のサンプルを追加(表 prim の行高さ合計 = 表高さ、重ね図形が行位置と一致、を検証)

**受け入れ**: 添付「プロジェクトアプローチ」相当のスライドが `table{headShape:"chevron"}` の JSON 1 つで実機に出る(ヘッダーは無地+強調色罫)。

### Phase 3 — grid 規則の追加(半日)

対象: `public/layout.js`

- [ ] 兄弟列の行揃え(§2.3-1)
- [ ] 自動リフロー(§2.3-2)
- [ ] `palette.accent === "none"` のモノトーンモード(タスクペインの「なし」スウォッチは UI 刷新計画 Phase 3 に 1 行追記して委ねる)

**受け入れ**: `cols[rows[cell{head},cell{items}]×3]` で見出しの下端が揃う。`cols[cell×6]` が 2 行に折れる。accent none で prims に強調色が 1 つも出ない。

### Phase 4 — プロンプト改訂(半日 + 計測 1 日)

対象: `server/prompt.js`, `server/mock.js`

- [ ] `schemaDoc()`: 型カタログを 4 葉 + 属性に書き換え。「構成の選び方」を §3.1 の関係→部品表に置換
- [ ] few-shot をフル 1 本 + 断片 4 本に削減
- [ ] `server/mock.js` に §3.1 の各関係を 1 つずつ含むサンプルを追加(列挙 P3 / 拡散 / スイムレーン / 階層 / 4 象限 / RACI)
- [ ] 本物の LLM で 20 プロンプト(比較・順序・因果・数値・階層・ロードマップ・議事録要約 等を混ぜる)を流し、Phase 0 のベースラインと比較:
  - 生成時間(ms)、completion_tokens
  - 構成の分布(4 形に偏っていないか)
  - `normNode` で捨てられたフィールド数(LLM が語彙を誤用していないか)
  - 実機スクショで崩れゼロ

**受け入れ**: 生成時間がベースラインと同等以下、構成分布が入力の関係性に追従、崩れゼロ。

### Phase 5 — 文書更新(半日)

- [ ] `docs/architecture/generation-pipeline.md` §2.2/§3 を To-Be に書き換え(As-Is 文書なので「実装後」に更新)
- [ ] swimlane 提案のステータスを superseded に
- [ ] `README.md` の Node 型一覧を更新、`CLAUDE.md` の Office.js 要点に「表の行見出し矢羽は重ね描き」を 1 行追記

## 7. 非スコープ(今回は作らない。近似で逃がす)

| 型 | 理由 | 近似 |
|---|---|---|
| 回転型(循環矢印) | 戻り矢印の描画は Office.js の曲線コネクタが必要で崩れやすい | フロー + 末尾 cell に「→ 計画へ戻る」 |
| 上昇/下降型(階段配置) | 座標の斜め配置は grid で表せない | 横フロー(chevron)で順序だけ示す |
| 組織図・業務フロー図(分岐) | ツリー/分岐は grid の対象外 | `table` か `rows/cols` の入れ子 |
| 見積書・免責事項 | 定型文書。`table` と `cell{items}` で足りる | そのまま |

## 8. リスクと検証項目(着手前に必読)

- **desugar の等価性**: `card` → `cell` 移行で見出し下罫・アイコン位置・highlight 地の 3 点が現行と一致すること。Phase 0 のスクショ比較でしか検知できない。
- **表の重ね図形と `insertSlidesFromBase64` 後の座標**: clone モードでは表位置が `profile.body` 基準になる。重ね矢羽は表と同じ `bodyRect` 座標系で計算するので理論上ずれないが、実機で必ず確認する(UI 刷新計画 §2-1 の geoKey の脆さと同根)。
- **ネスト items の Office.js 描画**: 箇条書き ON の状態で先頭にインデント文字を置くと、点が 1 階層目と同じ位置に出る。第 2 階層は `bulletFormat.visible=false` の別段落として描くか、`indentLevel` を使うかを Phase 1 で実機検証して決める(**未検証**)。
- **LLM の語彙移行**: `shape`/`fill` を LLM が乱用(全 cell に `fill:"tint"` 等)する可能性。Phase 4 の「捨てられたフィールド数」で監視し、バジェット(§2.5)で機械的に抑える。
- **`colGroups` の幅同期**: 表の `colWidths` は `addTable` に渡す丸め値で確定するので、群見出し帯も同じ丸め値で描く(小数のままだと 1〜2pt ずれて見える)。

## 9. 完了の定義

- `npm test` グリーン(前提)
- Phase 0 と同じ 8 モック × 2 モードの実機スクショで崩れゼロ、意図しない見た目差ゼロ
- 本物の LLM 20 プロンプトで生成時間がベースライン同等以下、構成が 4 形に偏らない
- 添付「プロジェクトアプローチ」相当が JSON 1 つ(`table{headShape:"chevron"}`)で実機に出る
- 未検証項目(§8)が残る場合は完了報告に「未検証」として列挙する
