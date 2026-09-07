# 更新: 新規生成の構成規則

2026-09-06 の追加修正により、新規 LLM 応答は [パネル数と主図を先に決める方式](panel-grammar.md) に移行した。以下の自由なグリッド合成・強調属性の記述は旧保存 JSON／モックの互換経路を説明している。現在の新規生成では 1/2/3 パネルを先に選び、主図の混在と任意の塗り強調を禁止する。

---

# スライド生成パイプライン 技術仕様書(As-Is / v4)

- 作成日: 2026-09-06 / 最終更新: 2026-09-06(部品合成エンジン v4 に更新)
- ステータス: **現状(As-Is)の技術仕様。実装済みの挙動の記述であり、変更提案ではない。**
- 対象: `server.js` / `server/prompt.js` / `server/llm.js` / `public/layout.js` / `public/render.js` / `public/taskpane.js`
- 設計の背景: [../feature-proposals/2026-09-06-composable-cell-grid-engine-plan.md](../feature-proposals/2026-09-06-composable-cell-grid-engine-plan.md)(なぜ「型」を捨てて部品合成にしたか)

---

## 1. パイプライン全体像

```
taskpane.js
  │ prompt, hint, model, mock        (素材は prompt に貼るだけ。context 欄は廃止)
  ▼
POST /api/generate (server.js)
  │ buildMessages() → system(schemaDoc: 部品 4 種 + 関係→部品の対応表 + 規律 + 例) + user(指示/希望/枚数)
  ▼
LLM (LiteLLM 経由。推論は 1 回・jsonMode。response_format / max_tokens / temperature 非対応モデルは該当パラメータを外して再試行)
  │ raw JSON text
  ▼
extractJson() → shapeResponse() → SlideLayout.normalizeSpec()   (server/prompt.js + layout.js)
  │ { slides: [spec, …] }  ※入力過多なら { sections: […] } を返し中断
  │   normalizeSpec = 脱糖(旧型→cell/grid)+ ゆらぎ吸収({table:{…}} 包み等)+ 強調バジェット(highlight 2 / tint 1)
  ▼
taskpane.js: SlideRender.renderSpecs(specs, {prepared, palette, layoutId})
  │   prepared = 起動時/選択変更時に済ませた参照解析(解析中は生成ボタンを無効化)
  ├─ SlideLayout.layout(spec, {width, height, palette, profile})   … Office.js 非依存の純関数
  │     spec(意味構造) → prims[](座標付き rect/line/table/image 命令) + header(差し替え用テキスト)
  └─ render.js: PowerPoint.run(...) で prims を Office.js シェイプへ変換・挿入
        └─ 画像(icon)は /api/icons で事前PNG化 → setSelectedDataAsync で貼付
```

骨格は不変: 「LLM 推論は 1 回」「座標・色・pt は LLM に出させず純関数で決める」「Office.js 呼び出しは最小」。
v4 で変わったのは **語彙(14 型 → 4 部品 + 属性)** と **規律の置き場所(プロンプトのお願い → エンジンのコード)**。

---

## 2. LLM 入出力契約(`server/prompt.js`)

### 2.1 リクエスト(`buildMessages`)

- `system`: ロール設定 + `schemaDoc()`。内容は (a) 出力形式 (b) 部品 4 種の属性表 (c) **要素間の関係 → 部品の組み方**の対応表(列挙 / 背景・合流 / 拡散 / 順序 / 比較 / 段階×観点 / 役割 / 2 軸 / 数値 / 前提→帰結 / 階層 / 分析→解説。「型ではなく組み合わせ例」と明記)(d) デザイン規律 9 条 (e) フル例 1 本 + body 断片例 4 本。
- `user`: `【指示】`(素材ごと貼られたプロンプト)→ `【レイアウトの希望】`(任意)→ `【枚数】`。
- 呼び出し: `llm.chat({ messages, maxTokens: 3000, temperature: 0.2, jsonMode: true })`。**推論はこの 1 回のみ**。

### 2.2 部品(`body` の語彙)

| 部品 | 主な属性 | 備考 |
|---|---|---|
| `cell` | `head` / `text` \| `items`(ネスト配列で第 2 階層)/ `shape:"chevron"` / `fill: light\|dark\|tint` / `highlight` / `icon` / `align` / `size` / `level` | テキストを持つ唯一の部品。旧 `label`/`card`/`bullets`/`text`/`callout`/`icon`/`chevrons` は `normNode` でこれに脱糖 |
| `table` | `corner` / `colHeaders` / `rows[{head, cells, highlight}]` / `highlightCol` / `headShape:"chevron"` / `colGroups[{text,span}]` / `axes{x,y}` / セル `{text, fill, highlight}` | **セル合成の格子**。比較・評価・RACI・スイムレーン・ガント・4 象限・2 階層列見出し |
| figure | `bars` / `column` / `line` / `stacked` / `kpi` / `ntable`(PPT ネイティブ表) | 数値。`ntable` は 8 行以上の密な数表向け |
| `arrow` | `direction` | 前提→帰結の塗り三角(cols の間) |
| 容器 | `rows` / `cols` + `weights` + `gap` | 3 段まで |

### 2.3 応答の抽出・分岐(`extractJson` / `shapeResponse`)

変更なし: コードフェンス除去 → `JSON.parse` → `{…}` 切り出し → 末尾カンマ除去。`sections` が返れば章構成確認 UI へ。

---

## 3. レイアウトエンジン(`public/layout.js` v4)

### 3.1 デザイントークン(`STYLE`)と参照テンプレートへの追従

参照スライドがあるときは、その **本文の文字サイズ(最頻値)を基準に文字階層を組み替える**(`applyFontProfile`)。
本文=テンプレ値、`min`=同値(これを下回ったときだけ縮小扱い)、`floor`=×0.72、`large`=×1.25、格子=×0.85、見出し=テンプレの太字最頻値、KPI=×2.6、出典=テンプレの出典サイズ。
テンプレートが無い(既定レイアウト)ときだけ以下の既定値を使う。`layout()` の最後で既定に戻すので呼び出し間で引きずらない。

`FONT`(title 26 / head 20 / body 18 / floor 12 / tableMin 12 / large 22 / kpi 48 / footnote 11)、`SPACE`(margin 36 / pad 8 / gap 20 / cellPad 6 / chevronOverlap 10)、`RULE`(thick 1.8 / thin 0.75)、`BUDGET`(highlight 2 / tint 1 / leaves 12 / depth 3)。全描画はここだけを参照する。`palette.accent === "none"` で完全モノトーン(強調は太字 + 濃い地)。

### 3.2 `normalizeSpec` → `normNode`(脱糖・防御・バジェット)

- エイリアス吸収(`headline`/`message`/`source`、`children`/`blocks`/`columns` …)、`cleanTitle`(常体化・ラベル前置き除去)。
- **脱糖**: `label→cell{head}`、`card→cell{head,text|items}`、`bullets→cell{items}`、`callout→cell{fill:tint}`、`icon→cell{icon}`、`chevrons→rows[cols[cell{shape}…], cols[cell{items}…]]`(6 段以上は `table{headShape:"chevron"}`)、`boxes/columns→cols[cell…]`、`matrix→table`、`figure{kind}→各 figure`。
- **ゆらぎ吸収**: `{"table":{…}}` のように型名をキーにした包み方を展開(`WRAP_TYPES`)。
- **正規化規則**: 2 パネル(+arrow)は左右等幅強制、葉 5 個以上の `cols` は 2 行に自動リフロー、葉 12 個・深さ 3 で打ち切り。
- **強調バジェット**(`enforceBudget`): `highlight` は文書順で 2 個まで(cell / table の行・列・セル / bars・column の項目を合算)、`fill:"tint"` は 1 個まで。kpi と単系列の折れ線は数えない。

### 3.3 `layout(spec, options)`

1. パレット確定(`palette.accent` > `profile.accent` > 既定 `#FD5108`)。
2. タイトル: `profile.title`(clone 時)があれば `fitSlot`(1 行で入るなら元サイズ、入らなければ行数×1.15 で枠に収まるまで縮小、下限 18)、無ければ既定位置 + 下線。
3. リード: `profile.lead` があり `lead` があるときだけ。出典・本文領域は従来どおり。
4. `body` を `renderNode` で再帰配分。

### 3.4 グリッド(`renderRows` / `renderCols`)

- 自然高さ `naturalHeight` と固定/伸縮フラグから `distributeHeights` で配分(収まらなければ固定行据え置き → 伸縮行に残り → 全体比例縮小。領域外には出さない)。
- **兄弟列の行揃え**: `cols` の全列が同じ段数の `rows` なら、段ごとの高さを列間で共有(見出しの下端・本文の上端が揃う)。
- 版面充填: 自然高さが 55% 未満なら垂直中央。

### 3.5 `cell` レンダラ(唯一のテキスト部品)

- `shape` あり: 矢羽/五角形の帯(既定 `fillDark`、highlight で accent、白文字・中央)。
- `head`: 太字 20pt(入らなければ 18pt×2 行)+ 下に太い罫(highlight で accent)。icon は見出し左。
- 本文: `fitBody`(18pt 基本、14pt まで縮小・22pt まで拡大)。`items` 2 件以上でネイティブ箇条書き、ネスト配列は「•/　–」の記号付き行(ネイティブ箇条書きは使わない)。数値だけのセルは右揃え。
- `fill`: light / dark(白文字)/ tint(強調色の 90% 薄地)/ accent。塗りのある箱に枠線は付けない。

### 3.6 `table` レンダラ(セル合成の格子)

- `matrixMetrics`: 行見出し幅(26%、矢羽なら 20%)、列幅均等、列グループ帯、列見出し帯、行高さ。フォントは 18pt から、**14 字以下の短いセルが折り返すなら 14pt まで縮小**、それでも高さが足りなければ **行高さを比例縮小して領域内に収める**(文字は PowerPoint の自動縮小に委ねる)。
- 描画: 列グループ(テキスト + 群幅の細罫)→ 列見出し(無地・太字・中央、下に全幅の強調色太罫)→ 行(行見出し太字 / `headShape:"chevron"` なら 90° 回転した homePlate/chevron 図形 + 別テキストボックス)→ セル(塗り・強調地・数値右揃え・複数行は記号の有無で箇条書き判定)→ 行間の薄罫(最終行なし)→ `axes` の軸ラベル(y は 270° 回転)と列間の縦罫。
- `—` は明示的なダッシュか、塗りセルの無い表の空セルだけ(ガントの空白には出さない)。

### 3.7 figure / `ntable`

`bars` / `column` / `line` / `stacked` / `kpi` は v3 と同じネイティブ図形描画。`ntable` は旧 `LEAF.table`(`addTable` 1 コール、フォールバックは矩形)で、行高さが領域を超えるときは比例縮小。

### 3.8 テキスト計測

`estimateLines` は文字幅の概算(CJK 1em / 大文字数字 0.62 / 小文字 0.5)に **幅 5% の安全率**を掛ける(実フォントの折り返しは概算より 1 行早く起きる)。

---

## 4. Office.js 描画層(`public/render.js`)

### 4.1 `prepare()`

1 回の `PowerPoint.run` で、サイズ・選択スライド・マスタ/レイアウト一覧・選択スライドの全シェイプ詳細を取得。加えて **レイアウトのプレースホルダ既定サイズ**(Title / Subtitle 等の `font.size`)を読み、`refShapes[].layoutFontSize` に付ける(縮小したタイトルを持つ生成物を参照にしても、次の複製で元のサイズに戻る)。

### 4.2 `analyzeReference()`

- **自前シェイプ(名前 `FS_*`)を解析対象から除外**する(直前の生成物を参照にしても、その本文を装飾として保持しない)。
- タイトル: Title プレースホルダ > 上部の幅広テキスト。表紙判定は CenterTitle / タイトルが下寄せ / 「Title + Subtitle だけで本文が無い」。**Subtitle プレースホルダはリード文として扱う**(実務テンプレはこの形が多い)。
- `paragraphFormat.horizontalAlignment` は文字列以外が返ることがあるため `alignOf` で防御。
- 保持対象(ヘッダ帯・フッタ帯・余白帯の装飾、出典行)、本文領域、強調色(最頻の有彩色)、フォント名は v3 と同じ。

### 4.3 `renderSpecs(specs, opts)`

- `opts.prepared` を使う(無ければ `prepare()`)。`opts.layoutId` が現在のマスタに存在すれば `mode:"layout"` の `slides.add` で自動選択より優先。
- clone: 複製 → `geoKey` でタイトル/リード/出典を特定し文字だけ差し替え。**タイトル/リードのフォントサイズは常に計算値を設定**(`fitSlot` の結果。縮小の連鎖を防ぐ)。
- prims → `applyRect`(回転・塗り・枠・テキスト・太字レンジ)/ `applyLine` / `applyTable`(ntable)/ 画像。

### 4.4 タスクペイン(`public/taskpane.js`)

- 起動時と **選択スライド変更時**(`DocumentSelectionChanged`、400ms デバウンス)に `refreshReference()`。再入ガード(シーケンス番号)で最後の呼び出しだけを反映。解析中は生成ボタン無効、完了(成功/失敗)で有効。
- 進捗はボタン内(「生成中… 2.3s」→「描画中…」→「再生成」)。プレビュー画像は表示しない(PNG は `DEBUG_SNAPSHOT` 用に取得は継続)。
- 地味ゾーン: モデル / レイアウト(clone 時は「自動(テンプレ追従)」で無効)/ 色スウォッチ(自動・5 色・なし・カスタム。単一の状態変数 `selectedAccent`)/ 読み取り状態アイコン(analyzing / clone / layout / error、ツールチップで詳細、クリックで再読み取り)。
- 開発者モード(`?dev=1` or `/api/health.devMode`)で「スライド仕様 JSON」編集欄。`DEV_TRIGGER` ポーリングは `?v=4` 付き(`reload` / `reanalyze` / 生成トリガ)。

---

## 5. なぜ速いか(v4 での実測を含む)

1. LLM 呼び出し 1 回・jsonMode。語彙を 4 部品に絞ったので出力 JSON が安定し、再試行が発生しない。
2. 座標・フォント・色は純関数(ミリ秒)。品質は検証ループではなく **構成的制約**(バジェット・自動リフロー・行揃え・領域内クランプ)で担保。
3. 参照解析は起動時・選択変更時に済ませておく(生成時は待たない)。
4. Office.js は `sync` 3〜4 回/枚。格子はセル図形の一括追加、ntable は `addTable` 1 コール。

実測(2026-09-06、`test/cases.json` の実プロンプト 10 件、`scripts/bench.js`): モデル選定の結果は README の「モデル」節と `debug/bench-*.json` を参照。

---

## 6. 添付スライド例の扱い(v4 での結論)

- 横棒グラフの報告スライド → `bars` + `kpi` + `cell{items}`(モック `report`)。
- 「プロジェクトアプローチ」(縦矢羽 × 3 列) → **新 Node 型は追加せず** `table{headShape:"chevron"}`(モック `swimlane`)。列ヘッダーは無地 + 強調色罫の規律を維持し、矢羽は濃い地 + 白文字(強調段は accent)。`debug/approach-test.pptx`(同テンプレに add-in を埋め込んだ検証デッキ)でテンプレ追従(Georgia タイトル・Subtitle のリード・会社フッタ・ページ番号・#D04A02)を実機確認済み。

---

## 7. 参考: 関連ファイル一覧

| ファイル | 役割 |
|---|---|
| `server.js` | `/api/generate`(所要時間・トークン数をログ)、`/api/health`(devMode)、`/api/debug/*` |
| `server/prompt.js` | システムプロンプト(部品 4 種・関係→部品表・規律・例)と応答の抽出 |
| `server/llm.js` | LiteLLM 呼び出し(/v1 フォールバック、response_format / max_tokens / temperature 非対応の再試行) |
| `server/mock.js` | モック 15 種(flow / insight / pyramid / roadmap / report / compare / trend / share / swimlane / enumerate / diffuse / hierarchy / quadrant / raci / gantt) |
| `public/layout.js` | v4 エンジン(STYLE・cell・table・figure・grid・normalizeSpec・enforceBudget) |
| `public/render.js` | Office.js 描画、参照解析(FS_ 除外・レイアウト既定サイズ)、clone / layout |
| `public/taskpane.{html,css,js}` | タスクペイン UI(刷新版) |
| `test/layout.test.js` | 品質ゲート(全モック × 3 サイズ × プロファイル、v4 規則) |
| `test/reference.test.js` | 参照解析(A/B/C + 表紙/空 + D: Subtitle をリードに使う実務テンプレ) |
| `test/cases.json` | 実プロンプト 10 件(ベンチ・実機検証用) |
| `scripts/bench.js` / `scripts/run-triggers.js` / `scripts/raw.js` | モデル選定ベンチ / 実機連続生成 / 生応答の確認 |
| `scripts/uia-click.ps1` / `scripts/ui.ps1` / `scripts/ppt-*.ps1` | 前面化なしのリボン操作 / スクショ / COM 検査 |
