# Flash Slide — プロンプトから高速でコンサル風スライドを作る PowerPoint アドイン

Office.js(PowerPoint JavaScript API 1.4〜1.10)と LiteLLM(OpenAI 互換)だけで、
**1 プロンプト → 1〜2 枚**のスライドを高速生成する。複雑な推論は行わず、**1 回の LLM 呼び出し**で
スライド構造(JSON)を得て、レイアウト計算はローカルの純関数エンジンで行う。

- **既定はグレースケール**(色はタスクペインで選んだときだけ。「自動」で開いているスライドの強調色を採用、プリセット 5 色、カスタム)
- **自律的なレイアウト**: 固定テンプレートではなく、内容に応じて LLM がブロック(グリッド)を構成
- **ノンデザイナーズ原則**: 余白・整列・Z の視線誘導・情報の階層・本文 18pt 以上
- **テンプレート追従**: 開いているスライドのデザイン(タイトル位置・フォント・色・フッター/ページ番号)を
  解析し、それを引き継いで新規 1 枚を挿入する(コピペ運用デッキでも前後で揃う)
- **大量入力**: 1〜2 枚に収まらないときは章構成を提案し、ユーザー調整後に各章 1 枚ずつ並列生成

## セットアップ

```bash
npm install
npx office-addin-dev-certs install   # 初回のみ: https://localhost 用の開発証明書
npm run icons                        # マニフェスト用アイコン生成(public/assets/)
```

### .env(秘密情報の唯一の置き場所)

パス: リポジトリ直下の **`.env`**(`.env.example` をコピー)

```
LLM_BASE_URL="https://<litellm-host>"   # 末尾 /v1 は付けても付けなくてもよい
LLM_API_KEY="<key>"
LLM_MODEL="<model-name>"                # 既定モデル(タスクペインで上書き可)
PORT=3455                               # manifest.xml の URL と一致させること
```

- LLM_BASE_URL / LLM_API_KEY が未設定なら**モックモード**(サンプルで描画)で起動する。
- GET が WAF で拒否されるゲートウェイに対応(モデル一覧は GET→POST フォールバック、生成は POST のみ)。
- `/chat/completions` が 404 の場合は `/v1/chat/completions` を自動再試行。
- `response_format`/`max_tokens` 非対応モデルには自動でフォールバック。重い推論モデルにもタイムアウトを設けない。

## モデル選定(実測)

`test/cases.json` の 13 ケース × 2 回を `scripts/bench.js --models … --repeat 2` で流し、
速度(サーバ計測)と品質(レイアウト警告ゼロ・縮小ゼロ・期待構造の一致)で比較した。
モデル名は環境依存なので、ここでは規模で表記する(`LLM_MODEL` と `--models` には実際の名前を渡す)。

| モデル | 平均 | 合格 | 構造一致 | 備考 |
|---|---|---|---|---|
| **軽量モデル(既定)** | 2.3s | 25/26 | **26/26** | 不合格 1 件はテンプレ無し時のガント行数警告 |
| 中位モデル | 5.4s | 22/26 | 24/26 | 2 ケースで構成の取り違え |
| 上位モデル | 6.6s | 20/26 | 24/26 | 1 ケースで縮小、1 ケースで構成の取り違え |

生成文法を「パネル数 + 主図 1 つ + 補足 1 つ」に絞った結果、**軽量モデルでも構造一致が 26/26** になり、
上位モデルへ替える理由が無くなった。旧文法(自由な入れ子)では軽量モデルの構造一致は 15/20 で、
上位モデルが 10/10 と明確に勝っていた。**品質はモデルではなく文法とレイアウトエンジンで決まる**というのが本プロジェクトの結論。

## 実行

```bash
npm start        # HTTPS サーバ(https://localhost:3455)
npm run sideload # PowerPoint を起動しアドインを読み込む(別ターミナル)
npm run unload   # サイドロード解除
npm test         # 単体+統合テスト(Office.js 不要)
npm run validate # マニフェスト検証
```

実運用の使い方: **対象デッキを開き**、ホームタブの「スライド生成」でタスクペインを開き、
プロンプトを入力して生成。選択中スライドの直後に、そのデッキのデザインを引き継いで挿入される。

## 構成

| ファイル | 役割 |
|---|---|
| `server.js` | HTTPS + API(`/api/generate` `/api/health` `/api/models`) |
| `server/llm.js` | LiteLLM 呼び出し(フォールバック・モデル一覧) |
| `server/prompt.js` | システムプロンプト(スキーマ・デザイン規律)と応答の抽出・整形 |
| `server/mock.js` | LLM 未設定時のサンプル spec(実務的な密度) |
| `public/layout.js` | **spec → 描画プリミティブ**の純関数エンジン(Node で単体テスト可)。cell 単一レンダラ・セル合成の格子(table)・figure・グリッド配置・自動フォント縮小・強調バジェット |
| `public/render.js` | Office.js 描画。参照スライド解析(`analyzeReference`)、clone(複製+本文差替)/ layout(既定)の二経路 |
| `public/taskpane.{html,css,js}` | タスクペイン UI(入力 1 欄・ボタン内進捗・下部の地味ゾーン: モデル/レイアウト/色スウォッチ/読み取り状態・章構成・開発者向け JSON 再描画) |
| `manifest.xml` | アドインマニフェスト(PowerPoint / ReadWriteDocument) |
| `test/` | `layout` `reference` `integration` テスト、`cases.json`(実プロンプト 10 件)、`fixtures/`(3 種デザインのデッキ生成) |
| `scripts/` | `bench.js`(モデル選定ベンチ)、`run-triggers.js`(実機での連続生成)、`uia-click.ps1`(前面化なしでリボン操作)、`ui.ps1`(スクショ)、`ppt-*.ps1`(COM でデッキ検査) |

## スライド仕様(spec)の形(v4: 部品合成)

`body` は入れ子グリッド。部品は **4 種だけ**(`cell` / `table` / figure / `arrow`)で、見た目は属性で決まる。
「型」を増やさず、要素間の関係(列挙・因果・順序・比較・数値・階層)に合わせて部品を組み合わせる。

```jsonc
{ "title": "結論(メッセージ)", "lead": "根拠 1 文", "footnote": "出典",
  "body": { "cols": [ {"type":"cell","head":"課題","icon":"warning","items":["..."]},
                      {"type":"arrow"},
                      {"type":"cell","head":"打ち手","highlight":true,"items":["..."]} ],
            "weights": [5,0.4,5] } }
```

| 部品 | 属性 | 表せるもの |
|---|---|---|
| `cell` | `head` / `text` \| `items`(ネスト配列で第 2 階層)/ `shape:"chevron"` / `fill: light\|dark\|tint` / `highlight` / `icon` / `align` | 見出し・カード・箇条書き・地の文・示唆の強調地・矢羽・ピクトグラム(旧 label/card/bullets/text/callout/chevrons/icon はすべてこれに脱糖) |
| `table` | `corner` / `colHeaders` / `rows[{head,cells}]` / `highlightCol` / `headShape:"chevron"` / `colGroups` / `axes` / セル `{text,fill,highlight}` | **セルを組み合わせた格子**(ネイティブ表ではない)。比較・評価・RACI・スイムレーン(行見出し矢羽)・ガント(塗りセル)・4 象限(軸ラベル)・2 階層列見出し |
| figure | `bars` / `column` / `line` / `stacked` / `kpi` / `ntable`(PPT ネイティブ表、密な数表向け) | 数値 |
| `arrow` | `direction` | 前提→帰結の塗り三角 |

エンジン(`public/layout.js`)がコードで強制する規律: 強調(highlight)は 1 枚 2 箇所まで、tint 塗りは 1 つまで、
塗りと枠線の併用禁止、角丸なし、2 パネル等幅、葉 5 個以上の cols は 2 行に自動リフロー、兄弟列の見出し行揃え、
格子は 14pt まで縮めてから行高さを比例縮小(領域外に出さない)。`palette.accent:"none"` で完全モノトーン。
`**太字**` で強調。旧形式(`type: card|label|bullets|callout|chevrons|boxes|columns|matrix|process`)も自動変換。

## デザイン規律(consulting-pptx-skill の slide-rules を単スライド向けに実装)

- 角丸禁止。塗りのあるボックスに枠線を付けない。罫線は区切る対象がある場所だけ(最終行の下に引かない)。
- 表(格子)はセルの組み合わせで描く。列見出しは濃い帯 + 白文字(`headFill:"none"` で罫だけに)、行名は薄い箱(`rowHeadFill:"none"` で抑止)。行間は薄罫、縦罫なし、ゼブラなし、非該当は「—」。ネイティブ表は `ntable`(figure 扱い)。
- 左=事実・分析、右=意味合い(So what)の 2 カラム。下部の結論帯・浮遊ボックスを作らない。
- 前提→帰結の強い因果は列間に塗り三角 1 つ。矢羽は細い全幅の帯、説明は帯の下にカラム(縦罫で区切る)。
- 本文は上端揃え(中央に浮かせない)。文字だけの行・箱は自然高さの 1.5 倍までしか伸ばさず、格子・図・ガント・体制図が領域を使い切る。
- タイトルは常体の結論文(「〜です/ます」「ラベル：」前置きは正規化で除去)。
- 本文はテンプレの本文サイズ(既定 14pt)でスライド全体を揃える(部品ごとの拡大なし)。版面充填率 70% 未満なら全体を +2pt ずつ(20pt まで)拡大、溢れ・縮小が出たら全体を 1pt ずつ縮小してから、それでも溢れる部品だけ段階縮小(警告)。
- 強調は太字が基本。`fill:"dark"` は見出し帯の反転(濃い帯 + 白文字)として描き、箱全体を黒く塗らない。箱フローの一部だけを塗らない。示唆帯(takeaway)は左バー無しの薄い地。lead の内容を本文で繰り返さない。
- 構成の方法論(プロンプト + `applyMethodology`): 並列の事象が 2〜3 で短ければ等幅の下線付きボックス、3〜6 で長ければ行名の箱 + 横長の説明の格子(縦に並べる)、共通の観点が 2 つ以上なら列見出し帯付きの格子、時間軸はガント(ホームベース型の帯)、組織は体制図。
- 数値がある入力は表の羅列より Figure(推移=折れ線、構成比=積み上げ、量の差=棒)を優先する。

## 生成の流れ(タスクペイン)
1. 「内容」に伝えたい中身(メモ・議事録・AI の回答)をそのまま貼る。文字数の横に枚数の目安が出る。
2. 「スライド生成」→ **4 案を並列生成**し、最初に返った案を開いているデッキの選択スライドの直後に挿入(体感 3〜6 秒)。
3. 入力欄は小さく畳まれ、「生成資料」に **4 案のサムネが 2×2 の格子**で並ぶ(挿入済みは太枠)。押すと差し替え。
4. 「枚数」(自動(1〜2枚) / 1枚 / 2枚)を変えると同じ入力で再生成し差し替え。入力を書き換えれば「スライド生成」に戻り、追加になる。
5. レイアウトの希望を文章で書きたいときだけ「仕上げの希望」を開く(色・レイアウト・モデルもここ)。

## 開発メモ / 既知の挙動

- **アイコン/ピクトグラム**は [Tabler Icons](https://tabler.io/icons)(MIT)の SVG を `@tabler/icons` から読み、
  サーバ(`/api/icons`)で色付き PNG に描画(`@resvg/resvg-js`)。タスクペインは Office 共通 API の
  `setSelectedDataAsync(image)` で位置指定して貼る。LLM は名前を選ぶだけ(`/api/icons` で一覧)。
  ※ フォントグリフ方式は PUA 文字が CJK ランで豆腐(□)化するため不採用。
- **table ノードは自身の `rows` を持つ**ため、コンテナ判定は `type` が無い場合に限る(過去に表が描かれない原因)。
- **直前に生成したスライドを参照にすると clone モードになる**(前後の統一)。本文領域内の細線・画像・注記は
  保持せず、ヘッダ帯・フッタ帯・余白帯の装飾と、出典(最下段の小文字)だけを引き継ぐ。
- **参照スライド解析**は表紙(CenterTitle)や下寄せタイトルを既定レイアウトに退避。
  タイトル/リード/フッター/出典/罫線/ロゴを判定し、装飾は保持して本文だけ差し替える。
- テストデッキ: `python test/fixtures/make-decks.py` で A(プレースホルダ)/B(コピペ運用)/C(4:3 ダークバンド)を生成。
- **ポート競合注意**: 別プロジェクト(PPTADDIN_CODEX)が 3000/3443 を使うため本プロジェクトは **3455**。
- 開発検証用に `DEV_TRIGGER=1` で起動するとタスクペインがサーバのトリガをポーリングする
  (`POST /api/debug/trigger`、ペイン側は `?v=4` 付きで取得。旧ペインが横取りしない)。`DEBUG_SNAPSHOT=1` で生成 PNG を `debug/` に保存。
  `DEV_MODE=1`(または `DEV_TRIGGER=1`、URL `?dev=1`)でタスクペインに「スライド仕様 JSON」編集欄が出る。いずれも本番は無効。
- 実機検証の手順: `node scripts/run-triggers.js reload reanalyze swimlane insight …`(モック連続描画)/
  `node scripts/run-triggers.js --file test/cases.json --model <model>`(本物の LLM)。トリガ `reload` でペインを再読み込み、`reanalyze` で参照解析だけ実行。
- モデル選定: `node scripts/bench.js --models a,b,c`(サーバ経由で実 LLM を叩き、速度・トークン・構造一致・レイアウト警告を表にする)。
  `--rejudge debug/bench-*.json` で保存済み spec をエンジン現行版で再判定。
- タスクペインを前面化せずに開く/撮る: `pwsh scripts/uia-click.ps1 "<window>" "ホーム" "スライド生成"`、`pwsh scripts/ui.ps1 shotwin "<window>" out.png`。
- Office は `taskpane.html` を強くキャッシュする。`manifest.xml` の URL に `?v=N` を付けて再サイドロードするか、ペインをリボンから開き直す。
