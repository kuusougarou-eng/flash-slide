"use strict";
/**
 * LLM 未設定時 / mock 指定時に返すサンプル spec。
 * 語彙は v4(cell / table / figure / arrow + rows/cols)。旧型(card/label/bullets/callout/chevrons)も normalizeSpec が吸収する。
 * slide-rules 準拠: 左=事実/右=意味合い、軸のある格子、下部帯なし、常体タイトル、実務的な密度。
 * 各サンプルは「要素間の関係 → 部品の組み方」の 1 例(型ではない)。
 */
const SAMPLES = {
  // 前提 → 帰結(格子 + 三角 + 示唆)
  flow: {
    title: "一次回答の自動化で問合せ対応の遅延を 6 割解消できる",
    lead: "遅延の 62% は定型質問への一次回答待ちで発生している。",
    footnote: "出典: 2026年7月 問合せログ分析(n=1,240)",
    body: {
      cols: [
        {
          rows: [
            { type: "cell", head: "遅延の内訳と原因" },
            {
              type: "table",
              corner: "要因",
              colHeaders: ["構成比", "背景"],
              rows: [
                { head: "一次回答待ち", cells: ["62%", "定型質問に担当者が手作業で回答"], highlight: true },
                { head: "担当者依存", cells: ["21%", "回答品質と速度が個人で異なる"] },
                { head: "夜間・休日", cells: ["17%", "翌営業日まで無応答"] },
              ],
            },
          ],
        },
        { type: "arrow" },
        {
          rows: [
            { type: "cell", head: "だから、一次回答を自動化する" },
            { type: "cell", items: ["**FAQ ボット**が定型質問に即時回答する", "**回答テンプレート 40 件**を整備し品質を揃える", "**エスカレーション基準**を明文化し非定型を人が担う"] },
            { type: "cell", text: "平均回答時間を **8 時間 → 3 時間** に短縮し、担当者は非定型案件に集中する", fill: "tint" },
          ],
        },
      ],
      weights: [5, 0.4, 4],
    },
  },
  // 比較・評価(格子)→ 解説
  insight: {
    title: "3 案のうち B 案だけが市場性と実現性を両立する",
    lead: "",
    footnote: "評価: ◎ 高 / ○ 中 / △ 低",
    body: {
      cols: [
        {
          type: "table",
          corner: "評価軸",
          colHeaders: ["A 案", "B 案", "C 案"],
          highlightCol: 1,
          rows: [
            { head: "市場規模", cells: ["○ 300 億円", "◎ 1,200 億円", "△ 80 億円"] },
            { head: "自社資産の活用", cells: ["△ 新規開発が必要", "◎ 既存販路を活用", "○ 一部転用可"] },
            { head: "競合状況", cells: ["△ 大手 3 社が先行", "○ 中堅が分散", "◎ 競合なし"] },
            { head: "投資規模", cells: ["○ 5 億円", "○ 8 億円", "◎ 1 億円"] },
          ],
        },
        {
          rows: [
            { type: "cell", head: "意味合い" },
            { type: "cell", items: ["B 案は**両軸で高評価**となる唯一の案", "C 案は競合不在だが市場が小さく、投資回収が難しい", "A 案は市場は中位だが**新規開発**の負荷が大きい", "**B 案を優先**し、初年度は既存販路 2 地域で検証する"] },
          ],
        },
      ],
      weights: [3, 2],
    },
  },
  // 列挙(関係のない並列要素。cell ×3 に icon)
  pyramid: {
    title: "生成AI 導入は業務を絞って小さく始めた企業ほど定着している",
    lead: "先行導入 12 社のうち、2〜3 業務に絞って 3 か月で ROI を実証した企業は全社で定着した。",
    footnote: "出典: 先行導入企業 12 社ヒアリング(2026年8月)",
    body: {
      cols: [
        { type: "cell", head: "業務選定", icon: "filter", text: "定型・高頻度・テキスト中心の業務を優先する。先行企業の 9 割が問合せ対応か文書作成から開始した" },
        { type: "cell", head: "評価指標", icon: "chart-bar", highlight: true, text: "処理時間と再作業率の 2 指標に絞る。指標が 4 つ以上の企業は判断が遅れ、定着率が半減した" },
        { type: "cell", head: "推進体制", icon: "users", text: "現場リーダー 1 名と IT 1 名の小チームが週次で改善サイクルを回す" },
      ],
    },
  },
  // 順序(矢羽の帯 + 各段の要点)
  roadmap: {
    title: "データ基盤は参照系から 4 段階で移行し、業務影響を最小化する",
    lead: "各段階で 2 週間の並行稼働期間を設け、稼働率と再処理件数で切替を判断する。",
    footnote: "",
    body: {
      rows: [
        {
          type: "chevrons",
          steps: [
            { head: "現状把握", items: ["1,200 テーブルを棚卸しする", "依存関係を可視化する"] },
            { head: "参照系移行", highlight: true, items: ["BI ダッシュボード 38 本を移す", "2 週間並行稼働する"] },
            { head: "更新系移行", items: ["基幹連携バッチ 64 本を移す", "切替リハーサルを 2 回行う"] },
            { head: "旧環境停止", items: ["年 1,800 万円のライセンスを解約する", "運用手順を更新する"] },
          ],
        },
        { type: "cell", text: "**切替判断基準**: 各段階の並行稼働期間中に、稼働率 99.5% 以上かつ再処理 週 3 件以下を 2 週連続で満たしたときに切り替える" },
      ],
      weights: [4, 1],
    },
  },
  // 数値(横棒 + KPI)→ 示唆
  report: {
    title: "中小企業の DX は人材不足と費用対効果の不透明さが二大障壁になっている",
    lead: "調査対象 2,400 社のうち DX に着手済みは 38%。未着手企業の 7 割が人材不足を理由に挙げる。",
    footnote: "出典: 中小企業 DX 実態調査(2026年6月、n=2,400、複数回答)",
    body: {
      cols: [
        {
          rows: [
            { type: "cell", head: "未着手企業の DX 阻害要因" },
            {
              type: "bars",
              unit: "%",
              items: [
                { label: "DX 人材の不足", value: 71, highlight: true },
                { label: "費用対効果が不透明", value: 62, highlight: true },
                { label: "既存システムとの連携", value: 44 },
                { label: "経営層の理解不足", value: 31 },
                { label: "セキュリティへの不安", value: 27 },
              ],
              note: "n=1,488(未着手企業)",
            },
          ],
        },
        {
          rows: [
            { type: "cell", head: "示唆" },
            { type: "kpi", value: "38%", label: "DX 着手済み企業の割合" },
            { type: "cell", items: ["**人材**: 外部人材の共同活用と伴走支援を広げる", "**費用対効果**: 業種別の効果事例と KPI 例を示す", "**連携**: 既存 SaaS との標準 API を整える"] },
          ],
        },
      ],
      weights: [2.7, 2.3],
    },
  },
  // 対比(cell ×3)
  compare: {
    title: "内製と外注は変化頻度で使い分けるとコストと速度を両立できる",
    lead: "",
    footnote: "",
    body: {
      cols: [
        { type: "cell", head: "内製する領域(変化が速い)", highlight: true, icon: "code", items: ["**顧客接点**: 週次で UI/UX を改善する", "**データ活用**: 分析知見を社内に蓄積する", "**採用**: エンジニアを 1 年で 5 名増員する"] },
        { type: "cell", head: "外注する領域(安定している)", icon: "building", items: ["**基盤運用**: SLA 99.9% で品質を担保する", "**レガシー保守**: 変更頻度が低く固定費化する", "**コスト**: 固定費を変動費へ転換する"] },
        { type: "cell", head: "共通ルール", icon: "checklist", items: ["**設計原則**: API で疎結合にし、契約は 1 年更新にする", "**ガバナンス**: 月次レビューで境界を見直す", "**ナレッジ**: 設計書と運用手順を社内に保持する"] },
      ],
    },
  },
  // 数値(折れ線)→ 示唆
  trend: {
    title: "問合せ件数は 3 年で 1.8 倍に増え、人員は横ばいのままである",
    lead: "",
    footnote: "出典: 社内問合せ管理システム(各年 4 月)",
    body: {
      cols: [
        { rows: [{ type: "cell", head: "月間問合せ件数と対応人員の推移" }, { type: "line", unit: "件", labels: ["2023", "2024", "2025", "2026"], series: [{ label: "問合せ件数", values: [1200, 1500, 1900, 2160], highlight: true }, { label: "対応人員(人×100)", values: [900, 900, 1000, 1000] }] }] },
        { rows: [{ type: "cell", head: "読み取れること" }, { type: "cell", items: ["件数は**年 20% 超**で増加し、人員は 9〜10 名で変化なし", "1 人あたり月 133 件 → **216 件**に増え、残業で吸収している", "増加分の 6 割は定型質問で、**自動化の余地**が大きい"] }] },
      ],
    },
  },
  // 数値(100% 積み上げ + 縦棒)
  share: {
    title: "売上構成は法人向けに移り、直販比率が 3 年で 2 倍になった",
    lead: "",
    footnote: "出典: 管理会計データ(2023〜2026 年度)",
    body: {
      cols: [
        { rows: [{ type: "cell", head: "販路別の売上構成比" }, { type: "stacked", parts: ["直販", "代理店", "EC"], items: [{ label: "2023 年度", values: [18, 62, 20] }, { label: "2024 年度", values: [24, 55, 21] }, { label: "2025 年度", values: [31, 48, 21] }, { label: "2026 年度", values: [37, 42, 21] }] }] },
        { rows: [{ type: "cell", head: "販路別の粗利率" }, { type: "column", unit: "%", items: [{ label: "直販", value: 41, highlight: true }, { label: "代理店", value: 24 }, { label: "EC", value: 33 }], note: "2026 年度実績" }] },
      ],
    },
  },
  // 段階 × 観点(スイムレーン: 行見出しが矢羽の格子)
  swimlane: {
    title: "AI PoC の課題整理から将来像の合意まで、4 段階 12 週で進める",
    lead: "各段階の成果物を経営会議で確認し、次段階の着手を判断する。",
    footnote: "",
    body: {
      type: "table",
      headShape: "chevron",
      corner: "",
      colHeaders: ["概要", "主なタスク", "アウトプット"],
      rows: [
        { head: "現状把握", cells: ["AI PoC の現状と課題を明確化する", "ヒアリング 12 名と PoC 8 件の棚卸し", "現状分析レポート"] },
        { head: "構想策定", cells: ["業務別の活用シナリオを描く", "優先業務の選定、効果と投資の試算", "活用構想書"], highlight: true },
        { head: "ロードマップ策定", cells: ["段階的な実行計画に落とす", "体制・予算・KPI の設計、リスク整理", "3 か年ロードマップ"] },
        { head: "将来像の合意", cells: ["経営層と到達点を合意する", "経営会議での審議、部門説明会", "合意文書と次年度計画"] },
      ],
    },
  },
  // 列挙パターン 3(横配置 + 下に説明。兄弟列の行揃え)
  enumerate: {
    title: "インタビューでは入社理由より現状課題の指摘が多く、組織制度への不満が最も強い",
    lead: "",
    footnote: "出典: 中堅社員 18 名インタビュー(2026年8月)",
    body: {
      cols: [
        { rows: [{ type: "cell", head: "パーパス" }, { type: "cell", items: ["社員への浸透が不十分", "実務との乖離を感じる"] }] },
        { rows: [{ type: "cell", head: "事業" }, { type: "cell", items: ["既存事業の成長鈍化が顕著", "新規事業の立ち上げが遅い"] }] },
        { rows: [{ type: "cell", head: "ヒト" }, { type: "cell", items: ["優秀な人材は多いが連携不足", "中堅層の離職率が高い"] }] },
        { rows: [{ type: "cell", head: "組織制度", highlight: true }, { type: "cell", items: ["評価制度が不透明で納得感が低い", "キャリアパスが見えず中堅が流出する"] }] },
      ],
    },
  },
  // 拡散(1 つ → 複数)
  diffuse: {
    title: "営業員の 2 割削減が新規開拓・既存深耕・提案品質の 3 面で売上減少を招いた",
    lead: "",
    footnote: "出典: 営業本部 月次報告(2026年4〜8月)",
    body: {
      cols: [
        { type: "cell", head: "起点", text: "営業員を **2 割削減**(120 名 → 96 名)。担当顧客数は 1 人あたり 1.4 倍に増加", fill: "light" },
        { type: "arrow" },
        {
          rows: [
            { type: "cell", head: "新規開拓の停滞", text: "新規訪問 月 480 → 290 件、新規受注 3 割減", highlight: true },
            { type: "cell", head: "既存顧客の深耕不足", text: "上位 50 社の訪問 月 2 → 1 回、追加受注 15% 減" },
            { type: "cell", head: "提案品質の低下", text: "提案書の作成時間が半減、勝率 42% → 31%" },
          ],
        },
      ],
      weights: [4, 0.4, 6],
    },
  },
  // 階層(非対称・ネストした箇条書き)
  hierarchy: {
    title: "売上減少の 5 割は商品力と営業力の内部要因で説明できる",
    lead: "",
    footnote: "出典: 経営企画部 要因分析(2026年8月)",
    body: {
      cols: [
        { type: "cell", head: "外部要因(30%)", items: ["**市場環境**", ["市場全体の縮小(10%)", "規制強化の影響(5%)"], "**競合要因**", ["新規参入の増加(10%)", "価格競争の激化(5%)"]] },
        { type: "cell", head: "内部要因(50%)", highlight: true, items: ["**商品力**", ["商品の陳腐化(15%)", "品質問題の発生(10%)", "ラインナップ不足(5%)"], "**営業力**", ["営業人員の減少(10%)", "営業効率の低下(10%)"]] },
      ],
    },
  },
  // 2 軸で分類(4 象限)
  quadrant: {
    title: "重要かつ緊急の 2 施策に来月の投資を集中させる",
    lead: "",
    footnote: "",
    body: {
      type: "table",
      corner: "",
      colHeaders: ["緊急度 低", "緊急度 高"],
      axes: { x: "緊急度", y: "重要度" },
      rows: [
        { head: "重要度 高", cells: ["計画的に着手\n・基幹システム刷新\n・人材育成体系の整備", { text: "即対応\n・受注管理の障害対策\n・主要顧客の契約更新", highlight: true }] },
        { head: "重要度 低", cells: ["見送り\n・社内報のリニューアル", "委任\n・備品発注の効率化\n・会議室予約の自動化"] },
      ],
    },
  },
  // 役割分担(RACI。セル単位の強調)
  raci: {
    title: "設計と開発の実行責任を IT 部門に一本化し、業務部門は要件と受入に集中する",
    lead: "",
    footnote: "R: 実行責任 / A: 説明責任 / C: 相談 / I: 情報共有",
    body: {
      type: "table",
      corner: "タスク",
      colHeaders: ["PM", "業務部門", "IT 部門", "品質管理"],
      rows: [
        { head: "要件定義", cells: [{ text: "R", highlight: true }, "A", "C", "I"] },
        { head: "設計", cells: ["A", "C", { text: "R", highlight: true }, "C"] },
        { head: "開発", cells: ["A", "I", "R", "C"] },
        { head: "テスト", cells: ["C", "C", "C", "R"] },
        { head: "リリース", cells: ["A", "R", "R", "C"] },
      ],
    },
  },
  // スケジュール(ガント。塗りセル)
  gantt: {
    title: "9 月末リリースに向け、開発とテストを 6 月から並行させる",
    lead: "",
    footnote: "",
    body: {
      type: "table",
      corner: "タスク",
      colHeaders: ["4 月", "5 月", "6 月", "7 月", "8 月", "9 月"],
      rows: [
        { head: "要件定義", cells: [{ text: "", fill: "dark" }, "", "", "", "", ""] },
        { head: "設計", cells: ["", { text: "", fill: "dark" }, { text: "", fill: "dark" }, "", "", ""] },
        { head: "開発", cells: ["", "", { text: "", fill: "dark" }, { text: "", fill: "dark" }, { text: "", fill: "dark" }, ""] },
        { head: "テスト", cells: ["", "", "", "", { text: "", fill: "dark" }, { text: "", fill: "dark" }] },
        { head: "リリース", cells: ["", "", "", "", "", { text: "▲", fill: "accent" }] },
      ],
    },
  },
  // ガント(専用レンダラ。期間 × タスクの帯、四半期の上位見出し、マイルストーン、今日)
  ganttchart: {
    title: "基幹刷新は設計を 8 月に確定し、来年 3 月中旬に本番稼働する",
    lead: "テストと移行リハーサルを年明けに集中させ、本番前の山場を 2 か月に収める。",
    footnote: "",
    body: {
      type: "gantt",
      groups: [{ text: "2026 年度上期", span: 3 }, { text: "下期", span: 6 }, { text: "2027", span: 3 }],
      periods: ["7月", "8月", "9月", "10月", "11月", "12月", "1月", "2月", "3月", "4月", "5月", "6月"],
      today: "9月",
      tasks: [
        { label: "要件定義", start: "7月", end: "8月" },
        { label: "基本設計", start: "8月", end: "9月" },
        { label: "詳細設計", start: "9月", end: "10月" },
        { label: "開発", start: "10月", end: "1月", highlight: true },
        { label: "結合テスト", start: "1月", end: "2月" },
        { label: "受入テスト", start: "2月", end: "3月" },
        { label: "移行リハーサル", start: "3月", end: "3月", note: "2 回" },
        { label: "本番移行・安定化", start: "4月", end: "5月" },
      ],
      milestones: [{ label: "設計完了", at: "10月" }, { label: "開発完了", at: "1月" }, { label: "本番稼働", at: "4月" }],
    },
  },
  // 体制図(箱 + コネクタ)
  org: {
    title: "推進体制は経営・実行・品質の 3 層に分け、責任を明確にする",
    lead: "",
    footnote: "",
    body: {
      type: "org",
      root: {
        label: "プロジェクトオーナー",
        sub: "執行役員 1 名",
        children: [
          {
            label: "ステアリングコミッティ",
            sub: "経営企画・IT・業務部門長",
            children: [
              {
                label: "プロジェクトマネージャー",
                sub: "専任 1 名",
                highlight: true,
                children: [
                  { label: "業務チーム", sub: "3 名 / 要件・受入" },
                  { label: "IT チーム", sub: "5 名 / 設計・開発", children: [{ label: "開発ベンダー", sub: "8 名" }, { label: "インフラベンダー", sub: "3 名" }] },
                  { label: "品質管理", sub: "2 名 / テスト" },
                  { label: "PMO", sub: "1 名 / 進捗・課題" },
                ],
              },
            ],
          },
        ],
      },
    },
  },
  sections: {
    sections: [
      { title: "市場環境と競合動向", summary: "国内市場は年率 4% で縮小し、上位 3 社の寡占が進む", slides: 1 },
      { title: "自社の現状と課題", summary: "主力製品の粗利率低下と営業リソースの分散が課題", slides: 1 },
      { title: "戦略オプションの評価", summary: "B 案(既存販路活用)が市場性と実現性を両立", slides: 1 },
      { title: "実行計画とマイルストーン", summary: "6 か月で PoC、12 か月で 2 地域展開", slides: 1 },
    ],
  },
};
SAMPLES.bullets = SAMPLES.pyramid;
SAMPLES.boxes = SAMPLES.compare;
SAMPLES.columns = SAMPLES.compare;
SAMPLES.matrix = SAMPLES.insight;
SAMPLES.process = SAMPLES.roadmap;
SAMPLES.bars = SAMPLES.report;

function pickMock(hint, prompt) {
  if (hint && SAMPLES[hint]) return SAMPLES[hint];
  const p = (hint || "") + " " + (prompt || "");
  if (/section|セクション|分割|構成案/i.test(p)) return SAMPLES.sections;
  if (/swimlane|スイムレーン|アプローチ|段階.*観点/i.test(p)) return SAMPLES.swimlane;
  if (/raci|役割分担/i.test(p)) return SAMPLES.raci;
  if (/gantt|ガント|スケジュール/i.test(p)) return SAMPLES.gantt;
  if (/象限|quadrant|緊急/i.test(p)) return SAMPLES.quadrant;
  if (/階層|要因分析|hierarchy/i.test(p)) return SAMPLES.hierarchy;
  if (/matrix|マトリクス|比較表|評価|table|表/i.test(p)) return SAMPLES.insight;
  if (/process|ステップ|手順|プロセス|ロードマップ|roadmap/i.test(p)) return SAMPLES.roadmap;
  if (/report|調査|報告|グラフ|数値|kpi/i.test(p)) return SAMPLES.report;
  if (/compare|比較|vs|内製|外注/i.test(p)) return SAMPLES.compare;
  if (/flow|課題|打ち手|効果|因果/i.test(p)) return SAMPLES.flow;
  return SAMPLES.pyramid;
}

module.exports = { SAMPLES, pickMock };
