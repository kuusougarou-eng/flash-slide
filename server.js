"use strict";
/**
 * Flash Slide — PowerPoint アドイン用ローカルサーバ
 *  - https://localhost:3455/            タスクペイン静的ファイル (public/)
 *  - POST /api/generate                 prompt/context → slide spec JSON (LiteLLM 経由)
 *  - GET  /api/health                   設定状態
 *  - GET  /api/models                   LiteLLM のモデル一覧
 *  - POST /api/debug/snapshot           生成スライドの PNG を debug/ に保存 (DEBUG_SNAPSHOT=1 のとき)
 */
require("dotenv").config();
const fs = require("fs");
const os = require("os");
const path = require("path");
const https = require("https");
const express = require("express");

const llm = require("./server/llm");
const { buildMessages, extractJson, shapeResponse } = require("./server/prompt");
const coverage = require("./server/coverage");
const { pickMock, SAMPLES } = require("./server/mock");
const icons = require("./server/icons");
const SlideLayout = require("./public/layout.js");

const PORT = Number(process.env.PORT || 3000);
// 推論は 1 回。密度・取りこぼし・章立ての「聞き直し」は 1 回あたり 3〜6 秒かかり(実測)、
// 同じ入力でも発動するかが変わるため体感が安定しない。品質は 1 回の応答 + エンジン側の
// 決定論的な処理で担保する。比較したいときだけ MULTI_PASS=1 で従来の聞き直しを有効にする。
const MULTI_PASS = process.env.MULTI_PASS === "1";
// 収まらないときにエンジンが増やしてよい上限。既定の 1〜2 枚を超えるのは、読める大きさで入らないときだけ
const SPLIT_CAP = Math.max(2, Number(process.env.SPLIT_CAP) || 4);
const app = express();
app.use(express.json({ limit: "20mb" }));
app.use((req, res, next) => {
  res.setHeader("Cache-Control", "no-store");
  if (process.env.LOG_REQUESTS === "1") console.log(new Date().toISOString(), req.method, req.url, req.headers["user-agent"] ? req.headers["user-agent"].slice(0, 60) : "");
  next();
});
// 静的ファイルは常に最新を配信(WebView2 のキャッシュで古い layout.js が使われる事故を防ぐ)
app.use(
  express.static(path.join(__dirname, "public"), {
    etag: false,
    lastModified: false,
    cacheControl: false,
    setHeaders: (res) => {
      res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate");
      res.setHeader("Pragma", "no-cache");
      res.setHeader("Expires", "0");
    },
  })
);

icons.mount(app);

app.get("/api/health", async (req, res) => {
  const c = llm.cfg();
  res.json({
    ok: true,
    configured: c.configured,
    baseUrl: c.base ? c.base.replace(/^(https?:\/\/[^/]+).*$/, "$1") : "",
    model: c.model,
    mockDefault: !c.configured,
    // 開発用: AUTO_RUN="flow,report" のように指定するとタスクペイン起動時にモック生成を自動実行
    autoRun: process.env.AUTO_RUN || "",
    devTrigger: process.env.DEV_TRIGGER === "1",
    // 開発者モード(タスクペインの「スライド仕様 JSON」編集欄を表示)。URL の ?dev=1 でも有効化できる
    devMode: process.env.DEV_MODE === "1" || process.env.DEV_TRIGGER === "1",
  });
});

app.get("/api/models", async (req, res) => {
  const c = llm.cfg();
  const models = await llm.listModels();
  res.json({ models, default: c.model });
});

app.post("/api/generate", async (req, res) => {
  const t0 = Date.now();
  const { prompt = "", context = "", hint = "", model = "", mock = false, maxSlides = 2, variant = 0, allowSections = true, split = true } = req.body || {};
  const c = llm.cfg();
  const useMock = mock === true || req.query.mock === "1" || !c.configured;
  const intent = ["auto", "matrix", "outline", "relational"][Math.max(0, Math.min(3, Math.trunc(Number(variant) || 0)))];
  const normalizeResponseSpec = (s) => SlideLayout.normalizeSpec(useMock ? s : { ...s, layoutIntent: intent });
  try {
    let raw, meta;
    if (useMock) {
      raw = pickMock(hint === "auto" ? "" : hint, prompt);
      // 並列 N 案の検証用: variant>0 なら別サンプルへ回す(本物の LLM は同じ入力でも毎回別構成を返すので不要)
      const v = Number(variant) || 0;
      if (v > 0) {
        const keys = Object.keys(SAMPLES).filter((k) => k !== "sections");
        const i = keys.findIndex((k) => SAMPLES[k] === raw);
        raw = SAMPLES[keys[((i < 0 ? 0 : i) + v) % keys.length]];
      }
      meta = { model: "mock", usage: null, mock: true };
    } else {
      if (!prompt.trim() && !context.trim()) return res.status(400).json({ error: "prompt が空です" });
      const messages = buildMessages({ prompt, context, hint, maxSlides, variant, allowSections });
      // LLM の往復回数を数える。密度・取りこぼし・章立ての再依頼で 1 回の生成が 2〜4 往復になることがあり、
      // 体感の遅さはほぼここで決まる(レイアウト計算は 1 枚 0.5ms 前後で効かない)
      meta = { roundTrips: 0, phases: [] };
      // どの段階に何 ms 掛かったかを測る。体感の遅さの内訳を推測でなく実測で出すため
      const phase = async (name, fn) => {
        const t = Date.now();
        try {
          return await fn();
        } finally {
          meta.phases.push(name + "=" + (Date.now() - t));
        }
      };
      const ask = (name, msgs) =>
        phase(name, () => {
          meta.roundTrips++;
          return llm.chat({ messages: msgs, model: model || undefined, maxTokens: 4000, temperature: 0.2, jsonMode: true });
        });
      meta.roundTrips++;
      const r = await phase("llm1", () => llm.chat({ messages, model: model || undefined, maxTokens: 3000, temperature: 0.2, jsonMode: true }));
      raw = extractJson(r.content);
      meta = { model: r.model, usage: r.usage, mock: false, endpoint: r.url, roundTrips: meta.roundTrips, phases: meta.phases };
      // 章立ての 1 章分では sections を受け付けない。章がさらに章立てを返すと、その章のスライドが 0 枚のまま黙って落ちる
      if (!allowSections && raw && Array.isArray(raw.sections) && raw.sections.length && !(Array.isArray(raw.slides) && raw.slides.length)) {
        const msgsNo = messages.concat([
          { role: "assistant", content: JSON.stringify(raw) },
          { role: "user", content: "sections は返さないでください。この依頼は章立ての 1 章分です。上で章に分けた内容を一切落とさず、slides(1〜2 枚)の JSON だけを返してください。" },
        ]);
        const rNo = await ask("llmChapter", msgsNo);
        const rawNo = extractJson(rNo.content);
        if (rawNo && ((Array.isArray(rawNo.slides) && rawNo.slides.length) || rawNo.title)) {
          raw = rawNo;
          meta.usage = { prompt_tokens: ((r.usage && r.usage.prompt_tokens) || 0) + ((rNo.usage && rNo.usage.prompt_tokens) || 0), completion_tokens: ((r.usage && r.usage.completion_tokens) || 0) + ((rNo.usage && rNo.usage.completion_tokens) || 0) };
          console.log("[density] chapter re-asked for slides (sections not allowed)");
        }
      }
      // 密度チェック: 1 枚に収めると全体を 3pt 以上縮めるか溢れるなら、2 枚に分けるよう 1 回だけ再依頼する(情報を落とさず、文字を小さくしない)
      const maxN = Math.max(1, Math.min(2, Number(maxSlides) || 2));
      const first = shapeResponse(raw, normalizeResponseSpec, maxN);
      if (maxN >= 2 && first.slides && first.slides.length >= 1) {
        try {
          if (first.slides.length === 1) {
            const tLay = Date.now();
            const lay = SlideLayout.layout(first.slides[0], { width: 960, height: 540, profile: { bodyFontSize: 14 } });
            meta.phases.push("densityLayout=" + (Date.now() - tLay));
            // 「密」= 全体を 3pt 以上縮めても入らない / 格子・ガントが溢れる。部品単位の軽い縮小(shrunk)だけでは分けない(4 段の箱フローを 2 枚に割るような過剰分割を防ぐ)
            // ガントは行を詰めて 1 枚に収める(分割しない)。格子の溢れと、全体 3pt 以上の縮小だけを「密」とみなす
            const isGantt = first.slides[0] && first.slides[0].body && first.slides[0].body.type === "gantt";
            // 「密」= 全体を 3pt 以上縮めないと入らない / 格子が溢れる / 上限で内容が落ちた
            const dropped = lay.warnings.some((w) => /content dropped|nested list overflow/.test(w)) || lay.prims.some((p) => p.shrunk && p.fontSize < 16);
            const dense = dropped || (!isGantt && ((lay.fonts.grown || 0) <= -3 || lay.warnings.some((w) => /table overflow/.test(w))));
            if (dense && MULTI_PASS) {
              const msgs2 = messages.concat([
                { role: "assistant", content: JSON.stringify(raw) },
                { role: "user", content: "上の構成は 1 枚に収まらない密度です(文字を縮めないと入らない)。内容を一切落とさずに slides を 2 枚に分けて JSON だけを返してください。1 枚目=結論と主要な根拠、2 枚目=詳細(格子・残りの項目)。同じ内容を 2 枚で繰り返さない。" },
              ]);
              const r2 = await ask("llmSplit", msgs2);
              const raw2 = extractJson(r2.content);
              const second = shapeResponse(raw2, normalizeResponseSpec, maxN);
              if (second.slides && second.slides.length === 2) {
                raw = raw2;
                meta.split = true;
                meta.usage = { prompt_tokens: (r.usage && r.usage.prompt_tokens || 0) + (r2.usage && r2.usage.prompt_tokens || 0), completion_tokens: (r.usage && r.usage.completion_tokens || 0) + (r2.usage && r2.usage.completion_tokens || 0) };
                console.log(`[density] split into 2 slides (grown=${lay.fonts.grown || 0}, warnings=${lay.warnings.length})`);
              }
            }
          }
          // 取りこぼしチェック: 入力にあった数値・固有名詞が spec に現れないなら、
          // それだけを挙げて 1 回聞き直す(コンテキストは足さず、与えられた内容を漏らさない)。
          // allowSections=false(章の生成)では prompt に元の入力全体がそのまま含まれたまま渡ってくるため、
          // この章だけを見て「入力全体」を分母にすると常に大半が欠落扱いになる(意味のある検知にならない)。
          // 章の情報保持は章立て自体(summary に事実を持たせる指示)で担保し、ここでは測らない
          if (allowSections && MULTI_PASS) {
            const tCov = Date.now();
            const src = prompt + "\n" + context;
            const want = coverage.atoms(src);
            // 日本語の内容語も対象にする(数値だけでは「兼業モデル」のような語の取りこぼしを拾えない)
            const wantTerms = coverage.terms(src);
            const shapedNow = shapeResponse(raw, normalizeResponseSpec, maxN);
            const nowText = JSON.stringify(shapedNow.slides || shapedNow);
            const lost = coverage.missing(want, nowText).concat(coverage.missingTerms(wantTerms, nowText));
            meta.phases.push("coverageScan=" + (Date.now() - tCov));
            if (want.length + wantTerms.length >= 8 && lost.length >= 2) {
              const msgs3 = messages.concat([
                { role: "assistant", content: JSON.stringify(raw) },
                { role: "user", content: "上の構成は入力にあった次の内容を落としています: " + lost.join(" / ") + "\n入力に無いことは足さず、これらを本文・note・2 枚目のいずれかに載せた JSON を返してください。1 枚に収まらなければ slides を 2 枚にしてください。" },
              ]);
              const r3 = await ask("llmCoverage", msgs3);
              const raw3 = extractJson(r3.content);
              const third = shapeResponse(raw3, normalizeResponseSpec, maxN);
              const t3 = JSON.stringify(third.slides || third);
              const lost3 = coverage.missing(want, t3).concat(coverage.missingTerms(wantTerms, t3));
              if (lost3.length < lost.length) {
                raw = raw3;
                meta.recovered = lost.length - lost3.length;
                meta.usage = { prompt_tokens: (meta.usage && meta.usage.prompt_tokens || 0) + (r3.usage && r3.usage.prompt_tokens || 0), completion_tokens: (meta.usage && meta.usage.completion_tokens || 0) + (r3.usage && r3.usage.completion_tokens || 0) };
                console.log(`[coverage] recovered ${lost.length - lost3.length}/${lost.length} dropped items (slides=${(third.slides || []).length})`);
              } else {
                console.log(`[coverage] still missing ${lost3.length}/${want.length + wantTerms.length}: ${lost3.slice(0, 6).join(" / ")}`);
              }
            }
          }
          // 最終確認: 2 枚構成でもまだ縮めないと入らない/溢れるなら、文字を潰して押し込む代わりに
          // 章立て(sections)へ切り替える(情報量が2枚という器を超えている。1〜2 枚に収める試みは既に上で尽くした)
          const tFin = Date.now();
          const finalCheck = shapeResponse(raw, normalizeResponseSpec, maxN);
          if (allowSections && MULTI_PASS && finalCheck.slides && finalCheck.slides.length) {
            const stillDense = finalCheck.slides.some((s) => {
              try {
                const lay2 = SlideLayout.layout(s, { width: 960, height: 540, profile: { bodyFontSize: 14 } });
                const dropped2 = lay2.warnings.some((w) => /content dropped|nested list overflow|table overflow/.test(w)) || lay2.prims.some((p) => p.shrunk && p.fontSize < 16);
                // 格子・表のセルは fitBody を通らず shrunk フラグも grown も付かないまま既定下限(14pt)を割ることがある。
                // 「縮んだ形跡があるか」ではなく「実際に版面に出ている文字が読める大きさか」を直接見る
                const bodyFonts = lay2.prims.filter((p) => p.body && p.text && p.fontSize).map((p) => p.fontSize);
                const tooSmall = bodyFonts.length > 0 && Math.min.apply(null, bodyFonts) < 14;
                return dropped2 || tooSmall || (lay2.fonts.grown || 0) <= -3;
              } catch (_) {
                return false;
              }
            });
            meta.phases.push("finalLayout=" + (Date.now() - tFin));
            if (stillDense) {
              const msgsSec = messages.concat([
                { role: "assistant", content: JSON.stringify(raw) },
                { role: "user", content: "上の構成は2枚に収めても文字を大きく縮めないと入らない密度です。内容を一切落とさず、章立て(sections)の JSON だけを返してください。入力に列挙されている論点(項目・領域・施策など)がある場合は、1 章に複数の論点を詰め込まず論点数に応じて章を分けてください(3〜8章。論点が8を超える場合も8章までにまとめてよいが、1章に複数論点を無理に押し込めるくらいなら章を増やす方を優先する)。各章の summary にはその章に入れる事実・数値をすべて含めてください。" },
              ]);
              const rSec = await ask("llmSections", msgsSec);
              const rawSec = extractJson(rSec.content);
              if (rawSec && Array.isArray(rawSec.sections) && rawSec.sections.length) {
                raw = rawSec;
                meta.escalatedToSections = true;
                meta.usage = {
                  prompt_tokens: ((meta.usage && meta.usage.prompt_tokens) || 0) + ((rSec.usage && rSec.usage.prompt_tokens) || 0),
                  completion_tokens: ((meta.usage && meta.usage.completion_tokens) || 0) + ((rSec.usage && rSec.usage.completion_tokens) || 0),
                };
                console.log(`[density] escalated to sections (${rawSec.sections.length} chapters, still dense after 2-slide compression)`);
              }
            }
          }
        } catch (e) {
          console.log("[density] check skipped: " + (e && e.message));
        }
      }
    }
    const shaped = shapeResponse(raw, normalizeResponseSpec, Math.max(1, Math.min(2, Number(maxSlides) || 2)));
    // 収まらないときはエンジンが割る。LLM には聞き直さない(推論は 1 回)。
    // 枚数を指定されているとき(1 枚固定)は割らない。既定は 1〜2 枚だが、内容が読める大きさで
    // 入らないなら最大 4 枚まで増やす。文字を潰す・内容を落とすより枚数を増やす方がよい
    if (shaped.slides && Number(maxSlides) !== 1 && split !== false) {
      const cap = SPLIT_CAP;
      const out = [];
      for (const s of shaped.slides) for (const part of SlideLayout.splitToFit(s, cap - out.length)) out.push(part);
      if (out.length > shaped.slides.length) {
        meta.splitInto = out.length;
        console.log(`[split] ${shaped.slides.length} → ${out.length} slides (deterministic, no re-ask)`);
      }
      shaped.slides = out.slice(0, cap);
    }
    const ms = Date.now() - t0;
    if (process.env.DEBUG_SNAPSHOT === "1" && !useMock) {
      // 生応答と正規化後の spec を残す(崩れの再現・回帰テストの材料)
      try {
        const dir = path.join(__dirname, "debug", "specs");
        fs.mkdirSync(dir, { recursive: true });
        fs.writeFileSync(path.join(dir, `spec-${Date.now()}.json`), JSON.stringify({ prompt: prompt.slice(0, 200), model: meta.model, ms, raw, shaped }, null, 1));
      } catch (_) {}
    }
    // 計測ログ: モデル・所要時間・トークン数(ベンチと本番の比較に使う)
    if (!useMock) {
      const u = meta.usage || {};
      console.log(`[generate] ${meta.model} ${ms}ms x${meta.roundTrips || 1} [${(meta.phases || []).join(" ")}] in=${u.prompt_tokens || "?"} out=${u.completion_tokens || "?"} slides=${shaped.slides ? shaped.slides.length : 0}${shaped.sections ? " sections=" + shaped.sections.length : ""}`);
    }
    // 互換: spec = 1 枚目
    res.json({ ...shaped, spec: shaped.slides ? shaped.slides[0] : null, ms, ...meta });
  } catch (e) {
    console.error("[generate]", e && e.stack || e);
    res.status(500).json({ error: String(e && e.message || e), ms: Date.now() - t0 });
  }
});

// デバッグ: タスクペインから送られたスライド画像を保存
app.post("/api/debug/snapshot", (req, res) => {
  if (process.env.DEBUG_SNAPSHOT !== "1") return res.status(404).end();
  const { png, name = "slide" } = req.body || {};
  if (!png) return res.status(400).json({ error: "png required" });
  const dir = path.join(__dirname, "debug");
  fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, `${name}-${Date.now()}.png`);
  fs.writeFileSync(file, Buffer.from(png.replace(/^data:image\/png;base64,/, ""), "base64"));
  console.log("[snapshot]", file);
  res.json({ file });
});
// 開発用: アクティブなデッキに対して任意のタイミングで生成を発火(clone 検証)
let pendingTrigger = null;
app.post("/api/debug/trigger", (req, res) => {
  if (process.env.DEV_TRIGGER !== "1") return res.status(404).end();
  pendingTrigger = Object.assign({ hint: "", prompt: "", mock: true, tag: "trigger" }, req.body || {});
  console.log("[trigger set]", JSON.stringify(pendingTrigger));
  res.json({ ok: true });
});
const lastPolls = {};
app.get("/api/debug/polls", (req, res) => res.json(lastPolls));
app.get("/api/debug/trigger", (req, res) => {
  if (process.env.DEV_TRIGGER !== "1") return res.status(404).end();
  // 旧バージョンのタスクペイン(別文書に残っているもの)がトリガを横取りしないよう、v=4 以上だけに渡す
  if (String(req.query.v || "") !== "4") return res.json({ trigger: null });
  lastPolls[String(req.query.doc || "") || "(no doc)"] = new Date().toISOString(); // どのペイン(文書)が生きているかの確認用
  // doc を指定したトリガは、その文書(URL 部分一致)を開いているペインにだけ渡す(複数ペイン/複数エージェントの取り合い防止)
  if (pendingTrigger && pendingTrigger.doc) {
    const want = String(pendingTrigger.doc).toLowerCase();
    const have = String(req.query.doc || "").toLowerCase();
    const match = want === "(no doc)" ? !have : have.includes(want);
    if (!match) return res.json({ trigger: null });
  }
  const t = pendingTrigger;
  pendingTrigger = null;
  res.json({ trigger: t });
});
app.post("/api/debug/log", (req, res) => {
  console.log("[taskpane]", JSON.stringify(req.body));
  res.json({ ok: true });
});

// HTTPS (office-addin-dev-certs の証明書)。Office アドインは HTTPS でしか読み込めない。
// 証明書が無い環境(CI・クローン直後)では、ALLOW_HTTP=1 のときだけ HTTP で起動する(テスト用)。
const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
let tls = null;
try {
  tls = { key: fs.readFileSync(path.join(certDir, "localhost.key")), cert: fs.readFileSync(path.join(certDir, "localhost.crt")) };
} catch (_) {
  if (process.env.ALLOW_HTTP !== "1") {
    console.error("開発証明書が見つかりません: " + certDir);
    console.error("  npx office-addin-dev-certs install   を実行してから起動してください。");
    console.error("  (テスト目的で HTTP で起動する場合は ALLOW_HTTP=1)");
    process.exit(1);
  }
}
const banner = () => {
  const c = llm.cfg();
  console.log(`Flash Slide server: ${tls ? "https" : "http"}://localhost:${PORT}`);
  console.log(`LLM configured: ${c.configured} (model=${c.model || "-"})${c.configured ? "" : "  → mock モードで動作"}`);
  if (!tls) console.log("注意: 証明書が無いため HTTP で起動しました。Office アドインとしては読み込めません(テスト専用)。");
};
if (tls) https.createServer(tls, app).listen(PORT, banner);
else require("http").createServer(app).listen(PORT, banner);
