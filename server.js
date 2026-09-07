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
const { pickMock, SAMPLES } = require("./server/mock");
const icons = require("./server/icons");
const SlideLayout = require("./public/layout.js");

const PORT = Number(process.env.PORT || 3000);
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
  const { prompt = "", context = "", hint = "", model = "", mock = false, maxSlides = 2, variant = 0 } = req.body || {};
  const c = llm.cfg();
  const useMock = mock === true || req.query.mock === "1" || !c.configured;
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
      const messages = buildMessages({ prompt, context, hint, maxSlides });
      const r = await llm.chat({ messages, model: model || undefined, maxTokens: 3000, temperature: 0.2, jsonMode: true });
      raw = extractJson(r.content);
      meta = { model: r.model, usage: r.usage, mock: false, endpoint: r.url };
      // 密度チェック: 1 枚に収めると全体を 3pt 以上縮めるか溢れるなら、2 枚に分けるよう 1 回だけ再依頼する(情報を落とさず、文字を小さくしない)
      const maxN = Math.max(1, Math.min(2, Number(maxSlides) || 2));
      const first = shapeResponse(raw, SlideLayout.normalizeSpec, maxN);
      if (maxN >= 2 && first.slides && first.slides.length === 1) {
        try {
          const lay = SlideLayout.layout(first.slides[0], { width: 960, height: 540, profile: { bodyFontSize: 14 } });
          // 「密」= 全体を 3pt 以上縮めても入らない / 格子・ガントが溢れる。部品単位の軽い縮小(shrunk)だけでは分けない(4 段の箱フローを 2 枚に割るような過剰分割を防ぐ)
          // ガントは行を詰めて 1 枚に収める(分割しない)。格子の溢れと、全体 3pt 以上の縮小だけを「密」とみなす
          const isGantt = first.slides[0] && first.slides[0].body && first.slides[0].body.type === "gantt";
          const dense = !isGantt && ((lay.fonts.grown || 0) <= -3 || lay.warnings.some((w) => /table overflow/.test(w)));
          if (dense) {
            const msgs2 = messages.concat([
              { role: "assistant", content: JSON.stringify(raw) },
              { role: "user", content: "上の構成は 1 枚に収まらない密度です(文字を縮めないと入らない)。内容を一切落とさずに slides を 2 枚に分けて JSON だけを返してください。1 枚目=結論と主要な根拠、2 枚目=詳細(格子・残りの項目)。同じ内容を 2 枚で繰り返さない。" },
            ]);
            const r2 = await llm.chat({ messages: msgs2, model: model || undefined, maxTokens: 4000, temperature: 0.2, jsonMode: true });
            const raw2 = extractJson(r2.content);
            const second = shapeResponse(raw2, SlideLayout.normalizeSpec, maxN);
            if (second.slides && second.slides.length === 2) {
              raw = raw2;
              meta.split = true;
              meta.usage = { prompt_tokens: (r.usage && r.usage.prompt_tokens || 0) + (r2.usage && r2.usage.prompt_tokens || 0), completion_tokens: (r.usage && r.usage.completion_tokens || 0) + (r2.usage && r2.usage.completion_tokens || 0) };
              console.log(`[density] split into 2 slides (grown=${lay.fonts.grown || 0}, warnings=${lay.warnings.length})`);
            }
          }
        } catch (e) {
          console.log("[density] check skipped: " + (e && e.message));
        }
      }
    }
    const shaped = shapeResponse(raw, SlideLayout.normalizeSpec, Math.max(1, Math.min(2, Number(maxSlides) || 2)));
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
      console.log(`[generate] ${meta.model} ${ms}ms in=${u.prompt_tokens || "?"} out=${u.completion_tokens || "?"} slides=${shaped.slides ? shaped.slides.length : 0}${shaped.sections ? " sections=" + shaped.sections.length : ""}`);
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
