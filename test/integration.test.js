"use strict";
/**
 * サーバ ⇄ LLM の統合テスト(ダミー LLM を使用。実 LiteLLM は .env 設定後に手動で確認)。
 *   node test/integration.test.js
 * 検証項目:
 *  - /chat/completions 404 → /v1/chat/completions フォールバック
 *  - Authorization ヘッダ / response_format(json_object) の送信
 *  - レイアウト指定(hint)がプロンプトに反映され、spec が正規化されて返る
 *  - コードフェンス付き応答の抽出(response_format 非対応モデル相当)
 */
const assert = require("assert");
const { spawn } = require("child_process");
const path = require("path");
const https = require("https");
const http = require("http");

const FAKE_PORT = 4555;
const APP_PORT = 3999;
// 開発証明書が無い環境(CI・クローン直後)では HTTP で起動して同じ検証をする
const HAS_CERTS = (() => {
  try {
    require("fs").accessSync(require("path").join(require("os").homedir(), ".office-addin-dev-certs", "localhost.crt"));
    return true;
  } catch (_) {
    return false;
  }
})();
const SCHEME = HAS_CERTS ? "https" : "http";
const root = path.join(__dirname, "..");

function wait(ms) {
  return new Promise((r) => setTimeout(r, ms));
}
function req(url, body) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const mod = u.protocol === "https:" ? https : http;
    const r = mod.request(
      u,
      { method: body ? "POST" : "GET", headers: { "Content-Type": "application/json" }, rejectUnauthorized: false },
      (res) => {
        let d = "";
        res.on("data", (c) => (d += c));
        res.on("end", () => resolve({ status: res.statusCode, json: d ? JSON.parse(d) : null }));
      }
    );
    r.on("error", reject);
    if (body) r.write(JSON.stringify(body));
    r.end();
  });
}

(async () => {
  const fake = spawn(process.execPath, [path.join(__dirname, "fake-llm.js"), String(FAKE_PORT)], { stdio: "inherit" });
  const app = spawn(process.execPath, [path.join(root, "server.js")], {
    stdio: "inherit",
    env: Object.assign({}, process.env, {
      PORT: String(APP_PORT),
      LLM_BASE_URL: `http://localhost:${FAKE_PORT}`, // 末尾 /v1 無し → フォールバック検証
      LLM_API_KEY: "test-key",
      LLM_MODEL: "fake-fast",
      DEBUG_SNAPSHOT: "0",
      ALLOW_HTTP: HAS_CERTS ? "0" : "1",
      LOG_REQUESTS: "0",
    }),
  });
  try {
    const base = `${SCHEME}://localhost:${APP_PORT}`;
    // サーバの起動待ち: 固定待ちにすると初回起動(証明書読み込み)が間に合わずに落ちるのでポーリングする
    let ready = false;
    for (let i = 0; i < 60 && !ready; i++) {
      try {
        await req(`${base}/api/health`);
        ready = true;
      } catch (_) {
        await wait(500);
      }
    }
    if (!ready) throw new Error(`server did not start on ${base} within 30s`);

    const h = await req(`${base}/api/health`);
    assert.strictEqual(h.json.configured, true, "health.configured");
    assert.strictEqual(h.json.model, "fake-fast");

    const models = await req(`${base}/api/models`);
    assert.deepStrictEqual(models.json.models, ["fake-fast", "fake-reasoning"], "models via POST fallback (GET is WAF-blocked)");

    // 1) 通常生成(json_object)
    const g = await req(`${base}/api/generate`, { prompt: "新規事業 3 案を比較", context: "A案…", hint: "matrix" });
    assert.strictEqual(g.status, 200, JSON.stringify(g.json));
    assert.strictEqual(g.json.mock, false);
    assert.strictEqual(g.json.slides.length, 1);
    assert.ok(g.json.spec.body.cols || g.json.spec.body.type, "spec normalized to grid");
    assert.strictEqual(g.json.endpoint, `http://localhost:${FAKE_PORT}/v1/chat/completions`, "fell back to /v1");
    assert.strictEqual(g.json.usage.total_tokens, 800);

    const last = (await req(`http://localhost:${FAKE_PORT}/last`)).json;
    assert.strictEqual(last.model, "fake-fast");
    assert.deepStrictEqual(last.response_format, { type: "json_object" });
    assert.ok(last.messages[0].role === "system" && /JSON/.test(last.messages[0].content), "system prompt present");
    assert.ok(/【素材/.test(last.messages[1].content) && /matrix/.test(last.messages[1].content), "context + hint in user msg");
    assert.ok(last.max_tokens > 0 && last.stream === false);

    // 2) モデル上書き
    const g2 = await req(`${base}/api/generate`, { prompt: "テスト", hint: "bars", model: "fake-reasoning" });
    assert.strictEqual(g2.json.model, "fake-reasoning");
    assert.strictEqual(g2.json.spec.body.type, "bars", "one full-width chart remains the main visual");
    assert.strictEqual(g2.json.spec.note.type, "cell", "one plain note accompanies the chart");
    assert.strictEqual(g.json.spec.panelCount, 2);
    assert.ok(g.json.spec.body.cols.every((p) => p.type === "cell" && !p.fill && !p.highlight), "generation removes invented panel emphasis");

    // 3) 空プロンプトは 400
    const g3 = await req(`${base}/api/generate`, { prompt: "", context: "" });
    assert.strictEqual(g3.status, 400);

    // 4) mock 強制 & sections モード
    const g4 = await req(`${base}/api/generate`, { prompt: "x", hint: "roadmap", mock: true });
    assert.strictEqual(g4.json.mock, true);
    assert.strictEqual(g4.json.spec.body.rows[0].rows[0].cols[0].shape, "home", "roadmap sample starts with a chevron band (desugared)");
    // 4b) mock の variant ローテーション(並列 N 案の検証用): variant が違えば別サンプル、同じなら同じ
    const v0 = await req(`${base}/api/generate`, { prompt: "x", hint: "roadmap", mock: true, variant: 0 });
    const v1 = await req(`${base}/api/generate`, { prompt: "x", hint: "roadmap", mock: true, variant: 1 });
    const v1b = await req(`${base}/api/generate`, { prompt: "x", hint: "roadmap", mock: true, variant: 1 });
    assert.notStrictEqual(JSON.stringify(v0.json.slides), JSON.stringify(v1.json.slides), "variant 1 は variant 0 と別サンプル");
    assert.strictEqual(JSON.stringify(v1.json.slides), JSON.stringify(v1b.json.slides), "同じ variant は同じサンプル");
    const g5 = await req(`${base}/api/generate`, { prompt: "x", hint: "sections", mock: true });
    assert.ok(Array.isArray(g5.json.sections) && g5.json.sections.length === 4, "sections passthrough");
    assert.strictEqual(g5.json.spec, null);

    console.log("integration tests OK");
  } finally {
    fake.kill();
    app.kill();
  }
})().catch((e) => {
  console.error("integration test FAILED:", e);
  process.exit(1);
});
