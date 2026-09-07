"use strict";
/**
 * LiteLLM (OpenAI 互換) 呼び出し。
 * - LLM_BASE_URL / LLM_API_KEY / LLM_MODEL は .env から
 * - GET が WAF で拒否される環境があるため、生成は POST のみ。
 * - /chat/completions が 404 の場合は /v1/chat/completions を再試行。
 * - response_format / max_tokens / temperature を受け付けないモデルには、該当パラメータを外して 1 回だけ再試行。
 * - タイムアウトは設けない(重いモデルは待つ)。
 */

function cfg() {
  const base = (process.env.LLM_BASE_URL || "").trim().replace(/\/+$/, "");
  const key = (process.env.LLM_API_KEY || "").trim();
  const model = (process.env.LLM_MODEL || "").trim();
  return { base, key, model, configured: !!(base && key) };
}

function headers(key) {
  return {
    "Content-Type": "application/json",
    Authorization: `Bearer ${key}`,
    "api-key": key, // Azure 互換ゲートウェイ対策
  };
}

async function postJson(url, key, body) {
  const res = await fetch(url, { method: "POST", headers: headers(key), body: JSON.stringify(body) });
  const text = await res.text();
  let json = null;
  try {
    json = JSON.parse(text);
  } catch (_) {}
  return { status: res.status, ok: res.ok, json, text };
}

/**
 * chat completion を実行し、content 文字列を返す。
 * @param {object} opts { messages, model, maxTokens, temperature, jsonMode }
 */
async function chat(opts) {
  const { base, key, model: defModel } = cfg();
  if (!base || !key) throw new Error("LLM_BASE_URL / LLM_API_KEY が未設定です (.env)");
  const model = opts.model || defModel;
  if (!model) throw new Error("LLM_MODEL が未設定です (.env) か、モデルを選択してください");

  const body = {
    model,
    messages: opts.messages,
    temperature: opts.temperature == null ? 0.2 : opts.temperature,
    stream: false,
  };
  if (opts.maxTokens) body.max_tokens = opts.maxTokens;
  if (opts.jsonMode) body.response_format = { type: "json_object" };

  const candidates = base.endsWith("/v1")
    ? [`${base}/chat/completions`]
    : [`${base}/chat/completions`, `${base}/v1/chat/completions`];

  let last = null;
  for (const url of candidates) {
    let r = await postJson(url, key, body);
    // response_format / max_tokens 非対応モデルへのフォールバック
    if (!r.ok && r.status >= 400 && r.status < 500 && r.status !== 404 && r.status !== 401) {
      const msg = (r.text || "").toLowerCase();
      const retryBody = Object.assign({}, body);
      let changed = false;
      if (retryBody.response_format && /response_format|json_object|json/.test(msg)) {
        delete retryBody.response_format;
        changed = true;
      }
      if (retryBody.max_tokens && /max_tokens|max_completion_tokens/.test(msg)) {
        delete retryBody.max_tokens;
        changed = true;
      }
      // temperature を既定値(1)しか受け付けないモデルへの対応
      if (retryBody.temperature != null && /temperature/.test(msg)) {
        delete retryBody.temperature;
        changed = true;
      }
      if (changed) r = await postJson(url, key, retryBody);
    }
    if (r.ok) {
      const content = r.json && r.json.choices && r.json.choices[0] && r.json.choices[0].message
        ? r.json.choices[0].message.content
        : null;
      if (typeof content !== "string") throw new Error("LLM 応答に content がありません: " + r.text.slice(0, 300));
      return { content, usage: r.json.usage || null, model: r.json.model || model, url };
    }
    last = r;
    if (r.status !== 404) break; // 404 のときだけ /v1 を試す
  }
  throw new Error(`LLM error ${last && last.status}: ${(last && last.text || "").slice(0, 500)}`);
}

/** モデル一覧(GET /models → 失敗時 POST も試す)。取れなければ空配列。 */
async function listModels() {
  const { base, key } = cfg();
  if (!base || !key) return [];
  const urls = base.endsWith("/v1") ? [`${base}/models`] : [`${base}/models`, `${base}/v1/models`];
  for (const url of urls) {
    for (const method of ["GET", "POST"]) {
      try {
        const res = await fetch(url, { method, headers: headers(key), body: method === "POST" ? "{}" : undefined });
        if (!res.ok) continue;
        const j = await res.json();
        const arr = Array.isArray(j.data) ? j.data : Array.isArray(j) ? j : [];
        const ids = arr.map((m) => m.id || m.model || m.name).filter(Boolean);
        if (ids.length) return ids;
      } catch (_) {}
    }
  }
  return [];
}

module.exports = { cfg, chat, listModels };
