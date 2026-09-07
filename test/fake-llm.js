"use strict";
/**
 * OpenAI 互換のダミー LLM サーバ(統合テスト用)。
 *  - GET は 403(WAF 模倣)。POST /v1/chat/completions のみ受け付け、/chat/completions は 404(→ /v1 フォールバック検証)。
 *  - response_format 指定時は素の JSON、無い時はコードフェンス付きで返す(extractJson 検証)。
 *  - 受け取った messages の内容を /last で確認できる。
 * 起動: node test/fake-llm.js [port]
 */
const http = require("http");
const { SAMPLES } = require("../server/mock");

const port = Number(process.argv[2] || 4555);
let last = null;

const server = http.createServer((req, res) => {
  const send = (code, obj) => {
    res.writeHead(code, { "Content-Type": "application/json" });
    res.end(JSON.stringify(obj));
  };
  if (req.url === "/last") return send(200, last || {});
  if (req.method === "GET") return send(403, { error: "WAF: GET blocked" });
  let body = "";
  req.on("data", (c) => (body += c));
  req.on("end", () => {
    if (req.url === "/v1/models") return send(200, { data: [{ id: "fake-fast" }, { id: "fake-reasoning" }] });
    if (req.url !== "/v1/chat/completions") return send(404, { error: "not found: " + req.url });
    if (!/^Bearer test-key$/.test(req.headers.authorization || "")) return send(401, { error: "bad key" });
    const j = JSON.parse(body);
    last = j;
    const user = (j.messages.find((m) => m.role === "user") || {}).content || "";
    const m = user.match(/【レイアウトの希望[^】]*】\s*(\w+)/);
    const spec = m && m[1] === "bars"
      ? { panelCount: 1, title: "売上は前年から増加した", body: { type: "bars", unit: "億円", items: [{ label: "前年", value: 12 }, { label: "今年", value: 18 }] }, note: "既存顧客の拡大が増加を支えた" }
      : { panelCount: 2, title: "実績を踏まえて次の施策を進める", panels: [{ head: "実績", items: ["売上が増えた", "既存顧客の取引が拡大した"] }, { head: "次の施策", text: "既存顧客への提案を進める", fill: "dark", highlight: true }] };
    const content = j.response_format ? JSON.stringify(spec) : "了解しました。\n```json\n" + JSON.stringify(spec, null, 1) + "\n```";
    send(200, {
      id: "chatcmpl-fake",
      model: j.model,
      choices: [{ index: 0, message: { role: "assistant", content }, finish_reason: "stop" }],
      usage: { prompt_tokens: 500, completion_tokens: 300, total_tokens: 800 },
    });
  });
});
server.listen(port, () => console.log("fake LLM on http://localhost:" + port));
