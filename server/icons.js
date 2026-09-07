"use strict";
/**
 * ピクトグラム: Tabler Icons(MIT, 5,000+ アウトライン SVG)をサーバ側で色付き PNG に描画して返す。
 * LLM は名前を選ぶだけ。タスクペインは PNG(base64)を受け取り、画像としてスライドに配置する。
 *
 *  GET  /api/icons                         … 利用可能なアイコン名(キュレーション済み)
 *  POST /api/icons { items:[{name,color}], size } … まとめて描画 → { icons: { "name|color": base64 } }
 *  GET  /api/icon?name=users&color=404040&size=256 … 単体(デバッグ用、PNG を直接返す)
 */
const fs = require("fs");
const path = require("path");
const { Resvg } = require("@resvg/resvg-js");

const ICON_DIR = path.join(__dirname, "..", "node_modules", "@tabler", "icons", "icons", "outline");
const CACHE_DIR = path.join(__dirname, "..", "icons-cache");

/** ビジネス資料で使う頻度が高い語彙に絞ったキュレーション(存在確認して公開) */
const CURATED = [
  // 人・組織
  "user", "users", "users-group", "user-check", "user-plus", "user-cog", "building", "building-bank", "building-factory-2",
  "building-store", "building-hospital", "building-community", "briefcase", "id-badge", "hierarchy-2", "sitemap", "heart-handshake",
  // 事業・財務
  "coin", "cash", "credit-card", "wallet", "receipt", "report-money", "percentage", "calculator", "scale", "gavel", "contract",
  "chart-bar", "chart-line", "chart-pie", "chart-dots", "chart-arrows", "trending-up", "trending-down", "report-analytics", "presentation",
  "target", "target-arrow", "flag", "trophy", "award", "certificate", "rocket", "growth", "seedling",
  // 業務・プロセス
  "settings", "tool", "tools", "adjustments", "refresh", "repeat", "arrows-exchange", "route", "git-branch", "checklist", "list-check",
  "clipboard-check", "clipboard-list", "calendar", "calendar-event", "clock", "hourglass", "history", "timeline", "stack-2", "layers-subtract",
  "package", "truck", "ship", "plane", "map-pin", "map-2", "world", "globe",
  // IT・データ
  "device-laptop", "device-desktop", "device-mobile", "server", "database", "cloud", "cloud-upload", "cpu", "code", "api",
  "robot", "brain", "network", "plug-connected", "lock", "lock-open", "shield-check", "shield-lock", "key", "fingerprint", "bug",
  "mail", "message", "message-circle", "phone", "video", "speakerphone", "bell", "search", "filter", "zoom-check",
  // 文書・知識
  "file-text", "file-analytics", "file-check", "files", "folder", "book", "notebook", "school", "bulb", "eye", "pencil", "edit",
  // 状態・評価
  "check", "circle-check", "x", "circle-x", "alert-triangle", "alert-circle", "info-circle", "help", "question-mark", "star", "thumb-up",
  "arrow-right", "arrow-down", "arrow-up", "arrows-right-left", "arrow-big-right", "plus", "minus", "equal",
  // 生活・産業
  "home", "car", "bus", "train", "stethoscope", "first-aid-kit", "heart", "leaf", "recycle", "bolt", "flame", "droplet", "sun", "moon",
  "shopping-cart", "basket", "gift", "tag", "tags", "ticket", "camera", "photo", "music", "movie", "coffee", "tent",
];

let availableCache = null;
function available() {
  if (availableCache) return availableCache;
  availableCache = CURATED.filter((n) => fs.existsSync(path.join(ICON_DIR, n + ".svg")));
  return availableCache;
}

function exists(name) {
  return /^[a-z0-9-]+$/.test(name) && fs.existsSync(path.join(ICON_DIR, name + ".svg"));
}

const mem = new Map();
function renderPng(name, color, size) {
  color = normColor(color);
  size = Math.max(16, Math.min(1024, Number(size) || 256));
  const key = `${name}|${color}|${size}`;
  if (mem.has(key)) return mem.get(key);
  fs.mkdirSync(CACHE_DIR, { recursive: true });
  const file = path.join(CACHE_DIR, `${name}_${color.slice(1)}_${size}.png`);
  if (fs.existsSync(file)) {
    const b = fs.readFileSync(file);
    mem.set(key, b);
    return b;
  }
  if (!exists(name)) return null;
  let svg = fs.readFileSync(path.join(ICON_DIR, name + ".svg"), "utf8");
  svg = svg.replace(/stroke="currentColor"/g, `stroke="${color}"`);
  // 太めのストロークで小サイズでも視認性を確保(Tabler 既定 2 → 1.75 だと細いので 2 を維持)
  const png = new Resvg(svg, { fitTo: { mode: "width", value: size }, background: "rgba(0,0,0,0)" }).render().asPng();
  const buf = Buffer.from(png);
  fs.writeFileSync(file, buf);
  mem.set(key, buf);
  return buf;
}
function normColor(c) {
  c = String(c || "404040").replace(/^#/, "");
  if (!/^[0-9a-fA-F]{6}$/.test(c)) c = "404040";
  return "#" + c.toUpperCase();
}

function mount(app) {
  app.get("/api/icons", (req, res) => res.json({ icons: available(), license: "Tabler Icons (MIT)" }));
  app.get("/api/icon", (req, res) => {
    const buf = renderPng(String(req.query.name || ""), req.query.color, req.query.size);
    if (!buf) return res.status(404).json({ error: "unknown icon" });
    res.setHeader("Content-Type", "image/png");
    res.send(buf);
  });
  app.post("/api/icons", (req, res) => {
    const { items = [], size = 256 } = req.body || {};
    const out = {};
    for (const it of items.slice(0, 40)) {
      const buf = renderPng(String(it.name || ""), it.color, size);
      if (buf) out[`${it.name}|${normColor(it.color)}`] = buf.toString("base64");
    }
    res.json({ icons: out });
  });
}

module.exports = { mount, available, renderPng, exists, CURATED };
