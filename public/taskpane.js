/* global Office, SlideLayout, SlideRender */
(function () {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const LS = "flashslide.settings.v3";

  let prepared = null; // Office 側の準備情報(マスタ/サイズ/参照解析)。読み取り完了までは生成不可
  let lastSpecs = null;
  let lastPrompt = null;
  let selectedAccent = "none"; // "auto" | "none" | "#RRGGBB"(単一の状態変数。既定はグレースケール。スウォッチ表示と palette() の両方がこれを見る)
  let customAccent = "#FD5108";
  let devMode = /[?&]dev=1/.test(location.search);
  let refSeq = 0; // 参照読み取りの再入ガード(最後に発火した呼び出しの結果だけを反映)
  let done = false;

  // ---------- 設定 ----------
  function loadSettings() {
    try {
      const s = JSON.parse(localStorage.getItem(LS) || "{}");
      if (/^#[0-9a-fA-F]{6}$/.test(s.customAccent || "")) customAccent = s.customAccent; // 旧版が "none" を保存していることがある
      if (s.accent) setAccent(s.accent);
      if (s.hint) $("hint").value = s.hint;
      if (s.model) $("model").dataset.pending = s.model;
      if (s.prompt) $("prompt").value = s.prompt;
    } catch (_) {}
  }
  function saveSettings() {
    try {
      localStorage.setItem(LS, JSON.stringify({ accent: selectedAccent, customAccent, hint: $("hint").value, model: $("model").value, prompt: $("prompt").value }));
    } catch (_) {}
  }
  function setAccent(v) {
    if (v !== "auto" && v !== "none" && !/^#[0-9a-fA-F]{6}$/.test(v)) return;
    selectedAccent = v === "auto" || v === "none" ? v : v.toUpperCase();
    // "auto" / "none" は小文字のまま保持するので、比較は大文字に揃える
    // (揃えないと isPreset が常に false になり、カスタムのスウォッチが選択表示されてしまう)
    const sel = selectedAccent.toUpperCase();
    const presets = Array.from(document.querySelectorAll(".sw")).map((b) => b.dataset.accent.toUpperCase());
    const isPreset = presets.includes(sel);
    document.querySelectorAll(".sw").forEach((b) => {
      const key = b.dataset.accent.toUpperCase();
      b.classList.toggle("selected", key === sel || (key === "CUSTOM" && !isPreset));
    });
    const custom = document.querySelector(".sw.custom");
    if (!isPreset) {
      customAccent = selectedAccent;
      custom.style.setProperty("--c", customAccent);
      custom.classList.add("hasColor");
      custom.title = "カスタム " + customAccent;
    }
    const uiColor = selectedAccent === "auto" ? "#FD5108" : selectedAccent === "none" ? "#404040" : selectedAccent;
    document.documentElement.style.setProperty("--accent", uiColor);
    updateTuneState();
  }
  /** パレット: 自動(テンプレ準拠)のときは accent を渡さない → 参照スライドの色 or 既定。none は完全モノトーン */
  function palette() {
    return selectedAccent === "auto" ? {} : { accent: selectedAccent };
  }

  // ---------- 入力欄(主役)の状態表示 ----------
  // 「今開いているスライドへの操作指示」ではなく「伝えたい中身」を置く場所だと分かるよう、
  // 例文は指示文ではなくメモそのもの、文字数は枚数の目安として返す。
  function updateCompose() {
    const n = $("prompt").value.trim().length;
    const c = $("counter");
    c.textContent = n ? `${n.toLocaleString()} 字 · 目安 ${n < 700 ? "1 枚" : n < 1800 ? "1〜2 枚" : "2 枚 + 章立ての確認"}` : "";
    c.classList.toggle("long", n >= 1800);
    if (n > 0) document.body.classList.add("typing");
    else if (document.activeElement !== $("prompt")) document.body.classList.remove("typing");
    if (!$("go").classList.contains("busy")) $("goLabel").textContent = goLabel();
  }
  /** 折りたたんだ「仕上げの希望」の現在値を summary に出す(開かなくても分かる) */
  function updateTuneState() {
    const acc = selectedAccent === "auto" ? "色は自動" : selectedAccent === "none" ? "グレースケール" : selectedAccent;
    const lay = $("layout");
    const bits = [acc, (lay.options[lay.selectedIndex] || {}).textContent || "自動"];
    const h = $("hint").value.trim();
    if (h) bits.unshift("「" + (h.length > 14 ? h.slice(0, 14) + "…" : h) + "」");
    if ($("mock").checked) bits.push("モック");
    $("tuneState").textContent = "— " + bits.join(" · ");
  }
  function setZen(on) {
    document.body.classList.toggle("zen", on);
    $("expand").title = on ? "入力欄を戻す(Esc)" : "入力欄を広げる(Esc で戻る)";
    if (on) $("prompt").focus();
  }
  let saveTimer = null;
  function schedulePromptSave() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(saveSettings, 400);
  }

  // ---------- 進捗(ボタン内ラップ) ----------
  let timer = null;
  function busy(btn, label) {
    const t0 = performance.now();
    btn.disabled = true;
    btn.classList.add("busy");
    const lbl = btn.querySelector("span") || btn;
    clearInterval(timer);
    const tick = () => (lbl.textContent = `${label}… ${((performance.now() - t0) / 1000).toFixed(1)}s`);
    tick();
    timer = setInterval(tick, 100);
    return t0;
  }
  function idle(btn, t0, label) {
    clearInterval(timer);
    btn.classList.remove("busy");
    const lbl = btn.querySelector("span") || btn;
    lbl.textContent = label;
    setGoEnabled();
    return t0 ? ((performance.now() - t0) / 1000).toFixed(1) : "";
  }
  function setGoEnabled() {
    const analyzing = $("refStatus").classList.contains("analyzing");
    $("go").disabled = analyzing || $("go").classList.contains("busy");
  }
  function showError(msg) {
    $("error").hidden = !msg;
    $("error").textContent = msg || "";
    if (msg) fetch("/api/debug/log", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ error: msg }) }).catch(() => {});
  }

  // ---------- サーバ ----------
  async function api(path, body) {
    const res = await fetch(path, { method: body ? "POST" : "GET", headers: { "Content-Type": "application/json" }, body: body ? JSON.stringify(body) : undefined });
    const j = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(j.error || res.status + " " + res.statusText);
    return j;
  }
  async function refreshHealth() {
    try {
      const h = await api("/api/health");
      const dot = $("statusDot");
      if (h.configured) {
        dot.className = "dot ok";
        $("statusText").textContent = h.model || "model 未指定";
        $("status").title = "LLM: " + (h.model || "model 未指定") + " @ " + h.baseUrl;
      } else {
        dot.className = "dot mock";
        $("statusText").textContent = "LLM 未設定 → モック動作(.env を設定)";
        $("mock").checked = true;
      }
      if (h.devMode) devMode = true;
      $("devJson").hidden = !devMode;
      const m = await api("/api/models");
      const sel = $("model");
      const pending = sel.dataset.pending || "";
      (m.models || []).forEach((id) => {
        const o = document.createElement("option");
        o.value = id;
        o.textContent = id + (id === m.default ? " (既定)" : "");
        sel.appendChild(o);
      });
      if (pending && Array.from(sel.options).some((o) => o.value === pending)) sel.value = pending;
      if (h.autoRun) autoRunSequence(String(h.autoRun).split(",").map((s) => s.trim()).filter(Boolean));
      if (h.devTrigger) startTriggerPolling();
    } catch (e) {
      $("statusDot").className = "dot err";
      $("statusText").textContent = "サーバに接続できません: " + e.message;
    }
  }

  // ---------- 参照スライドの読み取り(状態: analyzing / clone / layout / error) ----------
  function setRefState(state, tip, label) {
    const b = $("refStatus");
    b.className = "refStatus " + state;
    b.title = tip;
    if (label != null) $("refText").textContent = label;
    setGoEnabled();
  }
  async function refreshReference() {
    const seq = ++refSeq;
    setRefState("analyzing", "スライドを読み取り中…", "読み取り中…");
    try {
      const info = await SlideRender.prepare();
      if (seq !== refSeq) return; // 後から発火した読み取りが優先
      prepared = info;
      const size = `${Math.round(info.width)}×${Math.round(info.height)}pt`;
      const where = info.selectedIndex != null ? `スライド ${info.selectedIndex + 1}` : "スライド未選択";
      // 「開いているスライドへの指示」ではなく「ここに新しく足す」ことを、出力先の明示で言外に伝える
      const dest = info.selectedIndex != null ? `スライド ${info.selectedIndex + 1} の次に追加` : "デッキの最後に追加";
      if (info.mode === "clone") setRefState("clone", `${where} のデザインを引き継ぎます · ${size}\n${info.analysisNote || ""}`, `${dest}・デザイン継承`);
      else setRefState("layout", `${where} の直後に既定レイアウトで追加します · ${size}\n${info.analysisNote || ""}`, `${dest}・既定レイアウト`);
      fillLayouts(info);
    } catch (e) {
      if (seq !== refSeq) return;
      prepared = null;
      setRefState("error", "スライドを読み取れませんでした(既定レイアウトで生成します)\n" + e.message, "既定レイアウトで追加");
      fillLayouts(null);
    }
  }
  /** T6: マスタのレイアウト一覧をセレクトに反映。clone が成立するときは自動固定(disabled) */
  function fillLayouts(info) {
    const sel = $("layout");
    const prev = sel.value;
    sel.innerHTML = "";
    const auto = document.createElement("option");
    auto.value = "";
    auto.textContent = info && info.mode === "clone" ? "自動(テンプレ追従)" : "自動";
    sel.appendChild(auto);
    const layouts = (info && info.layouts) || [];
    layouts.forEach((l) => {
      const o = document.createElement("option");
      o.value = l.id;
      o.textContent = l.name;
      sel.appendChild(o);
    });
    // マスタが変わって以前の選択が無ければ自動に戻す
    sel.value = layouts.some((l) => l.id === prev) ? prev : "";
    sel.disabled = !!(info && info.mode === "clone") || !layouts.length;
    updateTuneState();
  }
  let selTimer = null;
  function onSelectionChanged() {
    clearTimeout(selTimer);
    selTimer = setTimeout(refreshReference, 400);
  }

  // ---------- 生成(N 案並列 → 最初の 1 案を即挿入 → 残りをカードで提示 → 押して差し替え) ----------
  // 実測(2026-09-06): 4 案並列でも壁時計は 1 案 +0.6〜2s、構成は 4/4 別物、title は同じ主旨。
  // 「仕上げの希望」を言語で書かせずに、見て選ぶだけでレイアウトを変えられるようにする。
  const VARIANTS = 4;
  let batch = null; // { prompt, prepared, cands: [{specs, gen}], insertedIdx, insertedIds, tag }
  let pagesMode = 0; // 0=おまかせ(最大 2) / 1=1 枚 / 2=2 枚

  function goLabel() {
    return done && batch && batch.prompt === $("prompt").value.trim() ? "再生成" : "スライド生成";
  }
  function pagesArgs() {
    if (pagesMode === 1) return { maxSlides: 1, hintExtra: "" };
    if (pagesMode === 2) return { maxSlides: 2, hintExtra: "必ず 2 枚(slides)に分ける。1 枚目=結論と全体像、2 枚目=詳細・補足。" };
    return { maxSlides: 2, hintExtra: "" };
  }
  function renderOpts(pre, ignoreIds) {
    return { prepared: pre, palette: palette(), layoutId: $("layout").value || undefined, ignoreIds: ignoreIds || [] };
  }
  /** 最初に成功した 1 件を返す(全滅なら最後のエラーで reject) */
  function firstSettled(tasks) {
    return new Promise((resolve, reject) => {
      let pending = tasks.length;
      let lastErr = null;
      tasks.forEach((t, i) =>
        t.then(
          (gen) => resolve({ i, gen }),
          (e) => {
            lastErr = e;
            if (--pending === 0) reject(lastErr);
          }
        )
      );
    });
  }
  function logResult(results, gen, tag) {
    results.forEach((r, i) => {
      fetch("/api/debug/log", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ slide: tag, slideId: r.slideId, debug: r.debug, llmMs: gen.ms, model: gen.model, usage: gen.usage }) }).catch(() => {});
      if (r.png) debugSnapshot(r.png, tag + (results.length > 1 ? "-" + (i + 1) : ""));
    });
  }

  async function generate(overrides) {
    overrides = overrides || {};
    const prompt = overrides.prompt != null ? overrides.prompt : $("prompt").value.trim();
    const mock = $("mock").checked;
    if (!prompt && !mock) {
      showError("スライドにしたい内容を貼り付けてください。");
      return;
    }
    showError("");
    saveSettings();
    setZen(false);
    $("sections").hidden = true;
    const go = $("go");
    const t0 = busy(go, "生成中");
    const n = Math.max(1, Number(overrides.n) || VARIANTS);
    const pg = pagesArgs();
    const hint = [overrides.hint != null ? overrides.hint : $("hint").value, pg.hintExtra].filter(Boolean).join(" / ");
    // 同じ入力での再生成は差し替え(前回挿入分を消す)。入力が変わっていれば追加
    const replacing = !!(batch && batch.prompt === prompt && batch.insertedIds.length);
    const oldIds = replacing ? batch.insertedIds.slice() : [];
    // 出力先はバッチ開始時のスナップショットで固定する。差し替えでは前回のスナップショットを使う
    // (生成後は選択が自分の生成スライドへ移るので、現在の prepared を使うと「消す対象」を参照に複製しようとして落ちる)
    const pre = replacing && batch.prepared ? batch.prepared : prepared;
    const nb = { prompt, prepared: pre, cands: [], insertedIdx: -1, insertedIds: [], tag: overrides.tag || "slide" };
    batch = nb;
    const tasks = [];
    for (let i = 0; i < n; i++) tasks.push(api("/api/generate", { prompt, hint, model: $("model").value, mock, maxSlides: pg.maxSlides, variant: i }));
    tasks.forEach((t) => t.catch(() => {})); // 未処理 rejection の警告を抑える(結果は allSettled で拾う)
    try {
      const first = await firstSettled(tasks);
      if (batch !== nb) return; // 新しいバッチが始まっていたら捨てる
      lastPrompt = prompt;
      if (first.gen.sections) {
        batch = null;
        idle(go, t0, goLabel());
        showSections(first.gen.sections, prompt);
        return;
      }
      go.querySelector("span").textContent = "描画中…";
      // 先に挿入し、成功してから旧スライドを消す(挿入に失敗しても既存スライドを失わない)
      const results = await SlideRender.renderSpecs(first.gen.slides, renderOpts(pre, oldIds));
      if (batch !== nb) return;
      if (oldIds.length) {
        await SlideRender.deleteSlides(oldIds);
        await SlideRender.selectSlides(results.map((r) => r.slideId));
      }
      nb.cands[first.i] = { specs: first.gen.slides, gen: first.gen };
      nb.insertedIdx = first.i;
      nb.insertedIds = results.map((r) => r.slideId);
      lastSpecs = first.gen.slides;
      done = true;
      const total = idle(go, t0, goLabel());
      setDone(true);
      showResult(first.gen.slides, results, first.gen, total);
      showCandidates(n, false); // 残りはスケルトン
      logResult(results, first.gen, nb.tag);
      // 残りの案を待ってカードを埋める
      const all = await Promise.allSettled(tasks);
      if (batch !== nb) return;
      all.forEach((r, i) => {
        if (r.status === "fulfilled" && r.value && Array.isArray(r.value.slides) && r.value.slides.length && !nb.cands[i]) nb.cands[i] = { specs: r.value.slides, gen: r.value };
      });
      // Keep the inserted candidate; show only genuinely different compositions.
      const seen = new Set([SlideLayout.compositionKey(nb.cands[nb.insertedIdx].specs)]);
      nb.cands.forEach((c, i) => {
        if (!c || i === nb.insertedIdx) return;
        const key = SlideLayout.compositionKey(c.specs);
        if (seen.has(key)) nb.cands[i] = null;
        else seen.add(key);
      });
      showCandidates(n, true);
    } catch (e) {
      if (batch === nb) batch = null;
      idle(go, t0, goLabel());
      showError("エラー: " + (e && e.message ? e.message : e) + (e && e.debugInfo ? "\n" + JSON.stringify(e.debugInfo) : ""));
      console.error(e);
    }
  }

  /** カードを押して別案に差し替える(前回挿入分を消してから同じ場所に入れる) */
  async function replaceWith(i) {
    const b = batch;
    const c = b && b.cands[i];
    if (!c || i === b.insertedIdx) return;
    showError("");
    const go = $("go");
    const t0 = busy(go, "差し替え中");
    document.querySelectorAll(".cand").forEach((el) => (el.disabled = true));
    try {
      const oldIds = b.insertedIds.slice();
      const results = await SlideRender.renderSpecs(c.specs, renderOpts(b.prepared, oldIds));
      await SlideRender.deleteSlides(oldIds);
      await SlideRender.selectSlides(results.map((r) => r.slideId)); // 旧スライドを消すと表示が隣へ逃げるので、新しい方を見せる
      b.insertedIdx = i;
      b.insertedIds = results.map((r) => r.slideId);
      lastSpecs = c.specs;
      const total = idle(go, t0, goLabel());
      showResult(c.specs, results, c.gen, total);
      showCandidates(b.cands.length, true);
      logResult(results, c.gen, b.tag + "-alt" + (i + 1));
    } catch (e) {
      idle(go, t0, goLabel());
      showError("差し替えに失敗しました: " + (e && e.message ? e.message : e));
      showCandidates(b.cands.length, true);
    }
  }

  /** 生成資料の一覧(元の案の順序で固定。差し替えても位置は動かず、印だけ移る)。settled=false のあいだは未着をスケルトンで示す */
  function showCandidates(n, settled) {
    const box = $("cands");
    box.innerHTML = "";
    const b = batch;
    $("result").classList.toggle("noCands", !b);
    if (!b) return;
    let alts = 0;
    for (let i = 0; i < n; i++) {
      const c = b.cands[i];
      if (!c && settled) continue; // 失敗した案は出さない
      const el = document.createElement("button");
      el.type = "button";
      el.className = "cand";
      if (!c) {
        el.classList.add("skel");
        el.disabled = true;
        el.innerHTML = '<div class="thumb"></div><span class="cap"><span class="capTitle">生成中…</span></span>';
      } else {
        const spec = c.specs[0];
        const inserted = i === b.insertedIdx;
        el.appendChild(thumbnail(spec, b.prepared));
        const cap = document.createElement("span");
        cap.className = "cap";
        const t = document.createElement("span");
        t.className = "capTitle";
        t.textContent = spec.title || "";
        t.title = spec.title || "";
        const st = document.createElement("span");
        st.className = "capState";
        st.textContent = (inserted ? "挿入済み" : "") + (c.specs.length > 1 ? (inserted ? " · " : "") + c.specs.length + "枚" : "");
        cap.appendChild(t);
        cap.appendChild(st);
        el.appendChild(cap);
        if (inserted) el.classList.add("inserted");
        else {
          el.title = "この案に差し替える";
          el.addEventListener("click", () => replaceWith(i));
          alts++;
        }
      }
      box.appendChild(el);
    }
    $("candsNote").textContent = !settled ? "残りの案を生成しています。" : alts ? "押した案に差し替えます。表示は近似プレビューです。" : "他の案は得られませんでした。再生成で引き直せます。";
  }

  // ---------- サムネ: layout.js のプリミティブを HTML で描き、CSS transform で縮小 ----------
  // 実描画(render.js)と同じ layout() 結果を使うので構成・比率は同じ。フォントメトリクス差で折り返しは 1 行ずれうる。
  const SHAPES = {
    homePlate: "polygon(0 0, 82% 0, 100% 50%, 82% 100%, 0 100%)",
    chevron: "polygon(0 0, 82% 0, 100% 50%, 82% 100%, 0 100%, 18% 50%)",
    trapezoid: "polygon(20% 0, 80% 0, 100% 100%, 0 100%)",
    triangle: "polygon(50% 0, 100% 100%, 0 100%)",
    rightArrow: "polygon(0 25%, 70% 25%, 70% 0, 100% 50%, 70% 100%, 70% 75%, 0 75%)",
    downArrow: "polygon(25% 0, 75% 0, 75% 70%, 100% 70%, 50% 100%, 0 70%, 25% 70%)",
    ellipse: "ellipse(50% 50% at 50% 50%)",
  };
  function esc(t) {
    return String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
  function richText(text, ranges, bullets) {
    const marks = new Array(text.length).fill(false);
    (ranges || []).forEach(([st, l]) => {
      for (let i = st; i < st + l && i < text.length; i++) marks[i] = true;
    });
    let out = bullets ? "• " : "";
    let open = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (ch === "\n") {
        if (open) {
          out += "</b>";
          open = false;
        }
        out += "<br>" + (bullets ? "• " : "");
        continue;
      }
      if (marks[i] && !open) {
        out += "<b>";
        open = true;
      } else if (!marks[i] && open) {
        out += "</b>";
        open = false;
      }
      out += esc(ch);
    }
    if (open) out += "</b>";
    return out;
  }
  function primBox(p) {
    const d = document.createElement("div");
    d.className = "pr";
    d.style.left = p.x + "px";
    d.style.top = p.y + "px";
    d.style.width = Math.max(1, p.w) + "px";
    d.style.height = Math.max(1, p.h) + "px";
    if (p.fill) d.style.background = p.fill;
    if (p.line) d.style.border = `${p.lineWeight || 0.75}px solid ${p.line}`;
    if (p.shape === "roundRect") d.style.borderRadius = "6px";
    else if (SHAPES[p.shape]) d.style.clipPath = SHAPES[p.shape];
    if (p.rotation) d.style.transform = `rotate(${p.rotation}deg)`;
    if (p.text) {
      const pad = p.pad == null ? 10 : p.pad;
      d.style.padding = `${Math.min(pad, 8)}px ${pad}px`;
      d.style.fontSize = (p.fontSize || 18) + "px";
      d.style.fontWeight = p.bold ? 700 : 400;
      d.style.color = p.color || "#1a1a1a";
      d.style.textAlign = p.align || "left";
      d.style.justifyContent = p.valign === "middle" ? "center" : p.valign === "bottom" ? "flex-end" : "flex-start";
      if (p.fontName) d.style.fontFamily = `"${p.fontName}", "Yu Gothic UI", Meiryo, sans-serif`;
      // .pr は縦 flex(valign 用)なので、テキストノードと <b> が別アイテムに割れないよう 1 つの span で包む
      d.innerHTML = "<span>" + richText(String(p.text), p.bold ? null : p.boldRanges, !!p.bullets) + "</span>";
    }
    return d;
  }
  function primLine(p) {
    const d = document.createElement("div");
    d.className = "pr";
    const w = Math.abs(p.x2 - p.x1), h = Math.abs(p.y2 - p.y1), wt = p.weight || 1;
    d.style.left = Math.min(p.x1, p.x2) + "px";
    d.style.top = (Math.min(p.y1, p.y2) - (h === 0 ? wt / 2 : 0)) + "px";
    d.style.width = (w === 0 ? wt : w) + "px";
    d.style.height = (h === 0 ? wt : h) + "px";
    d.style.background = p.color || "#7F7F7F";
    return d;
  }
  function thumbnail(spec, pre) {
    const W = (pre && pre.width) || 960, H = (pre && pre.height) || 540;
    const useClone = !!(pre && pre.mode === "clone" && pre.selectedId && pre.has18);
    const lay = SlideLayout.layout(spec, { width: W, height: H, palette: palette(), profile: useClone ? pre.profile : null });
    const wrap = document.createElement("div");
    wrap.className = "thumb";
    wrap.style.aspectRatio = `${W} / ${H}`;
    const inner = document.createElement("div");
    inner.className = "thumbIn";
    inner.style.width = W + "px";
    inner.style.height = H + "px";
    // clone ではタイトル/リードをテンプレの枠に流し込む(プリミティブに含まれない)ので、枠の位置に描く
    if (useClone && pre.profile) {
      const t = pre.profile.title, l = pre.profile.lead;
      if (t && t.reuse && lay.header.titleText) inner.appendChild(primBox({ x: t.x, y: t.y, w: t.w, h: t.h, text: lay.header.titleText, fontSize: lay.header.titleFontSize || t.fontSize || 24, bold: t.bold !== false, color: t.color || "#1a1a1a", align: t.align || "left", valign: t.valign || "middle", pad: 6, fontName: t.fontName || pre.profile.fontName }));
      if (l && l.reuse && lay.header.leadText) inner.appendChild(primBox({ x: l.x, y: l.y, w: l.w, h: l.h, text: lay.header.leadText, fontSize: lay.header.leadFontSize || l.fontSize || 16, bold: false, color: l.color || "#404040", align: l.align || "left", valign: l.valign || "top", pad: 6, fontName: l.fontName || pre.profile.fontName }));
    }
    for (const p of lay.prims) {
      if (p.kind === "rect") inner.appendChild(primBox(p));
      else if (p.kind === "line") inner.appendChild(primLine(p));
      else if (p.kind === "table") for (const f of p.fallback || []) inner.appendChild(primBox(f));
      else if (p.kind === "image") {
        const img = document.createElement("img");
        img.alt = "";
        img.src = "/api/icon?name=" + encodeURIComponent(p.name) + "&color=" + encodeURIComponent(p.color || "#252525") + "&size=256";
        Object.assign(img.style, { position: "absolute", left: p.x + "px", top: p.y + "px", width: p.w + "px", height: p.h + "px" });
        inner.appendChild(img);
      }
    }
    wrap.appendChild(inner);
    const fit = () => {
      const k = wrap.clientWidth / W;
      if (k > 0) inner.style.transform = `scale(${k})`;
    };
    if (typeof ResizeObserver !== "undefined") new ResizeObserver(fit).observe(wrap);
    else requestAnimationFrame(fit);
    return wrap;
  }

  /** 生成後: 入力欄を小さく畳む(内容は保持)。フォーカス/「入力を開く」で戻る */
  function setDone(on) {
    document.body.classList.toggle("done", on);
    document.body.classList.remove("editing");
    $("fold").hidden = !on;
    $("fold").textContent = "入力を開く";
  }
  function setEditing(on) {
    document.body.classList.toggle("editing", on);
    $("fold").textContent = on ? "入力を畳む" : "入力を開く";
  }

  /** 入力が多すぎるとき: LLM が提案した章構成を表示し、ユーザーが調整してから生成 */
  function showSections(sections, prompt) {
    const box = $("sectionList");
    box.innerHTML = "";
    sections.forEach((s, i) => {
      const row = document.createElement("div");
      row.className = "secRow";
      row.innerHTML =
        `<label><input type="checkbox" class="secOn" checked /> <span class="secNo">${i + 1}</span></label>` +
        `<div class="secFields"><input type="text" class="secTitle" value="" /><input type="text" class="secSummary" value="" /></div>`;
      row.querySelector(".secTitle").value = s.title;
      row.querySelector(".secSummary").value = s.summary;
      box.appendChild(row);
    });
    $("sections").hidden = false;
    $("secGo").onclick = async () => {
      const rows = Array.from(box.querySelectorAll(".secRow")).filter((r) => r.querySelector(".secOn").checked);
      const secs = rows.map((r) => ({ title: r.querySelector(".secTitle").value.trim(), summary: r.querySelector(".secSummary").value.trim() }));
      if (!secs.length) return;
      await generateSections(secs, prompt);
    };
    $("secCompress").onclick = () => generate({ prompt, hint: "必ず 2 枚以内に圧縮する(sections は返さない)。重要度の低い情報は捨てる。" });
  }

  /** 章ごとに LLM を並列実行し、順番に描画 */
  async function generateSections(secs, prompt) {
    showError("");
    const btn = $("secGo");
    const t0 = busy(btn, "生成中");
    $("go").disabled = true;
    try {
      const mock = $("mock").checked;
      const gens = await Promise.all(
        secs.map((s, i) =>
          api("/api/generate", {
            prompt: `${prompt}\n\n【今回作る章(全 ${secs.length} 章のうち ${i + 1} 章目)】${s.title} — ${s.summary}\nこの章の内容だけを載せる。他の章の内容は含めない。文字を縮めないと入らないほど密なら 2 枚に分けてよい(章の中の情報を削らない)。`,
            hint: $("hint").value,
            model: $("model").value,
            mock,
            maxSlides: 2,
            allowSections: false,
          })
        )
      );
      btn.querySelector("span").textContent = "描画中…";
      // 章が 1 枚も返さなかったら黙って落とさない(章がさらに章立てを返して 0 枚になる事故が実機で出た)
      const empty = secs.filter((s, i) => !(gens[i].slides || []).length).map((s, i) => s.title);
      const specs = gens.flatMap((g) => g.slides || []);
      lastSpecs = specs;
      const results = await SlideRender.renderSpecs(specs, { prepared, palette: palette(), layoutId: $("layout").value || undefined });
      const total = idle(btn, t0, "この構成で生成");
      $("sections").hidden = true;
      done = true;
      batch = { prompt, prepared, cands: [], insertedIdx: -1, insertedIds: results.map((r) => r.slideId), tag: "section" };
      idle($("go"), null, goLabel());
      setDone(true);
      showResult(specs, results, { model: gens[0].model, mock: gens[0].mock, ms: Math.max.apply(null, gens.map((g) => g.ms)) }, total);
      showCandidates(0, true);
      if (empty.length) showError("次の章は生成できませんでした: " + empty.join(" / "));
      results.forEach((r, i) => r.png && debugSnapshot(r.png, "section-" + (i + 1)));
    } catch (e) {
      idle(btn, t0, "この構成で生成");
      idle($("go"), null, goLabel());
      showError("エラー: " + (e && e.message ? e.message : e));
    }
  }

  async function redraw() {
    showError("");
    let specs;
    try {
      const j = JSON.parse($("specJson").value);
      specs = (Array.isArray(j) ? j : [j]).map(SlideLayout.normalizeSpec);
    } catch (e) {
      showError("JSON が不正です: " + e.message);
      return;
    }
    const btn = $("redraw");
    const t0 = busy(btn, "描画中");
    try {
      const results = await SlideRender.renderSpecs(specs, { prepared, palette: palette(), layoutId: $("layout").value || undefined });
      const total = idle(btn, t0, "この JSON で再描画");
      lastSpecs = specs;
      batch = null;
      showResult(specs, results, { mock: false, model: "(再描画)", ms: 0 }, total);
      showCandidates(0, true);
      results.forEach((r, i) => r.png && debugSnapshot(r.png, "redraw-" + (i + 1)));
    } catch (e) {
      idle(btn, t0, "この JSON で再描画");
      showError("エラー: " + (e && e.message ? e.message : e));
    }
  }

  function showResult(specs, results, gen, totalSec) {
    $("result").hidden = false;
    // 所要時間・モード・警告などの詳細は UI に出さない(開発者向けログにだけ残す)
    const warn = results.flatMap((r) => r.warnings || []);
    fetch("/api/debug/log", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ result: { slides: results.length, mode: (results[0] || {}).mode, totalSec, llmMs: gen.ms, model: gen.model, mock: gen.mock, warnings: warn } }) }).catch(() => {});
    $("specJson").value = JSON.stringify(specs.length === 1 ? specs[0] : specs, null, 2);
  }

  // ---------- 開発用 ----------
  async function autoRunSequence(hints) {
    $("mock").checked = true;
    while ($("refStatus").classList.contains("analyzing")) await new Promise((r) => setTimeout(r, 200));
    for (const h of hints) {
      await generate({ prompt: "自動テスト " + h, hint: h, tag: h, n: 1 });
      await new Promise((r) => setTimeout(r, 500));
    }
    fetch("/api/debug/log", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ autoRunDone: hints }) }).catch(() => {});
  }
  // 開発用: サーバに置かれたトリガをポーリングし、アクティブデッキに対して生成(clone 検証)
  let triggerBusy = false;
  let docName = null; // 文書の識別子(URL が取れない文書ではプレゼンのタイトル = ファイル名)
  async function resolveDocName() {
    let doc = "";
    try {
      doc = (Office.context.document && Office.context.document.url) || "";
    } catch (_) {}
    if (!doc) {
      try {
        await PowerPoint.run(async (ctx) => {
          ctx.presentation.load("title");
          await ctx.sync();
          doc = ctx.presentation.title || "";
        });
      } catch (_) {}
    }
    return doc;
  }
  function startTriggerPolling() {
    setInterval(async () => {
      if (triggerBusy) return;
      let t;
      try {
        if (docName == null) docName = await resolveDocName();
        t = (await api("/api/debug/trigger?v=4&doc=" + encodeURIComponent(docName || ""))).trigger;
      } catch (_) {
        return;
      }
      if (!t) return;
      triggerBusy = true;
      try {
        if (t.hint === "makeref") {
          await SlideRender.devBuildReference();
          await refreshReference();
          fetch("/api/debug/log", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ makeref: true, mode: prepared && prepared.mode, note: prepared && prepared.analysisNote }) }).catch(() => {});
        } else if (t.hint === "probe") {
          // 開発用: Office.js の機能有無を実機で確認する(ネイティブグラフの可否など)
          const req = ["1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10", "1.11", "1.12"].filter((v) => SlideRender.supports(v));
          let shapeApi = [];
          try {
            await PowerPoint.run(async (ctx) => {
              const sh = ctx.presentation.slides.getItemAt(0).shapes;
              shapeApi = Object.keys(Object.getPrototypeOf(sh)).filter((k) => /^add/.test(k));
            });
          } catch (e) {
            shapeApi = ["error: " + e.message];
          }
          fetch("/api/debug/log", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ probe: true, requirementSets: req, shapeAdders: shapeApi, hasChartInNamespace: typeof PowerPoint !== "undefined" && Object.keys(PowerPoint).filter((k) => /chart/i.test(k)) }) }).catch(() => {});
          return;
        } else if (t.hint === "probeindent") {
          // 開発用: 段落ごとに字下げ(indentLevel)を変えられるかを実機の描画で見る。
          // API は例外を出さないので、箱を残して PNG で目視する(確認後は probeclean で消す)
          await PowerPoint.run(async (ctx) => {
            const slide = ctx.presentation.slides.getItemAt(Number(t.i) || 0);
            const text = "親項目 1\n子項目 1-1\n子項目 1-2\n親項目 2";
            const box = slide.shapes.addTextBox(text, { left: 60, top: 120, width: 500, height: 200 });
            box.name = "FS_probeindent";
            const tr = box.textFrame.textRange;
            tr.font.size = 18;
            tr.paragraphFormat.bulletFormat.visible = true;
            const from = "親項目 1\n".length;
            const len = "子項目 1-1\n子項目 1-2".length;
            tr.getSubstring(from, len).paragraphFormat.indentLevel = 1; // 2〜3 段落目だけ 1 段下げたい
            await ctx.sync();
          });
          fetch("/api/debug/log", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ probeindent: "drawn" }) }).catch(() => {});
          return;
        } else if (t.hint === "probeclean") {
          await PowerPoint.run(async (ctx) => {
            const shapes = ctx.presentation.slides.getItemAt(Number(t.i) || 0).shapes;
            shapes.load("items/name");
            await ctx.sync();
            shapes.items.filter((s) => /^FS_probe/.test(s.name)).forEach((s) => s.delete());
            await ctx.sync();
          });
          return;
        } else if (t.hint === "probepara") {
          // 開発用: 段落まわりの API 面を実機で確認する(段落間隔・行間・ぶら下げが設定できるか)。
          // 使い捨てのテキストボックスを自分で作って試し、最後に必ず消す(参照デッキの中身は触らない)
          const out = { probepara: true, proto: {}, writes: {} };
          try {
            await PowerPoint.run(async (ctx) => {
              const slide = ctx.presentation.slides.getItemAt(0);
              const box = slide.shapes.addTextBox("段落1\n段落2\n段落3", { left: 10, top: 10, width: 200, height: 100 });
              box.name = "FS_probe";
              const tr = box.textFrame.textRange;
              const pf = tr.paragraphFormat;
              out.proto.textFrame = Object.keys(Object.getPrototypeOf(box.textFrame));
              out.proto.textRange = Object.keys(Object.getPrototypeOf(tr));
              out.proto.paragraphFormat = Object.keys(Object.getPrototypeOf(pf));
              out.proto.bulletFormat = Object.keys(Object.getPrototypeOf(pf.bulletFormat));
              out.scalars = { paragraphFormat: pf._scalarPropertyNames, bulletFormat: pf.bulletFormat._scalarPropertyNames, textFrame: box.textFrame._scalarPropertyNames };
              await ctx.sync();
              // 1 つずつ別の sync で試す(まとめると誰が落としたか分からない)
              for (const [key, value] of [["spaceAfter", 12], ["spaceBefore", 12], ["lineSpacing", 1.5], ["leftIndent", 24], ["hangingIndent", 18], ["firstLineIndent", -18], ["indentLevel", 1]]) {
                try {
                  box.textFrame.textRange.paragraphFormat[key] = value;
                  await ctx.sync();
                  // 追跡対象でないプロパティは「ただの JS 代入」で例外も出ない。load して読み戻せるかまで見る
                  try {
                    const pf2 = box.textFrame.textRange.paragraphFormat;
                    pf2.load(key);
                    await ctx.sync();
                    out.writes[key] = "set+load ok: " + JSON.stringify(pf2[key]);
                  } catch (e2) {
                    out.writes[key] = "set ok / load NG(未対応): " + (e2 && e2.message ? e2.message.slice(0, 70) : e2);
                  }
                } catch (e) {
                  out.writes[key] = "NG: " + (e && e.message ? e.message.slice(0, 90) : e);
                }
              }
              box.delete();
              await ctx.sync();
            });
          } catch (e) {
            out.error = String(e && e.message || e);
          }
          fetch("/api/debug/log", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(out) }).catch(() => {});
          return;
        } else if (t.hint === "reload") {
          location.reload(); // 静的ファイル(layout/render/taskpane.js)を読み直す(no-store 配信)
          return;
        } else if (t.hint === "renderspec" && t.doc && Array.isArray(t.specs)) {
          // Reproducible rendering QA: only a specifically addressed test deck.
          await refreshReference();
          // 本番と同じ入口(normalizeSpec)を通す。文法の正規化と決定論的な後処理の両方が掛かる
          const specs = t.specs.map(SlideLayout.normalizeSpec);
          const results = await SlideRender.renderSpecs(specs, { prepared, palette: palette() });
          fetch("/api/debug/log", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ renderedSpecs: t.tag || "composition", results: results.map(({ slideId, warnings, debug }) => ({ slideId, warnings, debug })) }) }).catch(() => {});
        } else if (t.hint === "setprompt") {
          // 開発用: 入力欄の中身だけ差し替える(空欄=初回状態の見た目を実機で確認するため)
          $("prompt").value = t.prompt || "";
          updateCompose();
          saveSettings();
        } else if (t.hint === "pick") {
          await replaceWith(Number(t.i) || 0); // 開発用: 他の案 i に差し替え(カードのクリック相当)
        } else if (t.hint === "secgo") {
          // 開発用: 章立ての「この構成で生成」を押す(ペインは WebView2 で、UI Automation から DOM に届かない)
          if ($("sections").hidden || typeof $("secGo").onclick !== "function") throw new Error("章立ての提案が出ていません");
          await $("secGo").onclick();
          fetch("/api/debug/log", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ secgo: true, tag: "section", slideIds: (batch && batch.insertedIds) || [], count: (lastSpecs || []).length }) }).catch(() => {});
        } else if (t.hint === "tune") {
          $("tune").open = t.open !== false; // 開発用: 「仕上げの希望」を開閉
        } else if (t.hint === "autoopen") {
          // 開発用: 自動オープンのチェックを切り替えて文書設定を保存(change イベント経由)
          $("autoOpen").checked = t.on !== false;
          $("autoOpen").dispatchEvent(new Event("change"));
        } else if (t.hint === "zen") {
          setZen(!document.body.classList.contains("zen"));
        } else if (t.hint === "reanalyze") {
          await refreshReference();
          fetch("/api/debug/log", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ reanalyze: true, mode: prepared && prepared.mode, note: prepared && prepared.analysisNote, layoutDefaults: prepared && prepared.layoutDefaults, layoutDefaultsError: prepared && prepared.layoutDefaultsError, refLayout: prepared && prepared.refLayoutName, layouts: prepared && prepared.layouts }) }).catch(() => {});
        } else {
          while ($("refStatus").classList.contains("analyzing")) await new Promise((r) => setTimeout(r, 200));
          $("mock").checked = t.mock !== false;
          if (t.accent) setAccent(t.accent);
          if (t.layoutId != null) $("layout").value = t.layoutId;
          if (t.model != null) $("model").value = t.model;
          if (t.pages != null) {
            pagesMode = Number(t.pages) || 0; // 開発用: 枚数 select 相当
            $("pages").value = String(pagesMode);
          }
          // useInput: 入力欄の文字をそのまま使う(ユーザー経路の再現。再生成=差し替えの判定を含めて検証できる)
          await generate({ prompt: t.useInput ? undefined : t.prompt || "自動テスト " + (t.hint || ""), hint: t.hint || "", tag: t.tag || "trigger", n: t.n || 1 });
        }
      } finally {
        triggerBusy = false;
      }
    }, 1200);
  }
  function debugSnapshot(png, name) {
    fetch("/api/debug/snapshot", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ png, name: name || "slide" }) }).catch(() => {});
  }

  // ---------- 自動オープン(文書スコープの設定) ----------
  const AUTO_KEY = "Office.AutoShowTaskpaneWithDocument";
  function initAutoOpen() {
    let ok = false;
    try {
      ok = Office.context.requirements.isSetSupported("AddInCommands", "1.1") && !!(Office.context.document && Office.context.document.settings);
    } catch (_) {}
    $("autoRow").hidden = !ok;
    if (!ok) return;
    try {
      $("autoOpen").checked = Office.context.document.settings.get(AUTO_KEY) === true;
    } catch (_) {}
    $("autoOpen").addEventListener("change", () => {
      const on = $("autoOpen").checked;
      try {
        if (on) Office.context.document.settings.set(AUTO_KEY, true);
        else Office.context.document.settings.remove(AUTO_KEY);
        Office.context.document.settings.saveAsync((r) => {
          if (r.status !== Office.AsyncResultStatus.Succeeded) showError("自動オープンの設定を保存できませんでした: " + (r.error && r.error.message));
          fetch("/api/debug/log", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ autoOpen: on, saved: r.status }) }).catch(() => {});
        });
      } catch (e) {
        showError("自動オープンの設定に失敗しました: " + e.message);
      }
    });
  }
  /** 起動時に文字が選択されていて入力欄が空なら、その文字を取り込む(右クリック「この内容でスライドを作る」からの導線) */
  function importSelection() {
    if ($("prompt").value.trim()) return;
    try {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (r) => {
        if (r.status !== Office.AsyncResultStatus.Succeeded) return;
        const t = String(r.value || "").trim();
        if (t.length < 4) return;
        $("prompt").value = t;
        updateCompose();
        saveSettings();
        fetch("/api/debug/log", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ importedSelection: t.length }) }).catch(() => {});
      });
    } catch (_) {}
  }

  // ---------- 初期化 ----------
  Office.onReady(() => {
    setAccent("none");
    loadSettings();
    $("devJson").hidden = !devMode;
    refreshHealth();
    refreshReference();
    try {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged, () => {});
    } catch (_) {}
    $("go").addEventListener("click", () => generate());
    $("redraw").addEventListener("click", redraw);
    $("refStatus").addEventListener("click", refreshReference);
    $("prompt").addEventListener("input", () => {
      updateCompose();
      schedulePromptSave();
    });
    $("prompt").addEventListener("focus", () => document.body.classList.add("typing"));
    $("prompt").addEventListener("blur", updateCompose);
    $("expand").addEventListener("click", () => setZen(!document.body.classList.contains("zen")));
    $("fold").addEventListener("click", () => setEditing(!document.body.classList.contains("editing")));
    $("prompt").addEventListener("focus", () => {
      if (document.body.classList.contains("done")) setEditing(true);
    });
    $("pages").addEventListener("change", () => {
      pagesMode = Number($("pages").value) || 0;
      if (batch && !$("go").disabled) generate(); // 同じ入力なので差し替えになる
    });
    document.addEventListener("keydown", (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
        e.preventDefault();
        if (!$("go").disabled) generate();
      } else if (e.key === "Escape" && document.body.classList.contains("zen")) {
        e.preventDefault();
        setZen(false);
      }
    });
    document.querySelectorAll(".sw").forEach((b) =>
      b.addEventListener("click", () => {
        const v = b.dataset.accent;
        if (v === "custom") {
          const picker = $("accentPicker");
          picker.value = customAccent;
          picker.click();
          return;
        }
        setAccent(v);
        saveSettings();
      })
    );
    $("accentPicker").addEventListener("input", (e) => {
      setAccent(e.target.value);
      saveSettings();
    });
    $("accentPicker").addEventListener("change", (e) => {
      setAccent(e.target.value);
      saveSettings();
    });
    ["hint", "model", "layout", "mock"].forEach((id) =>
      $(id).addEventListener("change", () => {
        saveSettings();
        updateTuneState();
      })
    );
    $("hint").addEventListener("input", updateTuneState);
    updateCompose();
    updateTuneState();
    initAutoOpen();
    importSelection();
  });
})();
