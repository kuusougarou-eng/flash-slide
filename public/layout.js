/* eslint-disable */
/**
 * layout.js v4 — スライド仕様(spec) → 描画プリミティブ(rect / line / table / image)への純粋変換。
 * Office.js に依存しないので Node でも単体テスト可能。単位は pt。既定 960x540 (16:9)。
 *
 * ■ 思想(docs/feature-proposals/2026-09-06-composable-cell-grid-engine-plan.md)
 *  - 「型」を増やさない。テキストを持つ部品は cell 1 つ。見た目は属性(head/text/items/shape/fill/highlight/align)で決まる。
 *  - コンサル資料の大半はマトリクス。table は「セルを組み合わせた格子」であり、PPT ネイティブ表ではない
 *    (ネイティブ表は figure の一種 ntable として残す)。
 *  - 色・座標・pt は LLM に出させない。ここの STYLE トークンだけが決める。
 *  - 強調は 1 枚 2 箇所まで、薄い強調地(tint)は 1 つまで。プロンプトの「お願い」ではなくコードで強制する。
 *
 * ■ デザイン規律(consulting-pptx-skill の slide-rules を単スライド生成向けに要約)
 *  - 角丸禁止。塗りのあるボックスに枠線を付けない。枠線は塗りなし要素だけ。
 *  - 罫線は「その先に区切る対象がある場所」だけ。最終行の下・版面最下部には引かない。
 *  - 表(格子): ヘッダーは塗りなし・太字・下に太い強調色罫。行間は細い薄罫。ゼブラなし。縦罫なし。
 *  - 左=事実・分析、右=意味合い(So what)。下部の結論帯や浮遊ボックスは作らない。
 *  - 前提→帰結の強い因果は列間に塗り三角 1 つ。
 *  - 本文はタイトル下〜出典行上の 55% 以上を使う。少なければ垂直中央揃え。
 *  - 色はモノトーン + 強調色 1 色(accent:"none" で完全モノトーン)。
 *  - 本文 18pt 以上(格子・表のみ 14pt まで許容)。
 *
 * ■ spec
 * { title, lead?, footnote?, body: Node }
 * ■ Node
 *   { rows:[Node…], weights?, gap? } | { cols:[Node…], weights?, gap? }
 *   { type:"cell", head?, text?, items?:[string | string[]…], shape?:"rect"|"chevron"|"home", fill?:"none"|"light"|"dark"|"tint",
 *     highlight?, icon?, align?, size?:"large", level? }
 *   { type:"table", corner?, colHeaders:[…], colGroups?:[{text, span}], rowGroups?:[{text, span}],
 *     rows:[{head, cells:[string | {text?, items?:[string…], fill?, highlight?}], highlight?}],
 *     headShape?:"chevron", highlightCol?, axes?:{x, y} }
 *   { type:"bars"|"column"|"line"|"stacked"|"kpi", … }   … Figure(数値)
 *   { type:"ntable", … table と同じ }                    … Figure: PPT ネイティブ表(密な数表向け)
 *   { type:"arrow", direction?:"right"|"down" }
 * 旧型(label/card/bullets/text/callout/chevrons/icon)は normalizeSpec で cell/grid に脱糖される。
 * テキスト中の **太字** は Bold。
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.SlideLayout = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // =====================================================================
  //  デザイントークン(全描画はここだけを参照する)
  // =====================================================================
  const DEFAULT_PALETTE = {
    text: "#1A1A1A",
    textMuted: "#595959",
    textOnDark: "#FFFFFF",
    line: "#7F7F7F",
    lineLight: "#C8C8C8",
    fillLight: "#F2F2F2",
    fillMid: "#D9D9D9",
    fillDark: "#404040",
    accent: "#FD5108",
    accentText: "#FFFFFF",
  };

  const BASE_FONT = {
    title: 26,
    lead: 14,
    head: 16,
    body: 14,
    min: 14, // 本文の基本下限(参考デッキの密度に合わせる。テンプレがあればその本文サイズで上書き)
    floor: 11, // 溢れるときだけ許す最終下限(詳細は小さくてよい。情報を落とすより小さい方がよい)
    tableMin: 11, // 格子セルの最終下限(詳細)
    ntableMin: 10, // ネイティブ表セルの最終下限
    large: 18,
    kpi: 44,
    caption: 10, // 小さな見出しラベル(「何が起きるか」のような前置き)
    num: 28, // 番号(01 / 02 / 03)
    footnote: 9,
    lineHeight: 1.35, // 行間はデジタル庁デザインシステムに倣い広め
  };
  // 現在有効な文字サイズ(既定 = BASE_FONT)。参照テンプレートがあれば layout() が本文サイズを上書きする
  const FONT = Object.assign({}, BASE_FONT);
  /** 参照スライドの本文サイズに合わせて全体の文字階層を組み替える(テンプレ準拠) */
  function applyFontProfile(prof) {
    Object.assign(FONT, BASE_FONT);
    const b = prof && prof.bodyFontSize;
    if (!(b >= 9 && b <= 28)) return;
    FONT.body = b;
    FONT.min = b; // テンプレの本文サイズが基準。これを下回ったときだけ「縮小」扱い
    FONT.floor = Math.max(8, Math.round(b * 0.72));
    FONT.large = Math.round(b * 1.25);
    FONT.tableMin = Math.max(8, Math.round(b * 0.85));
    FONT.ntableMin = Math.max(8, Math.round(b * 0.8));
    FONT.head = Math.max(b, Math.round(prof.headFontSize || b * 1.15));
    FONT.kpi = Math.round(b * 2.6);
    FONT.caption = Math.max(8, Math.round(b * 0.75));
    FONT.num = Math.round(b * 2);
    FONT.lead = Math.max(FONT.lead, b);
    if (prof.footnoteFontSize >= 6) FONT.footnote = prof.footnoteFontSize;
  }
  const SPACE = { margin: 36, pad: 8, gap: 20, bulletIndent: 18, levelIndent: 16, cellPad: 6, chevronOverlap: 10 };
  const RULE = { thick: 1.8, thin: 0.75 };
  const BUDGET = { highlight: 2, tint: 1, leaves: 12, depth: 3 };
  const STYLE = { font: BASE_FONT, space: SPACE, rule: RULE, budget: BUDGET };

  const MARGIN = SPACE.margin;
  const PAD = SPACE.pad;
  const GAP = SPACE.gap;
  const BULLET_INDENT = SPACE.bulletIndent;
  const RULE_THICK = RULE.thick;
  const RULE_THIN = RULE.thin;
  const FILL_RATIO_MIN = 0.55;

  // ---------- 色 ----------
  function tint(hex, ratio) {
    const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex || "");
    if (!m) return "#FFF3EC";
    const f = (h) => Math.round(parseInt(h, 16) + (255 - parseInt(h, 16)) * ratio);
    return "#" + [f(m[1]), f(m[2]), f(m[3])].map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase();
  }

  // ---------- テキスト計測(概算) ----------
  function charWidthEm(ch) {
    const c = ch.codePointAt(0);
    if ((c >= 0x3000 && c <= 0x9fff) || (c >= 0xac00 && c <= 0xd7af) || (c >= 0xf900 && c <= 0xfaff) || (c >= 0xff00 && c <= 0xffef) || (c >= 0x20000 && c <= 0x2ffff)) return 1.0;
    if (c === 0x20) return 0.3;
    if (/[A-Z0-9]/.test(ch)) return 0.62;
    if (/[a-z]/.test(ch)) return 0.5;
    return 0.55;
  }
  // 段落の区切りは "\n"、段落内のソフト改行(Shift+Enter 相当)は "\v"。行数の見積もりはどちらも 1 行として数える
  const BREAK = /[\n\v]/;
  function stripBold(s) {
    return String(s || "").replace(/\*\*/g, "");
  }
  function measure(text, fs) {
    let w = 0;
    for (const ch of text) w += charWidthEm(ch) * fs;
    return w;
  }
  function estimateLines(text, width, fontSize) {
    if (!text) return 0;
    width = width * 0.95; // 実フォントは概算より広い(PowerPoint の折り返しが 1 行早く起きる)ための安全率
    let lines = 0;
    for (const p of stripBold(text).split(BREAK)) {
      let w = 0,
        l = 1;
      for (const ch of p) {
        const cw = charWidthEm(ch) * fontSize;
        if (w + cw > width) {
          l++;
          w = cw;
        } else w += cw;
      }
      lines += l;
    }
    return lines;
  }
  function textHeight(text, width, fontSize, paraGap) {
    const lines = estimateLines(text, Math.max(20, width), fontSize);
    const paras = text ? text.split("\n").length : 0;
    return lines * fontSize * FONT.lineHeight + (paraGap || 0) * Math.max(0, paras - 1);
  }
  /** 既存プレースホルダ用: 1 行で入るなら元のサイズ。入らなければ行数×1.15 で枠(+少しの余裕)に収まるまで縮小 */
  function fitSlot(text, w, h, pref, min) {
    let fs = pref;
    const fits = (f) => estimateLines(text, w - 8, f) * f * 1.15 <= h + 6;
    while (fs > min && !fits(fs)) fs -= 1;
    return Math.max(fs, min);
  }
  /** リード用: 1 行で入るなら元サイズ。入らなければ 1 行に収まるサイズまで(元の 72% を下限に)縮め、それでも無理なら 2 行で fitSlot */
  function fitLeadSlot(text, w, h, pref) {
    const oneLine = (f) => estimateLines(text, w - 8, f) <= 1;
    if (oneLine(pref)) return pref;
    for (let f = pref - 1; f >= Math.max(12, Math.round(pref * 0.72)); f--) if (oneLine(f)) return f;
    return fitSlot(text, w, h, pref, 12);
  }
  function fitFont(text, w, h, pref, min) {
    let fs = pref;
    while (fs > min && textHeight(text, w, fs) > h) fs -= 1;
    return Math.max(fs, min);
  }
  /**
   * 本文用: 18pt を基本に、収まらないときだけ floor まで縮小(警告)。
   * 逆に余白が大きいときは max(既定 22pt)まで拡大して版面を埋める(文字量に応じた適切な大きさ)。
   */
  function fitBody(c, text, w, h, pref, paraGap, max, list) {
    let fs = pref || FONT.body;
    while (fs > FONT.floor && textHeight(text, w, fs, paraGap) > h) fs -= 1;
    // 箇条書きを「1 項目 1 行」に収めるための縮小はしない。ぶら下げインデントが効かないという前提で
    // 入れていたが、実機で左揃えにして確かめたところ、折り返した 2 行目は点の位置に戻らず
    // 1 行目の文字位置に揃う(hint:"probeindent" の G)。折り返しは受け入れてよく、
    // 縮小はむしろ版面全体の文字を小さいまま止めていた。
    if (fs < FONT.min) c.warnings.push("shrunk to " + fs + "pt: " + stripBold(text).slice(0, 18));
    // 部品ごとに拡大しない(隣の部品と文字サイズが揃わなくなる)。版面が余るときの拡大は layout() が全体を一括で行う
    return fs;
  }
  function parseRuns(s) {
    s = String(s || "");
    let out = "",
      bold = false,
      start = 0,
      i = 0;
    const ranges = [];
    while (i < s.length) {
      if (s.startsWith("**", i)) {
        if (!bold) start = out.length;
        else if (out.length > start) ranges.push([start, out.length - start]);
        bold = !bold;
        i += 2;
      } else {
        out += s[i];
        i++;
      }
    }
    if (bold && out.length > start) ranges.push([start, out.length - start]);
    return { text: out, boldRanges: ranges };
  }
  const NUMERIC_RE = /^[\s\d.,%+\-−▲△◎○×－/件人円億万pt]+$/; // 「4 月」「2026 年」のような期間・日付は左揃えのまま(右揃えは量だけ)
  function isNumeric(t) {
    return NUMERIC_RE.test(t) && /\d/.test(t);
  }
  /** 上限で切り捨てたときに「何をいくつ落としたか」を警告に残す(黙って消さない) */
  function capList(arr, n, ctx, what) {
    const a = arr || [];
    if (a.length > n && ctx && ctx.warnings) ctx.warnings.push(`content dropped: ${what} ${a.length - n} 件(上限 ${n})`);
    return a.slice(0, n);
  }
  function isNA(t) {
    return /^[\s—\-–ー]*$/.test(t);
  }

  // ---------- プリミティブ ----------
  function rect(x, y, w, h, opt) {
    return Object.assign({ kind: "rect", x, y, w, h, fill: null, line: null, lineWeight: RULE_THIN, text: "", boldRanges: [], shape: "rect" }, opt || {});
  }
  function line(x1, y1, x2, y2, color, weight) {
    return { kind: "line", x1, y1, x2, y2, color, weight: weight || RULE_THIN };
  }
  function textBox(x, y, w, h, text, opt) {
    const runs = parseRuns(text);
    const o = Object.assign({ fontSize: FONT.body, bold: false, color: null, align: "left", valign: "top", bullets: false, pad: PAD, wrap: true, autofit: "shrink" }, opt || {});
    return rect(x, y, w, h, Object.assign({ text: runs.text, boldRanges: runs.boldRanges }, o));
  }
  function image(name, x, y, w, h, color) {
    return { kind: "image", name: String(name || "").toLowerCase(), x, y, w, h, color };
  }

  // =====================================================================
  //  レイアウト本体
  // =====================================================================
  function layout(spec, options) {
    options = options || {};
    const W = options.width || 960;
    const H = options.height || 540;
    const prof = options.profile || null;
    const P = Object.assign({}, DEFAULT_PALETTE, options.palette || {});
    const mono = String(P.accent || "").toLowerCase() === "none";
    if (mono) P.accent = P.fillDark;
    else if (prof && prof.accent && !(options.palette && options.palette.accent)) P.accent = prof.accent;
    P.accentTint = mono ? P.fillLight : tint(P.accent, 0.9);
    // 参照スライドの地が暗いときは、文字・罫・塗りを反転させる(白地前提のトークンのままだと読めない)
    if (prof && prof.darkBackground) {
      P.text = "#FFFFFF";
      P.textMuted = "#B8BEC8";
      P.textOnDark = "#14161A";
      P.line = "#6B7280";
      P.lineLight = "#4B5563";
      P.fillLight = "#2A2F37";
      P.fillMid = "#3C424C";
      P.fillDark = "#E8EAED";
      if (mono) {
        P.accent = P.fillDark;
        P.accentTint = P.fillLight;
      }
    }
    P.mono = mono;
    applyFontProfile(prof);
    if (spec.compositionVersion) {
      FONT.title = Math.max(FONT.title, 32);
      // 本文の下限はテンプレの本文サイズ(最低 14pt)。18pt に底上げすると密度が出ず、コンサル資料の見た目にならない
      FONT.body = Math.max(FONT.body, 14);
      FONT.min = FONT.body;
      FONT.head = Math.max(FONT.head, FONT.body + 4);
      FONT.large = FONT.body;
    }
    const prims = [];
    const warnings = [];
    const header = { titleText: "", leadText: "", footnoteText: "" };
    const bodyFont = (prof && prof.fontName) || null;

    const innerX = MARGIN;
    const innerW = W - MARGIN * 2;
    let y = MARGIN * 0.75;

    // --- タイトル(結論。1〜2 行、縮小して 1 行に詰めない) ---
    const title = stripBold(spec.title || "").trim();
    header.titleText = title;
    if (prof && prof.title) {
      const t = prof.title;
      // 再利用する枠はプレースホルダ自身の余白・自動調整があるので、行数ベースで緩く判定する(1 行で入るなら縮めない)
      header.titleFontSize = t.reuse ? fitSlot(title, t.w, t.h, t.fontSize || FONT.title, 18) : fitFont(title, t.w - PAD * 2, t.h - PAD * 2, t.fontSize || FONT.title, Math.min(t.fontSize || FONT.title, 18));
      if (!t.reuse) {
        const fs = fitFont(title, t.w - PAD * 2, t.h - PAD * 2, t.fontSize || FONT.title, Math.min(t.fontSize || FONT.title, 20));
        prims.push(textBox(t.x, t.y, t.w, t.h, title, { fontSize: fs, bold: t.bold !== false, color: t.color || P.text, align: t.align || "left", valign: t.valign || "middle", fontName: t.fontName || bodyFont, autofit: "none", role: "title" }));
      }
      y = t.y + t.h;
    } else {
      const kicker = stripBold(spec.kicker || "").trim();
      if (kicker) {
        const kh = FONT.caption * FONT.lineHeight + 4;
        prims.push(textBox(innerX, y, innerW, kh, kicker, { fontSize: FONT.caption, bold: true, color: P.textMuted, valign: "bottom", pad: 0, autofit: "none", role: "kicker" }));
        y += kh + 2;
      }
      const lines = Math.min(2, estimateLines(title, innerW - PAD * 2, FONT.title));
      const titleH = lines * FONT.title * FONT.lineHeight + PAD * 2;
      prims.push(textBox(innerX, y, innerW, titleH, title, { fontSize: FONT.title, bold: true, color: P.text, valign: "bottom", autofit: "none", role: "title" }));
      y += titleH + 4;
      prims.push(line(innerX, y, innerX + innerW, y, P.accent, RULE_THICK));
      y += 10;
    }

    // --- リード(任意。So what を 1〜2 文。テンプレートに枠が無ければ描かない) ---
    const lead = (spec.lead || "").trim();
    if (prof) {
      if (prof.lead && lead) {
        header.leadText = lead;
        const l = prof.lead;
        header.leadFontSize = l.reuse ? fitLeadSlot(lead, l.w, l.h, l.fontSize || FONT.lead) : fitFont(lead, l.w - PAD * 2, l.h - PAD * 2, l.fontSize || FONT.lead, Math.min(l.fontSize || FONT.lead, 12));
        if (!l.reuse) {
          const fs = fitFont(lead, l.w - PAD * 2, l.h - PAD * 2, l.fontSize || FONT.lead, Math.min(l.fontSize || FONT.lead, 14));
          prims.push(textBox(l.x, l.y, l.w, l.h, lead, { fontSize: fs, color: l.color || P.text, align: l.align || "left", valign: l.valign || "top", fontName: l.fontName || bodyFont, autofit: "none", role: "lead" }));
        }
        y = Math.max(y, l.y + l.h);
      }
    } else if (lead) {
      header.leadText = lead;
      const leadH = textHeight(lead, innerW - PAD * 2, FONT.lead) + PAD * 2;
      prims.push(textBox(innerX, y, innerW, leadH, lead, { fontSize: FONT.lead, color: P.textMuted, valign: "top", autofit: "none", role: "lead" }));
      y += leadH + 2;
    }

    // --- 出典(左下の定位置)と本文領域 ---
    const foot = (spec.footnote || "").trim();
    header.footnoteText = foot;
    let bodyRect;
    if (prof && prof.body) {
      bodyRect = { x: prof.body.x, y: prof.body.y, w: prof.body.w, h: prof.body.h };
      if (lead && !prof.lead) {
        // テンプレにリード枠が無い(参照スライドがリード無しで作られた等)。So what を落とさず、タイトル直下に自前で描いて版面を下げる
        header.leadText = "";
        const lx = bodyRect.x,
          lw = bodyRect.w;
        const fs = Math.max(FONT.lead, Math.min(FONT.body, 16));
        const leadH = textHeight(lead, lw - PAD * 2, fs) + PAD * 2;
        const ly = Math.max(y, (prof.title ? prof.title.y + prof.title.h : y) + 2);
        prims.push(textBox(lx, ly, lw, leadH, lead, { fontSize: fs, color: P.text, valign: "top", autofit: "none", fontName: bodyFont, role: "lead" }));
        const shift = ly + leadH + 6 - bodyRect.y;
        if (shift > 0) {
          bodyRect.y += shift;
          bodyRect.h -= shift;
        }
      }
      if (foot) {
        const f = prof.footnote;
        if (f && !f.reuse) prims.push(textBox(f.x, f.y, f.w, f.h, foot, { fontSize: f.fontSize || FONT.footnote, color: P.textMuted, valign: "bottom", pad: 0, autofit: "none", fontName: bodyFont, role: "footnote" }));
        else if (!f) {
          const fh = FONT.footnote * FONT.lineHeight + 4;
          bodyRect.h -= fh + 4;
          prims.push(textBox(bodyRect.x, bodyRect.y + bodyRect.h + 4, bodyRect.w, fh, foot, { fontSize: FONT.footnote, color: P.textMuted, valign: "bottom", pad: 0, autofit: "none", fontName: bodyFont, role: "footnote" }));
        }
      }
    } else {
      const footH = foot ? FONT.footnote * FONT.lineHeight + 4 : 0;
      const bodyTop = y + 8;
      const bodyBottom = H - MARGIN * 0.75 - (footH ? footH + 6 : 0);
      bodyRect = { x: innerX, y: bodyTop, w: innerW, h: bodyBottom - bodyTop };
      if (foot) prims.push(textBox(innerX, H - MARGIN * 0.75 - footH, innerW, footH, foot, { fontSize: FONT.footnote, color: P.textMuted, valign: "bottom", pad: 0, autofit: "none", role: "footnote" }));
    }
    if (bodyRect.h < 120) warnings.push("body area too small: " + Math.round(bodyRect.h) + "pt");

    // --- 本文(グリッド) ---
    const body = spec.body && (spec.body.rows || spec.body.cols || spec.body.type) ? spec.body : { type: "cell", text: "" };
    /** 本文を 1 回描く。bump>0 なら文字階層をその分大きくして描く(版面充填) */
    const renderBodyPass = (bump) => {
      const saved = Object.assign({}, FONT);
      if (bump) {
        FONT.body += bump;
        FONT.min += bump;
        FONT.head = bump > 0 ? Math.max(FONT.body + 4, FONT.head + bump) : Math.max(FONT.body, FONT.head + bump);
        FONT.large += bump;
        FONT.num += Math.round(bump * 1.5);
        FONT.kpi += bump * 2;
        FONT.tableMin = Math.max(8, FONT.tableMin + Math.round(bump / 2));
      }
      const pr = [],
        wr = [];
      const ctx = { P, prims: pr, warnings: wr, leafCount: 0, fontName: bodyFont, bodyStart: 0 };
      try {
        if (spec.compositionVersion) renderComposition(spec, bodyRect, ctx);
        else renderNode(body, bodyRect, ctx, "rows", true);
      } catch (e) {
        wr.push("body layout failed: " + (e && e.message));
        LEAF.cell({ type: "cell", text: "(本文の描画に失敗しました)" }, bodyRect, ctx);
      }
      const fonts = { body: FONT.body, head: FONT.head, min: FONT.min, floor: FONT.floor };
      Object.assign(FONT, saved);
      let bottom = bodyRect.y,
        alloc = bodyRect.y;
      const bands = []; // 実際に何かが見えている縦の区間(重なりは後で畳む)
      for (const p of pr) {
        if (p.kind === "line") {
          bottom = Math.max(bottom, p.y1, p.y2);
          bands.push([Math.min(p.y1, p.y2), Math.max(p.y1, p.y2)]);
        } else if (p.y != null && p.h != null) {
          alloc = Math.max(alloc, p.y + p.h);
          // 見える高さ: 塗りのある箱・図形・画像は枠、上寄せの文字は文字の高さ(伸ばした空の枠は数えない)
          // 塗りのある箱・図形・画像は枠を、塗りのない文字箱は文字の高さを数える。
          // (縦中央に置いた文字を枠の高さで数えると、余白だらけの版面が「満杯」に見えて拡大が止まる)
          const solid = p.kind === "image" || (p.fill && p.fill !== "none") || (p.shape && p.shape !== "rect") || !p.text;
          const visible = solid ? p.h : Math.min(p.h, textHeight(p.text, p.w - (p.pad || 0) * 2, p.fontSize || FONT.body, 0) + (p.pad || 0) * 2);
          bottom = Math.max(bottom, p.y + visible);
          bands.push([p.y, p.y + visible]);
        }
      }
      // 充填率は「一番下の要素の位置」ではなく「実際に見えている縦の区間の合計」で測る。
      // 下端だけで測ると、少ない項目を大きな行間でばら撒いた版面が「埋まっている」と誤判定され、
      // 文字を大きくする段(下の版面充填)が発動しない(実測: 4 項目のパネルで下端 0.85 / 実面積 0.50)。
      bands.sort((a, b) => a[0] - b[0]);
      let covered = 0,
        curFrom = null,
        curTo = null;
      for (const [a, b] of bands) {
        const lo = Math.max(a, bodyRect.y),
          hi = Math.min(b, bodyRect.y + bodyRect.h);
        if (hi <= lo) continue;
        if (curTo == null || lo > curTo) {
          if (curTo != null) covered += curTo - curFrom;
          curFrom = lo;
          curTo = hi;
        } else curTo = Math.max(curTo, hi);
      }
      if (curTo != null) covered += curTo - curFrom;
      const fill = covered / Math.max(1, bodyRect.h);
      // 「悪い」= 実際に文字が縮んだ / 格子・ガントが溢れた / 割付が版面を越えた / 行の自然高さが 5% 超過(数 pt の超過は比例配分で吸収されるので許す)
      const rowsOver = wr.some((w) => {
        const m = w.match(/^rows overflow: natural (\d+)pt > (\d+)pt/);
        return m && Number(m[1]) > Number(m[2]) * 1.05;
      });
      // ガントは描画側で行数に応じて自前で文字を縮めるので、その警告では全体を縮めない
      const bad = rowsOver || wr.some((w) => /shrunk|too small|too wide|failed|table overflow|sequence overflow|tree overflow|note too large/.test(w)) || alloc > bodyRect.y + bodyRect.h + 1;
      return { prims: pr, warnings: wr, fonts, fill, bad, bump };
    };
    let pass = renderBodyPass(0);
    if (pass.bad && !options.noGrow) {
      // 溢れる・縮む: 部品ごとに縮めず、スライド全体の文字階層を 1pt ずつ下げて揃える(下限 floor)。それでも収まらなければ既定のまま(部品側の縮小に任せる)
      for (let b = -1; b >= -4; b--) {
        if (FONT.body + b < FONT.floor) break;
        const next = renderBodyPass(b);
        if (!next.bad) {
          pass = next;
          break;
        }
      }
    }
    // 版面充填: 文字量が版面に対して少ないときは本文を段階的に大きくして埋める(+2pt ずつ、本文 20pt まで)。溢れたら手前の段に戻す
    // 拡大するのは版面が余っているときだけ(composition でも同じ)。密なスライドを無条件に 20pt へ広げない
    // 入口の閾値は 0.85。0.70 だと、塊を離す空行の分だけ充填率が上がったスライドで拡大が止まり、
    // 本文が 14pt のまま余白の多い版面になっていた。溢れるかどうかは下のループが実際に描いて確かめる
    if (!pass.bad && pass.fill < 0.85 && !options.noGrow) {
      for (let b = 2; b <= 6; b += 2) {
        if (FONT.body + b > 20) break; // 上限 20pt(24pt は疎なスライドで幼く見える)
        const next = renderBodyPass(b);
        if (next.bad || (!spec.compositionVersion && next.fill > 0.98)) break;
        pass = next;
        if (!spec.compositionVersion && pass.fill >= 0.8) break;
      }
    }
    const bodyStart = prims.length;
    pass.prims.forEach((p) => prims.push(p));
    pass.warnings.forEach((w) => warnings.push(w));
    equalizeMatrixFonts(prims);
    equalizeCellFonts(prims);
    for (let i = bodyStart; i < prims.length; i++) {
      const p = prims[i];
      p.body = true;
      if (bodyFont && p.kind === "rect" && p.text && !p.fontName) p.fontName = bodyFont;
    }
    const fonts = pass.fonts;
    if (pass.bump) fonts.grown = pass.bump;
    applyFontProfile(null); // 既定に戻す(次の呼び出しに引きずらない)
    return { width: W, height: H, prims, warnings, header, bodyRect, fonts };
  }

  /** 同じ容器(rows / cols)に属する本文セルは同じ文字サイズに揃える。縮小方向なので溢れない */
  function equalizeCellFonts(prims) {
    const groups = new Map();
    for (const p of prims) {
      if (!p.gid || p.kind !== "rect" || !p.text || p.role) continue;
      const g = groups.get(p.gid) || [];
      g.push(p);
      groups.set(p.gid, g);
    }
    for (const g of groups.values()) {
      if (g.length < 2) continue;
      const fs = Math.min.apply(null, g.map((p) => p.fontSize));
      g.forEach((p) => {
        p.fontSize = fs;
        if (fs < FONT.min) p.shrunk = true; // 基準サイズを下回ったことは警告として残す
      });
    }
  }

  /** 横に並ぶ(縦位置が重なる)格子は同じ文字サイズに揃える。縮小方向なので溢れは起きない */
  function equalizeMatrixFonts(prims) {
    const groups = new Map();
    for (const p of prims) {
      if (p.role !== "matrixcell" || !p.mid) continue;
      const g = groups.get(p.mid) || { fs: Infinity, top: Infinity, bottom: -Infinity, items: [] };
      g.fs = Math.min(g.fs, p.fontSize);
      g.top = Math.min(g.top, p.y);
      g.bottom = Math.max(g.bottom, p.y + p.h);
      g.items.push(p);
      groups.set(p.mid, g);
    }
    const list = Array.from(groups.values());
    if (list.length < 2) return;
    for (let i = 0; i < list.length; i++) {
      for (let j = i + 1; j < list.length; j++) {
        const a = list[i], b = list[j];
        const overlap = Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top);
        if (overlap <= Math.min(a.bottom - a.top, b.bottom - b.top) * 0.5) continue; // 縦にずれた格子は別扱い
        const fs = Math.min(a.fs, b.fs);
        a.fs = b.fs = fs;
        a.items.concat(b.items).forEach((p) => (p.fontSize = Math.min(p.fontSize, fs)));
      }
    }
  }

  // =====================================================================
  //  コンテナ(grid)
  // =====================================================================
  const isContainer = (node) => !node.type && ((node.rows && node.rows.length) || (node.cols && node.cols.length));
  /** 見出しだけのセル(パネルの見出し) */
  const isHeadingCell = (n) => !!n && n.type === "cell" && !!(n.head && n.head.trim()) && !n.shape && !(n.text && n.text.trim()) && !(n.items && n.items.length);
  function isFixedCell(node) {
    // 見出しだけ / 矢羽 / アイコンだけ のセルは自然高さ固定(伸ばさない)
    if (node.type !== "cell") return false;
    if (node.style === "takeaway") return true;
    const hasBody = (node.text && node.text.trim()) || (node.items && node.items.length);
    return !hasBody || !!node.shape;
  }
  function isStretch(node) {
    if (isContainer(node)) return (node.rows || node.cols).some(isStretch);
    if (isBoxed(node)) return false;
    if (node.type === "cell") return !isFixedCell(node);
    return !["arrow", "kpi", "org", "tree"].includes(node.type);
  }
  const isStretchLeaf = (node) => isStretch(node) && !isContainer(node);
  /** 塗りのある箱(fill / step / takeaway / 濃い kpi)。高さは中身に合わせ、列いっぱいには伸ばさない */
  const isBoxed = (node) => !!node && ((node.type === "cell" && ((node.fill && node.fill !== "none") || node.shape === "step" || node.style === "takeaway")) || (node.type === "kpi" && (node.fill === "dark" || node.fill === "accent")));
  /** 箱の高さの上限。塗りの箱は中身の 1.6 倍 + 余白まで、示唆帯は 1.25 倍まで(黒い塊で版面を潰さない) */
  function maxBoxH(node, nat) {
    if (node.style === "takeaway") return nat * 1.25 + 6;
    if (node.shape === "step") return nat * 1.6;
    if (node.fill === "dark" || node.fill === "accent") return nat * 1.3 + 16;
    return nat * 1.6 + 24;
  }
  /** rows の中で伸ばしてよい上限(箱は maxBoxH、箱だけの容器は 1.5 倍、それ以外は無制限) */
  function stretchMaxOf(node, nat) {
    if (isBoxed(node)) return maxBoxH(node, nat);
    if (isContainer(node)) {
      const kids = node.rows || node.cols;
      if (kids.every(isBoxed)) return nat * 1.5;
      if (kids.some(hasFillingLeaf) || kids.some((k) => k && k.type === "cell" && k.head && String(k.head).trim())) return Infinity; // 下線パネルは列の高さを使う
      return nat * 1.2 + 12;
    }
    if (node.type === "cell") return (node.head && String(node.head).trim()) || node._panelBody ? Infinity : nat * 1.2 + 12; // 見出し無しの文章の箱だけ上に詰める
    return Infinity;
  }
  /** 領域いっぱいに使ってよい部品(格子・図・ガント・体制図)を含むか */
  function hasFillingLeaf(node) {
    if (!node) return false;
    if (isContainer(node)) return (node.rows || node.cols).some(hasFillingLeaf);
    return !["cell", "arrow", "kpi"].includes(node.type);
  }
  function defaultColWeight(node) {
    if (node.type === "arrow") return 0.28;
    if (node.type === "kpi") return 0.8;
    if (node.type === "cell" && node.icon && !node.head && !node.text && !(node.items && node.items.length)) return 0.5;
    return 1;
  }
  function gapOf(node) {
    if (typeof node.gap === "number") return node.gap;
    if (node.cols && node.cols.length >= 2 && node.cols.every((k) => k.type === "cell" && k.shape === "step")) return 26; // 箱フローは矢印の分だけ広く
    return GAP;
  }
  function renderNode(node, r, ctx, orientation, isRoot) {
    // 注意: table ノードも .rows を持つので、コンテナ判定は type が無い場合に限る
    if (isContainer(node) && node.rows) return renderRows(node, r, ctx, isRoot);
    if (isContainer(node) && node.cols) return renderCols(node, r, ctx, isRoot);
    const fn = LEAF[node.type] || LEAF.cell;
    ctx.leafCount++;
    node._orientation = orientation;
    if (isRoot && !isStretchLeaf(node)) {
      // 単独の固定ブロック(箱・矢羽など)は上端に置き、高さは中身に合わせる(中央に浮かせない)
      const nat = naturalHeight(node, r.w);
      const h = Math.min(r.h, isBoxed(node) ? maxBoxH(node, nat) : Math.max(nat, r.h * FILL_RATIO_MIN));
      return fn(node, { x: r.x, y: r.y, w: r.w, h }, ctx);
    }
    fn(node, r, ctx);
  }

  /** 自然高さ nat[] と固定フラグ fixed[] から、avail に収まる行高さを配分する */
  function minHeightOf(node) {
    if (!node || isContainer(node)) return 0;
    if (node.type === "table" || node.type === "ntable") return 40 + Math.min(8, (node.rows || []).length) * 28;
    if (node.type === "cell" && node.shape) return 40;
    return 0;
  }
  function distributeHeights(nat, fixed, avail, ctx, kids) {
    const maxs = kids ? kids.map((k, i) => stretchMaxOf(k, nat[i])) : null;
    const out = distributeHeightsRaw(nat, fixed, avail, ctx, maxs);
    if (!kids) return out;
    // 最低高さの保証: 不足分は最も高い伸縮要素から差し引く
    for (let i = 0; i < out.length; i++) {
      const min = minHeightOf(kids[i]);
      if (min && out[i] < min) {
        let need = min - out[i];
        out[i] = min;
        const order = out.map((h, j) => j).filter((j) => j !== i && !fixed[j]).sort((a, b) => out[b] - out[a]);
        for (const j of order) {
          const give = Math.min(need, Math.max(0, out[j] - 40));
          out[j] -= give;
          need -= give;
          if (need <= 0) break;
        }
      }
    }
    return out;
  }
  function distributeHeightsRaw(nat, fixed, avail, ctx, maxs) {
    const natSum = nat.reduce((a, b) => a + b, 0);
    if (natSum <= avail) {
      // 余りは伸縮要素に比例配分。上限のある要素(箱)は上限で止め、残りを他へ回す。全部が上限に達したら余りは下の余白にする
      const out = nat.slice();
      const open = nat.map((h, i) => !fixed[i] || (maxs && isFinite(maxs[i]) && maxs[i] > h));
      let extra = avail - natSum;
      for (let iter = 0; iter < 4 && extra > 0.5; iter++) {
        const sum = nat.reduce((a, b, i) => a + (open[i] ? b : 0), 0);
        if (!sum) break;
        let used = 0;
        nat.forEach((h, i) => {
          if (!open[i]) return;
          let add = (extra * h) / sum;
          const cap = maxs ? maxs[i] : Infinity;
          if (out[i] + add > cap) {
            add = Math.max(0, cap - out[i]);
            open[i] = false;
          }
          out[i] += add;
          used += add;
        });
        extra -= used;
      }
      return out;
    }
    // 収まらない: 固定行は据え置き、伸縮行に残りを配分。残りが乏しければ全体を比例縮小(領域外には出さない)
    const fixedSum = nat.reduce((a, b, i) => a + (fixed[i] ? b : 0), 0);
    const stretchSum = natSum - fixedSum || 1;
    ctx.warnings.push("rows overflow: natural " + Math.round(natSum) + "pt > " + Math.round(avail) + "pt");
    if (fixedSum < avail * 0.6) {
      const room = avail - fixedSum;
      return nat.map((h, i) => (fixed[i] ? h : (h * room) / stretchSum));
    }
    return nat.map((h) => (h * avail) / natSum);
  }

  function renderRows(node, r, ctx, isRoot, presetHeights, firstHeight) {
    const kids = node.rows;
    const gid = (ctx.gidSeq = (ctx.gidSeq || 0) + 1);
    const inner = node._inPanel || isHeadingCell(kids[0]);
    kids.forEach((k, i) => {
      k._gid = gid;
      k._inPanel = node._inPanel || (inner && i > 0); // 見出し罫の下にある要素は「パネルの中」
      k._panelBody = i > 0 && isHeadingCell(kids[i - 1]) && k.type === "cell" && !k.head; // 見出しだけの cell の直下の本文
    });
    const n = kids.length;
    const gap = gapOf(node);
    const avail = r.h - gap * (n - 1);
    let heights;
    const anyBoxed = kids.some((k) => isBoxed(k) || (isContainer(k) && (k.rows || k.cols).every(isBoxed)));
    if (presetHeights) heights = presetHeights;
    else if (firstHeight && n >= 2) {
      // 兄弟列と見出し行(1 段目)の高さを共有し、残りをこの列の中で配分する(見出し罫と本文の上端が列をまたいで揃う)
      const rest = distributeHeights(kids.slice(1).map((k) => naturalHeight(k, r.w)), kids.slice(1).map((k) => !isStretch(k)), avail - firstHeight, ctx, kids.slice(1));
      heights = [firstHeight].concat(rest);
    }
    else if (!anyBoxed && kids.some(hasFillingLeaf) && Array.isArray(node.weights) && node.weights.length === n && node.weights.every((w) => w > 0)) {
      const sum = node.weights.reduce((a, b) => a + b, 0);
      heights = node.weights.map((w) => (avail * w) / sum);
    } else {
      heights = distributeHeights(
        kids.map((k) => naturalHeight(k, r.w)),
        kids.map((k) => !isStretch(k)),
        avail,
        ctx,
        kids
      );
    }
    let y = r.y;
    kids.forEach((k, i) => {
      renderNode(k, { x: r.x, y, w: r.w, h: heights[i] }, ctx, "rows", false);
      y += heights[i] + gap;
    });
  }

  function renderCols(node, r, ctx, isRoot) {
    const kids = node.cols;
    const gid = (ctx.gidSeq = (ctx.gidSeq || 0) + 1);
    kids.forEach((k) => {
      k._gid = gid;
      k._inPanel = node._inPanel;
      // 横並びの各パネルで、先頭が格子ならその列見出し行をパネル見出しとして扱う(隣の見出し罫と同じ高さに太罫が並ぶ)
      if (!node._inPanel && kids.length >= 2) {
        const first = isContainer(k) && k.rows ? k.rows[0] : k;
        if (first && (first.type === "table" || first.type === "ntable")) first._panelHead = true;
      }
    });
    // 兄弟パネルが箇条書きのセルなら、行の高さを共有して行の罫を同じ y に揃える(段数が違っても)
    if (kids.length >= 2) {
      const listKids = kids.filter((k) => k.type === "cell" && Array.isArray(k.items) && k.items.length);
      if (listKids.length >= 2) {
        const rowsOf = (k) => k.items.reduce((acc, v) => acc + (Array.isArray(v) ? v.length : 1), 0);
        const maxRows = Math.max(...listKids.map(rowsOf));
        listKids.forEach((k) => (k._siblingRows = maxRows));
      }
    }
    const n = kids.length;
    const gap = gapOf(node);
    const avail = r.w - gap * (n - 1);
    // 下線付きボックス・文章の列は同じ幅で並べる(方法論)。weights に従うのは図・格子・矢印を含む列だけ
    const panelLike = kids.every((k) => (k.type === "cell" && !k.shape) || (isContainer(k) && !!k.rows && !hasFillingLeaf(k)));
    const weights = !panelLike && Array.isArray(node.weights) && node.weights.length === n && node.weights.every((w) => w > 0) ? node.weights : kids.map(defaultColWeight);
    const sum = weights.reduce((a, b) => a + b, 0);
    const colW = weights.map((w) => (avail * w) / sum);
    // 上端揃え(中央に浮かせない)。左右の列は同じ高さの枠を受け取り、上端・見出し罫が揃う
    const ry = r.y,
      rh = r.h;
    // 兄弟列の行揃え: 全列が同じ段数の rows なら、段ごとの高さを列間で共有する(見出しの下端・本文の上端が揃う)
    let shared = null;
    const rowKids = kids.filter((k) => k.type !== "arrow");
    if (rowKids.length >= 2 && rowKids.every((k) => isContainer(k) && k.rows && !k.weights && k.rows.length === rowKids[0].rows.length && k.rows.length >= 2)) {
      const m = rowKids[0].rows.length;
      const nat = [],
        fixed = [];
      for (let i = 0; i < m; i++) {
        let mx = 0,
          fx = true;
        kids.forEach((k, ci) => {
          if (k.type === "arrow") return;
          mx = Math.max(mx, naturalHeight(k.rows[i], colW[ci]));
          if (isStretch(k.rows[i])) fx = false;
        });
        nat.push(mx);
        fixed.push(fx);
      }
      const g = gapOf(rowKids[0]);
      const maxs = nat.map((h, i) => Math.max.apply(null, rowKids.map((k) => stretchMaxOf(k.rows[i], h))));
      shared = distributeHeightsRaw(nat, fixed, rh - g * (m - 1), ctx, maxs);
    }
    // 情報量が多い文章主体の列は、横に割らず縦に積んで横長に読ませる(狭い列に押し込むと行数が増えて読みにくい)
    if (!isRoot || true) {
      const textOnly = (k) => {
        if (k.type === "cell") return !k.shape;
        if (isContainer(k)) return (k.rows || k.cols).every(textOnly);
        return false;
      };
      // 1 行に入る文字数が少ない(= 列が文章に対して狭い)か、列に収まらない量なら縦積みにする
      const textLen = (k) => JSON.stringify(k).replace(/[^぀-ヿ一-鿿A-Za-z0-9]/g, "").length;
      const narrow = kids.some((k, i) => (colW[i] - PAD * 2) / FONT.body < 22 && textLen(k) > 110);
      const colsRatio = Math.max.apply(null, kids.map((k, i) => naturalHeight(k, colW[i]))) / Math.max(1, r.h);
      const stackRatio = (kids.reduce((a, k) => a + naturalHeight(k, r.w), 0) + GAP * (kids.length - 1)) / Math.max(1, r.h);
      // 縦積みにするのは「縦積みなら版面に収まる」か「横割りより縦積みの方が溢れが小さい」ときだけ(判定は版面全体の高さで行う)
      const dense = kids.length >= 2 && kids.every(textOnly) && (narrow || colsRatio > 1) && (stackRatio <= 1 || (colsRatio > 1 && stackRatio < colsRatio));
      if (dense) {
        const t = enumToTable(kids); // 行名の箱 + 横長の説明(方法論)。組めない形なら縦積み
        if (t) return LEAF.table(t, r, ctx);
        return renderRows({ rows: kids }, r, ctx, isRoot);
      }
    }

    // 段数が違う列同士でも、1 段目が見出しだけの cell なら見出し行の高さを揃える(グリッド整合)
    let firstH = 0;
    if (!shared) {
      const heads = kids.filter((k) => isContainer(k) && k.rows && k.rows.length >= 2 && k.rows[0].type === "cell" && isFixedCell(k.rows[0]) && !k.rows[0].shape);
      if (heads.length >= 2) firstH = Math.max.apply(null, kids.map((k, i) => (heads.includes(k) ? naturalHeight(k.rows[0], colW[i]) : 0)));
    }
    let x = r.x;
    const stepFlow = kids.length >= 2 && kids.every((k) => k.type === "cell" && k.shape === "step");
    // 箱(塗り)の高さ: 兄弟の自然高さの最大に揃える(番号カードは同じ高さ、黒い箱は中身の高さ)。列いっぱいには伸ばさない
    const natMax = Math.max.apply(null, kids.map((k, i) => naturalHeight(k, colW[i])));
    // 箱フローは行の高さ(自然高さの 1.6 倍まで)に揃える。それ以外の箱は兄弟の自然高さの最大に揃え、自分の上限(maxBoxH)を超えない
    const boxH = stepFlow ? Math.min(rh, natMax * 1.6) : kids.some(isBoxed) ? Math.min(rh, natMax * 1.08 + 8) : rh;
    kids.forEach((k, i) => {
      const w = colW[i];
      const kh = isBoxed(k) ? (stepFlow ? boxH : Math.min(boxH, maxBoxH(k, naturalHeight(k, w)))) : rh;
      if (stepFlow && i > 0) {
        // 箱と箱の間に小さな三角(進行方向)
        const d = Math.min(12, gap - 6);
        const bandH = kids.some((q) => q.text || (q.items && q.items.length)) ? Math.max(FONT.head * FONT.lineHeight + 8, boxH * 0.42) : boxH;
        c_pushTri(ctx, x - gap / 2 - d / 2, ry + bandH / 2 - d / 2, d);
      }
      if (shared && isContainer(k) && k.rows) renderRows(k, { x, y: ry, w, h: rh }, ctx, false, shared);
      else if (firstH && isContainer(k) && k.rows && k.rows.length >= 2 && k.rows[0].type === "cell" && isFixedCell(k.rows[0]) && !k.rows[0].shape) renderRows(k, { x, y: ry, w, h: rh }, ctx, false, null, firstH);
      else renderNode(k, { x, y: ry, w, h: kh }, ctx, "cols", false);
      x += w + gap;
    });
  }

  /**
   * 並列の事象(見出し + 説明)の列を「行名の箱 + 横長の説明」の格子に組み替える。
   * 方法論: 事象が 3〜6 と多く説明が長いときは横に割らず縦に並べ、横長の列で読ませる
   */
  function enumToTable(kids) {
    const pick = (k) => {
      if (!k) return null;
      if (k.type === "cell" && !k.shape && !k.style && k.head && String(k.head).trim() && ((k.text && k.text.trim()) || (k.items && k.items.length)))
        return { head: k.head, text: k.text, items: k.items, highlight: !!k.highlight, num: k.num };
      if (isContainer(k) && k.rows && k.rows.length === 2 && isHeadingCell(k.rows[0]) && k.rows[1].type === "cell" && !k.rows[1].shape && !k.rows[1].head)
        return { head: k.rows[0].head, text: k.rows[1].text, items: k.rows[1].items, highlight: !!(k.rows[0].highlight || k.rows[1].highlight) };
      return null;
    };
    const items = kids.map(pick);
    if (items.length < 3 || items.some((x) => !x)) return null;
    const flat = (arr) => (arr || []).flatMap((it) => (Array.isArray(it) ? it : [it])).map(String);
    return {
      type: "table",
      colHeaders: [],
      numbered: items.every((x) => x.num),
      rows: items.map((x) => ({ head: stripBold(String(x.head)), cells: [x.items && x.items.length ? flat(x.items).join("\n") : String(x.text || "")], highlight: x.highlight })),
    };
  }
  /** 方法論の適用(正規化の最後): 説明の長い並列事象の横並びは、縦並びの格子に組み替える */
  function applyMethodology(body, W) {
    const isStepRow = (k) => isContainer(k) && !!k.cols && k.cols.every((c) => c && c.type === "cell" && c.shape === "step");
    const walk = (n) => {
      if (!n || typeof n !== "object" || !isContainer(n)) return;
      const kids = n.rows || n.cols;
      // 箱フローが rows[cols[step,step], cols[step,step]] のように 2 段に割られていたら、5 段以下なら 1 列の流れに戻す(左→右の順序を切らない)
      if (n.rows && kids.length >= 2 && kids.every(isStepRow)) {
        const steps = kids.flatMap((k) => k.cols);
        if (steps.length <= 5) {
          delete n.rows;
          delete n.weights;
          n.cols = steps;
          return;
        }
      }
      if (n.cols && kids.length >= 3) {
        const t = enumToTable(kids);
        if (t) {
          const colW = ((W || 888) - GAP * (kids.length - 1)) / kids.length;
          const cpl = Math.max(8, Math.floor((colW - PAD * 2) / FONT.body));
          const linesOf = (rw) => rw.cells[0].split("\n").reduce((a, ln) => a + Math.max(1, Math.ceil(stripBold(ln).length / cpl)), 0);
          const maxLines = Math.max.apply(null, t.rows.map(linesOf));
          if (kids.length >= 5 || maxLines > 4) {
            delete n.cols;
            delete n.weights;
            Object.assign(n, t);
            return;
          }
        }
      }
      kids.forEach(walk);
    };
    walk(body);
  }

  function c_pushTri(ctx, x, y, d) {
    ctx.prims.push(rect(x, y, d, d, { shape: "triangle", rotation: 90, fill: ctx.P.fillMid }));
  }
  /** ブロックの自然高さ(pt) */
  function naturalHeight(node, w) {
    if (isContainer(node) && node.rows) return node.rows.reduce((a, k) => a + naturalHeight(k, w), 0) + gapOf(node) * (node.rows.length - 1);
    if (isContainer(node) && node.cols) {
      const ws = node.cols.map(defaultColWeight);
      const sum = ws.reduce((a, b) => a + b, 0);
      const avail = w - gapOf(node) * (node.cols.length - 1);
      return Math.max.apply(null, node.cols.map((k, i) => naturalHeight(k, (avail * ws[i]) / sum)));
    }
    const inner = w - PAD * 2;
    switch (node.type) {
      case "cell":
        return cellNaturalHeight(node, w);
      case "table":
        return matrixMetrics(node, w, null).natural;
      case "gantt":
        return ganttMetrics(node, w).natural;
      case "org": {
        const d = node.root ? orgDepth(node.root) : 1;
        return d * orgBoxH() + (d - 1) * 24;
      }
      case "tree": {
        // 横に伸びる樹形図は、縦の必要量が葉の数で決まる(深さではない)
        const lv = node.root ? orgLeaves(node.root) : 1;
        return lv * Math.max(34, FONT.body * 2.2) + (lv - 1) * 8;
      }
      case "ntable": {
        const rows = node.rows || [];
        const cols = (node.colHeaders || []).length;
        const cellW = cols ? (w - Math.min(w * 0.26, 220)) / cols - 12 : inner;
        let h = headerHeight();
        for (const rw of rows) h += Math.max.apply(null, [40].concat((rw.cells || []).map((c) => textHeight(cellText(c), cellW, FONT.body) + 12)));
        return h;
      }
      case "bars":
        return Math.max(1, (node.items || []).length) * 44 + (node.note ? 26 : 0);
      case "column":
        return 220;
      case "line":
        return 220;
      case "stacked":
        return 40 + Math.max(1, (node.items || []).length) * 50;
      case "kpi":
        return 92;
      case "arrow":
        return 40;
      default:
        return 60;
    }
  }
  function headerHeight() {
    return FONT.head * FONT.lineHeight + PAD + 6;
  }
  /** 見出しの高さ: 1 行で入らなければ 2 行(18pt)まで許容。文字縮小で 1 行に詰めない(slide-rules) */
  function headingMetrics(text, w, hasIcon) {
    const iconW = hasIcon ? FONT.head * FONT.lineHeight : 0;
    const avail = w - iconW - PAD * 2;
    const t = stripBold(text || "");
    if (estimateLines(t, avail, FONT.head) <= 1) return { h: headerHeight(), fs: FONT.head, lines: 1 };
    const fs = 18;
    const lines = Math.min(2, estimateLines(t, avail, fs));
    return { h: fs * FONT.lineHeight * lines + PAD + 6, fs, lines };
  }

  // ---------- テキスト整形 ----------
  function itemToText(it) {
    if (typeof it === "string") return it;
    const head = stripBold(it.head || "").trim();
    const text = (it.text || "").trim();
    if (head && text) return "**" + head + "**" + (/[:：]$/.test(head) ? " " : ":") + text;
    if (head) return "**" + head + "**";
    return text;
  }
  /** LLM が {"label":"受注","value":"前年比 **+18%**"} のように構造で返したときの正規化。
   *  ラベルは太字の見出しとして残す(描画側はこの形からラベル列を揃えるかどうかを決める) */
  function labelValueText(v) {
    if (!v || typeof v !== "object" || Array.isArray(v)) return v;
    if (v.label == null && v.value == null) return v;
    const label = stripBold(str(v.label != null ? v.label : v.head)).trim();
    const value = str(v.value != null ? v.value : v.text).trim();
    if (!label) return value;
    return value ? "**" + label + "**: " + value : "**" + label + "**";
  }
  /** items → 本文テキスト。ネスト配列は第 2 階層(先頭に記号を付けて表現し、ネイティブ箇条書きは使わない) */
  function itemsBody(items) {
    const arr = items || [];
    const nested = arr.some(Array.isArray);
    if (!nested) return { text: arr.map(itemToText).join("\n"), list: arr.length > 1, nested: false };
    // 第 2 階層はソフト改行("\v" = Shift+Enter 相当)で親の段落の中に置く。段落が変わらないので
    // 子には点が付かず、親の文字の左端(ぶら下げ位置)に自動で揃う。実機で確認済み(hint:"probeindent" の E)。
    // 段落を 1 段下げる手段は無い(indentLevel は何も起きず、タブ文字は点だけ左端に置き去りになる)ので、
    // 階層はこの形でしか作れない。点そのものは PowerPoint 本来の箇条書きに描かせる
    const lines = [];
    // 子は 1 段深く字下げする。子を持つ塊の後ろには空のソフト改行を 1 本入れて塊を離す
    // (空段落ではなく空のソフト改行なので点が付かない。段落間隔の API が無くても区切りが作れる)。
    // 記号と空行を入れるかは幅・文字サイズ・高さが決まってから決めるので、組み立て方だけ返す
    const INDENT = "　";
    const groups = []; // { head, children: [] }
    const childLines = [];
    for (const it of arr) {
      if (Array.isArray(it)) {
        if (!groups.length) groups.push({ head: "", children: [] });
        for (const sub of it) {
          const t = itemToText(sub);
          childLines.push(t);
          groups[groups.length - 1].children.push(t);
        }
      } else groups.push({ head: itemToText(it), children: [] });
    }
    const build = (opt) =>
      groups
        .map((g, i) => {
          let line = g.head;
          for (const t of g.children) line += "\v" + INDENT + (opt.marker ? "– " : "") + t;
          if (opt.spacer && g.children.length && i < groups.length - 1) line += "\v";
          return line;
        })
        .join("\n");
    return { text: build({ marker: false, spacer: true }), build, childLines, list: true, nested: true };
  }
  /** 入れ子の見せ方を、幅・文字サイズ・高さが決まってから決める。
   *  - 記号「–」: 子が 1 行に収まるなら付けない(字下げだけで階層が読める。付けると騒がしい)。
   *    折り返す子があるときだけ付ける(付けないと子と子の切れ目が消えて 1 つの文に見える)。
   *  - 塊を離す空行: 入るときだけ入れる。入らないなら捨てる(区切りより内容が入ることが先) */
  function nestedText(body, w, fs, h) {
    if (!body || !body.nested || !body.build) return body ? body.text : "";
    const marker = (body.childLines || []).some((l) => estimateLines(l, w, fs) > 1);
    const spaced = body.build({ marker, spacer: true });
    if (h == null || textHeight(spaced, w, fs) <= h) return spaced;
    return body.build({ marker, spacer: false });
  }
  function cellBody(node) {
    if (Array.isArray(node.items) && node.items.length) return itemsBody(node.items);
    return { text: node.text || "", list: false, nested: false };
  }
  function cellText(c) {
    return typeof c === "string" || typeof c === "number" ? String(c) : c && c.text ? c.text : "";
  }
  function formatNum(v) {
    if (Math.abs(v) >= 1000) return v.toLocaleString("en-US", { maximumFractionDigits: 1 });
    return String(Math.round(v * 10) / 10);
  }

  // =====================================================================
  //  ブロック描画
  // =====================================================================
  const LEAF = {};

  /** パネル見出し: 中央揃えのテキスト(+アイコン)を 1 組としてセンタリングし、下に太い罫を引く。塗り帯にしない */
  function heading(c, x, y, w, text, opt) {
    opt = opt || {};
    const m = headingMetrics(text, w, !!opt.icon);
    const h = m.h;
    const hl = !!opt.highlight;
    const align = opt.align || "center";
    const t = stripBold(text || "");
    const iconSize = Math.round(m.fs * 1.35);
    const gapI = 6;
    const color = opt.color || (hl ? c.P.accent : c.P.text);
    const ruleColor = opt.ruleColor || (hl ? c.P.accent : c.P.fillDark);
    const iconColor = opt.color === c.P.textOnDark ? c.P.textOnDark : hl ? c.P.accent : c.P.fillDark; // 濃い地の上では白
    // 見出しテキストは罫線(線分)と同じ幅のテキストシェイプに載せて中央揃え(アイコンは左端に添えるだけで幅は変えない)
    if (opt.icon) c.prims.push(image(opt.icon, x + PAD - 2, y + h - 4 - iconSize - 2, iconSize, iconSize, iconColor));
    // 2 行見出しが上下マージンで溢れないよう、見出しの内部余白は 4pt(既定 8pt だと 2 行×18pt が入らず 2 行目が消える)
    if (opt.band) c.prims.push(rect(x, y, w, h - 2, { fill: opt.band })); // 反転見出し(濃い帯 + 白文字)。罫の代わりに帯で塊を示す
    c.prims.push(textBox(x, y, w, h - 4, t, { fontSize: m.fs, bold: true, color, align, valign: opt.band ? "middle" : "bottom", autofit: "none", pad: 4 }));
    if (opt.rule !== false && !opt.band) c.prims.push(line(x, y + h - 2, x + w, y + h - 2, ruleColor, RULE_THICK)); // パネル見出しの罫(同じ塊の中で二重に引かない)
    return h;
  }

  function cellNaturalHeight(node, w) {
    const inner = w - PAD * 2;
    if (node.style === "takeaway") return Math.max(36, textHeight(cellBody(node).text, w - PAD * 4, FONT.body) + PAD * 2);
    if (node.fill && node.fill !== "none" && !node.shape) {
      // 塗りの箱: 中身 + 余白(上下 10pt)
      const b0 = cellBody(node);
      const capH0 = node.caption ? FONT.caption * FONT.lineHeight + 6 : 0;
      const hh0 = node.head && node.head.trim() ? (node.num ? FONT.num * FONT.lineHeight + 4 : headingMetrics(node.head, w, false).h) : 0;
      const bodyH0 = b0.text ? textHeight(b0.text, inner - (b0.list ? BULLET_INDENT : 0), FONT.body, 4) * 1.06 + PAD * 2 + 2 : 0;
      return capH0 + hh0 + (hh0 && bodyH0 ? GAP / 2 : 0) + bodyH0 + 20;
    }
    if (node.shape === "step") {
      const bs = cellBody(node);
      if (!bs.text) return 48;
      return Math.max(96, FONT.head * FONT.lineHeight + 8 + textHeight(bs.text, inner, FONT.body) * 1.06 + PAD * 2 + 12);
    }
    if (node.shape) return node.text || (node.items && node.items.length) ? 44 + GAP / 2 + 50 : 44;
    const b = cellBody(node);
    const hasBody = !!b.text;
    const hasHead = !!(node.head && node.head.trim());
    const iconOnly = node.icon && !hasHead && !hasBody;
    if (iconOnly) return 64;
    if (!hasHead && !hasBody) return 40;
    const pref = node.size === "large" ? FONT.large : FONT.body;
    const bodyH = hasBody ? textHeight(b.text, inner - (b.list ? BULLET_INDENT : 0) - (node.level ? SPACE.levelIndent : 0), pref, 4) * 1.06 + PAD * 2 + 2 : 0;
    const capH = node.caption ? FONT.caption * FONT.lineHeight + 6 : 0;
    if (!hasHead) return capH + bodyH + (node.fill && node.fill !== "none" ? 8 : 0);
    const hh = node.num ? FONT.num * FONT.lineHeight + 4 : headingMetrics(node.head, w, !!node.icon).h;
    if (!hasBody) return capH + hh;
    return capH + hh + GAP / 2 + Math.max(40, bodyH);
  }

  /**
   * cell: テキストを持つ唯一の部品。見た目は属性で決まる。
   *  - head: 太字見出し(塗りなし/薄地のとき下に太い罫)。
   *  - text / items: 本文(18pt 基本、fit)。items 2 件以上で箇条書き、ネスト配列は第 2 階層。
   *  - shape: chevron / home … 矢羽(塗り+白文字、中央揃え)。
   *  - fill: none / light / dark / tint … 塗り(塗りのある箱に枠線は付けない)。
   *  - highlight: 見出し・罫・矢羽塗りを強調色に。
   */
  LEAF.cell = function (node, r, c) {
    const P = c.P;
    const hl = !!node.highlight;
    const b = cellBody(node);
    const hasBody = !!b.text;
    const head = stripBold(node.head || "").trim();
    if (node._inPanel && !head && Array.isArray(node.items) && node.items.length > 1) return renderNestedItems(node, r, c);

    // 示唆帯: 版面幅の薄い地 + 太字の一文(左バーは付けない。callout は薄い地だけ)
    if (node.style === "takeaway") {
      const band = hl && !P.mono ? P.accentTint : P.fillLight;
      c.prims.push(rect(r.x, r.y, r.w, r.h, { fill: band }));
      const fs = fitBody(c, b.text, r.w - PAD * 4, r.h - PAD, FONT.body, 0, FONT.large);
      c.prims.push(textBox(r.x, r.y, r.w, r.h, b.text, { fontSize: fs, bold: true, color: P.text, valign: "middle", align: "left", pad: PAD * 1.5, gid: node._gid }));
      return;
    }
    // 箱フロー: 薄い箱に太字の見出し(強調は濃い塗り + 白文字)。矢印は容器側が箱の間に描く
    if (node.shape === "step") {
      // 箱フローは全段を同じ薄い箱で描く(一部だけ黒く塗る強調はしない。立てたい段は太字の見出しで十分)
      c.prims.push(rect(r.x, r.y, r.w, r.h, { fill: P.fillLight }));
      const color = P.text;
      if (hasBody) {
        const textNeed = textHeight(b.text, r.w - PAD * 2, FONT.body) + PAD * 2; // 見出し帯は本文が入る分だけ譲る
        const hh = Math.max(FONT.head * FONT.lineHeight + 8, Math.min(r.h * 0.42, r.h - textNeed));
        c.prims.push(textBox(r.x, r.y, r.w, hh, head, { fontSize: FONT.head, bold: true, color, align: "center", valign: "bottom", pad: 4, autofit: "none" }));
        const fs = fitBody(c, b.text, r.w - PAD * 2, r.h - hh - PAD, FONT.body, 3, FONT.body);
        c.prims.push(textBox(r.x, r.y + hh, r.w, r.h - hh, b.text, { fontSize: fs, color: P.textMuted, align: "center", valign: "top", pad: 4, bullets: false, gid: node._gid }));
      } else {
        c.prims.push(textBox(r.x, r.y, r.w, r.h, head || b.text, { fontSize: FONT.head, bold: true, color, align: "center", valign: "middle", pad: 6, autofit: "none" }));
      }
      return;
    }
    // 矢羽・五角形: 帯として描く(本文があれば帯の下)
    if (node.shape) {
      const bandH = hasBody ? 44 : Math.min(r.h, 56);
      const bandY = hasBody ? r.y : r.y + (r.h - bandH) / 2;
      // 矢羽はホームベース型。幅 100pt 未満・高さ 28pt 未満では矢の形が表現できないので長方形にする
      const shape = r.w >= 100 && bandH >= 28 ? "homePlate" : "rect";
      c.prims.push(textBox(r.x, bandY, r.w, bandH, head || stripBold(b.text), { shape, fill: hl ? P.accent : P.fillDark, fontSize: FONT.body, bold: true, color: P.textOnDark, align: "center", valign: "middle", pad: 4, autofit: "none" }));
      if (!hasBody || !head) return;
      const by = r.y + bandH + GAP / 2;
      const bh = r.h - bandH - GAP / 2;
      const fs = fitBody(c, b.text, r.w - PAD * 2 - (b.list ? BULLET_INDENT : 0), bh - PAD * 2, FONT.body, 4, undefined, b.list);
      c.prims.push(textBox(r.x, by, r.w, bh, b.text, { fontSize: fs, color: P.text, bullets: b.list, valign: "top", shrunk: fs < FONT.min }));
      return;
    }

    const fillKey = node.fill && node.fill !== "none" ? node.fill : null;
    // 濃い塗り(dark/accent)は箱全体を塗らず「見出し帯の反転(濃い帯 + 白文字)」で立てる。見出しが無ければ薄い塗りに落とす
    const invert = (fillKey === "dark" || fillKey === "accent") && !!head;
    const fill = invert ? null : fillKey ? { light: P.fillLight, dark: P.fillLight, tint: P.accentTint, accent: P.accentTint }[fillKey] || null : null;
    const dark = false;
    const textColor = P.text;
    if (fill) c.prims.push(rect(r.x, r.y, r.w, r.h, { fill }));
    else if (hl && !head && hasBody) c.prims.push(rect(r.x, r.y, r.w, r.h, { fill: P.accentTint })); // 見出しなしの強調セルは薄地で示す

    const indent = node.level ? SPACE.levelIndent * node.level : 0;
    const x = r.x + indent;
    const w = r.w - indent;

    // アイコンだけ
    if (node.icon && !head && !hasBody) {
      const d = Math.max(32, Math.min(64, r.h - 8, w - 8));
      c.prims.push(image(node.icon, x + (w - d) / 2, r.y + (r.h - d) / 2, d, d, hl ? P.accent : P.fillDark));
      return;
    }
    let y = r.y;
    let bh = r.h;
    // 小ラベル(caption): 見出しや本文の上に小さく灰色で置く(「何が起きるか」→「内部計算で推論を深める」)
    const cap = stripBold(node.caption || "").trim();
    if (cap) {
      const ch = FONT.caption * FONT.lineHeight + 6;
      c.prims.push(textBox(x, y, w, ch, cap, { fontSize: FONT.caption, color: dark ? P.textOnDark : P.textMuted, align: "left", valign: "bottom", pad: 4, autofit: "none", role: "caption" }));
      y += ch;
      bh -= ch;
    }
    // 番号(num): 「01」を大きく置き、見出しはその右
    const num = stripBold(node.num || "").trim();
    if (num && head) {
      const nh = FONT.num * FONT.lineHeight + 4;
      const nw = measure(num, FONT.num) + PAD * 2 + 4;
      if (invert) c.prims.push(rect(x, y, w, nh, { fill: hl && !P.mono ? P.accent : P.fillDark })); // 反転見出し帯
      c.prims.push(textBox(x, y, nw, nh, num, { fontSize: FONT.num, bold: true, color: invert ? P.textOnDark : hl ? P.accent : P.fillDark, align: "left", valign: "bottom", pad: 4, autofit: "none", role: "num" }));
      c.prims.push(textBox(x + nw, y, w - nw, nh, head, { fontSize: FONT.head, bold: true, color: invert ? P.textOnDark : P.text, align: "left", valign: "bottom", pad: 4, autofit: "none" }));
      if (!node._inPanel && !fill && !invert) c.prims.push(line(x, y + nh, x + w, y + nh, hl ? P.accent : P.fillDark, RULE_THICK));
      if (!hasBody) return;
      y += nh + GAP / 2;
      bh = r.y + r.h - y;
      const fsN = fitBody(c, b.text, w - PAD * 2 - (b.list ? BULLET_INDENT : 0), bh - PAD * 2, FONT.body, b.list ? 4 : 0, FONT.body, b.list);
      c.prims.push(textBox(x, y, w, bh, b.text, { fontSize: fsN, color: dark ? P.textOnDark : P.text, bullets: b.list, align: "left", valign: "top", pad: PAD, shrunk: fsN < FONT.min, gid: node._gid }));
      return;
    }
    if (head) {
      // アイコンは余白が余るときだけ(本文が枠の 6 割未満)。密な本文には付けない
      const bodyNeed = hasBody ? textHeight(b.text, w - PAD * 2 - (b.list ? BULLET_INDENT : 0), FONT.body, 4) + PAD * 2 : 0;
      const roomy = !hasBody || bodyNeed < (r.h - headingMetrics(head, w, false).h) * 0.6;
      // 黒下線はパネルの最上位だけ。入れ子の見出しは罫を引かず、太字の見出し行として置く
      const panelHead = !node._inPanel && !fill; // 塗りのある箱は塗り自体が塊を示すので罫を引かない
      const topLevel = !node._inPanel;
      const hh = heading(c, x, y, w, head, { highlight: hl, icon: roomy ? node.icon : undefined, color: invert ? P.textOnDark : undefined, band: invert ? (hl && !P.mono ? P.accent : P.fillDark) : null, align: topLevel || invert ? (node.align === "left" ? "left" : "center") : "left", rule: panelHead && !invert });
      if (!hasBody) return;
      y += hh + GAP / 2;
      bh = r.y + r.h - y; // caption 分を含めて枠の残り
    }
    if (!hasBody) return;
    const pref = node.size === "large" ? FONT.large : FONT.body;
    const padX = fillKey === "tint" || fillKey === "light" ? PAD * 1.5 : PAD;
    const fs = fitBody(c, b.text, w - padX * 2 - (b.list ? BULLET_INDENT : 0), bh - PAD * 2, pref, b.list ? 4 : 0, head ? FONT.large : FONT.large, b.list);
    const align = node.align || "left"; // 本文は常に左。右寄せは格子の数値列だけに限る(全体の平仄)
    // 下線パネルの本文は、枠に対して文字が少ないときは縦中央に置く(上に張り付いて下が空く密度の偏りをなくす)
    const bw = w - padX * 2 - (b.list ? BULLET_INDENT : 0);
    const bodyText = nestedText(b, bw, fs, bh - PAD * 2); // 記号と塊の空行は、実際の幅・文字サイズ・高さで決める
    const need = textHeight(bodyText, bw, fs, b.list ? 4 : 0) + PAD * 2;
    const roomyPanel = !fill && !node._inPanel && bh > need * 1.4;
    const valign = roomyPanel ? "middle" : head ? "top" : fill ? "middle" : node.valign || "top";
    c.prims.push(textBox(x, y, w, bh, bodyText, { fontSize: fs, bold: !!node.bold, color: textColor, bullets: b.list, align, valign, pad: padX, shrunk: fs < FONT.min, gid: node._gid }));
  };

  // ---------- table: セルを組み合わせた格子(ネイティブ表ではない) ----------
  /** 格子の寸法計算(自然高さ・列幅・行高さ・フォント)。r が null なら自然高さのみ */
  function matrixMetrics(node, w, r, ctx) {
    const colHeaders = capList(node.colHeaders, 6, ctx, "列");
    const rows = capList(node.rows, 12, ctx, "行");
    const ncols = Math.max(1, colHeaders.length || Math.max.apply(null, [1].concat(rows.map((rw) => (rw.cells || []).length))));
    const hasRowHead = !!(node.corner && node.corner.trim()) || rows.some((rw) => rw.head && String(rw.head).trim());
    const chevron = node.headShape === "chevron" && hasRowHead;
    const longestHead = hasRowHead ? Math.max.apply(null, [0].concat(rows.map((rw) => measure(stripBold(rw.head || ""), FONT.body)), [measure(stripBold(node.corner || ""), FONT.body)])) : 0;
    const numW = node.numbered ? measure("00", FONT.body) + SPACE.cellPad * 2 + 2 : 0;
    const rowHeadW = hasRowHead ? Math.max(w * 0.15, Math.min(w * 0.34, longestHead + SPACE.cellPad * 2 + 12 + numW)) : 0;
    // 行見出しのグループ化(2 階層の行見出し)。colGroups(列方向)と同じ {text,span} 形式を行方向に転用する
    const hasRowGroups = Array.isArray(node.rowGroups) && node.rowGroups.length > 0;
    const longestGroup = hasRowGroups ? Math.max.apply(null, [0].concat(node.rowGroups.map((g) => measure(stripBold(g.text || ""), FONT.body)))) : 0;
    const groupW = hasRowGroups ? Math.max(w * 0.08, Math.min(w * 0.22, longestGroup + SPACE.cellPad * 2 + 12)) : 0;
    const axisW = node.axes && node.axes.y ? FONT.body * FONT.lineHeight + 6 : 0;
    const axisH = node.axes && node.axes.x ? FONT.body * FONT.lineHeight + 6 : 0;
    const gridX = axisW + groupW;
    const gridW = w - axisW - groupW;
    // 列幅は内容量に比例させる(等分だと「稼働まで」のような短い列が広く、説明列が狭くなる)
    const cellPad = SPACE.cellPad;
    const contentW = gridW - rowHeadW;
    const want = [];
    for (let j = 0; j < ncols; j++) {
      let mx = measure(stripBold(colHeaders[j] || ""), FONT.body);
      for (const rw of rows) {
        const t = cellText((rw.cells || [])[j]);
        // 折り返す前提の長文は 1 行分の重みを抑える(平方根で圧縮)
        mx = Math.max(mx, Math.min(measure(t, FONT.body), Math.sqrt(Math.max(1, measure(t, FONT.body))) * 9));
      }
      want.push(Math.max(mx, 24));
    }
    const wantSum = want.reduce((a, b) => a + b, 0) || 1;
    const minW = contentW * (ncols > 1 ? 0.12 : 1);
    const maxW = contentW * (ncols > 1 ? 0.5 : 1);
    let colWs = want.map((w) => Math.min(maxW, Math.max(minW, (contentW * w) / wantSum)));
    const adj = contentW / (colWs.reduce((a, b) => a + b, 0) || 1);
    colWs = colWs.map((w) => w * adj);
    // 列見出しが 1 行で入る幅は必ず確保する(見出しが切れないように、広い列から融通する)
    for (let j = 0; j < ncols; j++) {
      const need = Math.min(maxW, measure(stripBold(colHeaders[j] || ""), FONT.body) + cellPad * 2 + 6);
      if (colWs[j] >= need) continue;
      let lack = need - colWs[j];
      colWs[j] = need;
      const order = colWs.map((w, k) => k).filter((k) => k !== j).sort((a, b) => colWs[b] - colWs[a]);
      for (const k of order) {
        const give = Math.min(lack, Math.max(0, colWs[k] - minW));
        colWs[k] -= give;
        lack -= give;
        if (lack <= 0.01) break;
      }
    }
    const colX = (j) => colWs.slice(0, j).reduce((a, b) => a + b, 0);
    const colW = contentW / ncols; // 高さ見積り等で使う平均値
    const groupH = Array.isArray(node.colGroups) && node.colGroups.length ? headerHeight() - 6 : 0;
    let headH = colHeaders.length ? (node._composition ? Math.max(headerHeight(), ...colHeaders.map((t, j) => textHeight(t, colWs[j] - cellPad * 2, FONT.body + 2) + cellPad * 2 + 6)) : headerHeight()) : 0;
    const rowNeed = (f) =>
      rows.map((rw) => {
        let h = hasRowHead ? textHeight(stripBold(rw.head || ""), rowHeadW - cellPad * 2, f) + cellPad * 2 : 0;
        (rw.cells || []).slice(0, ncols).forEach((cell, j) => {
          h = Math.max(h, textHeight(cellText(cell), colWs[j] - cellPad * 2, f, 3) + cellPad * 2);
        });
        return Math.max(h, f * FONT.lineHeight + cellPad * 2 + 6, chevron ? 40 : 0);
      });
    let fs = FONT.body;
    // 短いセル(14 字以下)が折り返すなら、折り返さなくなるまで 14pt を下限に縮める(「◎ 1,200 億円」が 2 行になるのを防ぐ)
    const shortWraps = (f) => {
      const texts = [];
      rows.forEach((rw) => {
        if (hasRowHead) texts.push([stripBold(rw.head || ""), rowHeadW]);
        (rw.cells || []).slice(0, ncols).forEach((cell, j) => texts.push([cellText(cell), colWs[j]]));
      });
      return texts.some(([t, w0]) => t && !t.includes("\n") && stripBold(t).length <= 14 && estimateLines(t, w0 - cellPad * 2, f) > 1);
    };
    while (!node._composition && fs > FONT.tableMin && shortWraps(fs)) fs -= 1;
    let needs = rowNeed(fs);
    const natural = groupH + headH + needs.reduce((a, b) => a + b, 0) + axisH;
    if (!r) return { natural };
    let availRows = r.h - groupH - headH - axisH;
    while (fs > FONT.tableMin && needs.reduce((a, b) => a + b, 0) > availRows) {
      fs -= 1;
      needs = rowNeed(fs);
    }
    if (node._composition && colHeaders.length) {
      const actualHeadH = Math.max(...colHeaders.map((t, j) => textHeight(t, colWs[j] - cellPad * 2, Math.min(FONT.head, fs + 2)) + cellPad * 2 + 6));
      availRows += headH - actualHeadH;
      headH = actualHeadH;
    }
    const needSum = needs.reduce((a, b) => a + b, 0) || 1;
    const extra = Math.max(0, availRows - needSum);
    // 収まらないときは行高さを比例縮小して領域内に収める。余るときは行を間延びさせず(自然高さの +60% まで)、残りは下の余白にする
    const rowHeights = needSum > availRows ? needs.map((h) => (h * Math.max(40, availRows)) / needSum) : needs.map((h) => h + (node._composition ? extra / Math.max(1, rows.length) : Math.min(extra / Math.max(1, rows.length), h * 0.6)));
    const hasFills = rows.some((rw) => (rw.cells || []).some((cl) => cl && typeof cl === "object" && cl.fill));
    return { natural, colHeaders, rows, ncols, hasRowHead, chevron, rowHeadW, colW, colWs, colX, cellPad, groupH, groupW, headH, fs, rowHeights, overflow: needSum > availRows, gridX, gridW, axisW, axisH, hasFills };
  }

  LEAF.table = function (node, r, c) {
    const P = c.P;
    const m = matrixMetrics(node, r.w, r, c);
    const mid = (c.matrixSeq = (c.matrixSeq || 0) + 1); // 隣り合う格子の文字サイズを揃えるための識別子
    // 揃えは列単位で決める(数値だけの列は右揃え、文字が混じる列は左揃え)。行ごとに右/左が混在すると平仄が崩れる
    const colNumeric = [];
    for (let j = 0; j < m.ncols; j++) {
      const vals = m.rows.map((rw) => cellText((rw.cells || [])[j])).filter((t) => t.trim() && !isNA(t));
      colNumeric.push(vals.length > 0 && vals.every((t) => isNumeric(t)));
    }
    const matStart = c.prims.length;
    if (m.overflow) c.warnings.push("table overflow at " + m.fs + "pt");
    const x0 = r.x + m.gridX;
    const headFs = Math.min(FONT.head, m.fs + 2);
    let y = r.y;

    // 列グループ見出し(2 階層の列見出し)
    if (m.groupH) {
      let cx = x0 + m.rowHeadW;
      let gi = 0;
      for (const g of node.colGroups) {
        const span = Math.max(1, Math.min(m.ncols - gi, Number(g.span) || 1));
        const gw = m.colWs.slice(gi, gi + span).reduce((a, b) => a + b, 0);
        gi += span;
        c.prims.push(textBox(cx, y, gw, m.groupH - 4, stripBold(g.text || ""), { fontSize: m.fs, bold: true, color: P.text, align: "center", valign: "bottom", pad: 2, autofit: "none", role: "matrixcell" }));
        c.prims.push(line(cx + 4, y + m.groupH - 3, cx + gw - 4, y + m.groupH - 3, P.fillDark, RULE_THIN));
        cx += gw;
      }
      y += m.groupH;
    }
    // 列見出し: 既定は塗りなし・太字。headFill:"dark" なら濃い帯 + 白文字(参考デッキの評価表)
    const headDark = m.colHeaders.length > 0 && node.headFill !== "none" && node.headFill !== "light"; // 列見出しは既定で濃い帯(参照デッキの列名帯)
    if (m.headH) {
      // 左上(行見出し × 列見出しの角)は基本的に空ける。corner に文字があるときだけ帯と文字を置く
      const cornerText = stripBold(node.corner || "").trim();
      const bandX = m.hasRowHead && !cornerText ? x0 + m.rowHeadW : x0;
      const bandW = m.gridW - (bandX - x0);
      // headDark(黒帯の反転見出し)は列をまたいで1本につながった帯そのものが意図した見た目なので、
      // 共有の背面シェイプのままにする。淡い地(_composition)は列ごとの塗りが要るだけなので、
      // 見出しシェイプの背面にもう1枚重ねず、各列見出しのテキストシェイプ自身に塗りを持たせる(隣とは小さな余白で分ける)
      if (headDark) c.prims.push(rect(bandX, y, bandW, m.headH - 2, { fill: P.fillDark }));
      if (m.hasRowHead && cornerText) c.prims.push(textBox(x0, y, m.rowHeadW, m.headH - 2, cornerText, { fontSize: headFs, bold: true, color: headDark ? P.textOnDark : P.text, align: "left", valign: headDark ? "middle" : "bottom", pad: m.cellPad, autofit: "none", role: "matrixcell" }));
      const colGap = 1.5;
      // 塗り(headFill:"light"・強調列)と罫は同じ「境界を示す」役目が重複する。塗りがあれば罫は引かない、
      // 塗りが無い(headFill:"none")ときだけ罫で区切る。dark は帯そのものが境界を兼ねるので元から罫を引かない
      const fillOf = (j) => (headDark ? null : node.highlightCol === j ? P.accentTint : node.headFill === "none" ? null : P.fillLight);
      m.colHeaders.forEach((h, j) => {
        const hl = node.highlightCol === j;
        const cx = x0 + m.rowHeadW + m.colX(j);
        const cellFill = fillOf(j);
        // 列見出しは中身の揃えに合わせる(数値列なら右、それ以外は左)。見出しだけ中央にしない
        const colAlign = colNumeric[j] ? "right" : "left";
        c.prims.push(
          textBox(cx + (cellFill ? colGap : 0), y, m.colWs[j] - (cellFill ? colGap * 2 : 0), m.headH - 2, stripBold(h), {
            fontSize: headFs,
            bold: true,
            color: headDark ? P.textOnDark : hl ? P.accent : P.text,
            align: colAlign,
            valign: headDark || cellFill ? "middle" : "bottom",
            pad: m.cellPad,
            fill: cellFill,
            autofit: "none",
            role: "matrixcell",
          })
        );
      });
      // 列見出しの罫: パネル先頭の格子なら太い罫(この行がパネル見出しを兼ねる。1本のパネル見出し罫として繋げる)、
      // 単体の格子で塗りが無い列だけ細い罫(塗りがある列は塗りの下端が境界を兼ねるので罫を重ねない)、パネルの中では引かない
      if (node._panelHead && !headDark) c.prims.push(line(bandX, y + m.headH - 2, x0 + m.gridW, y + m.headH - 2, P.fillDark, RULE_THICK)); // 隣のパネル見出し罫(h-2)と同じ y
      else if (!node._inPanel && !headDark) {
        const ruleY = y + m.headH - 2;
        m.colHeaders.forEach((h, j) => {
          if (fillOf(j)) return;
          const cx = x0 + m.rowHeadW + m.colX(j);
          c.prims.push(line(cx + colGap, ruleY, cx + m.colWs[j] - colGap, ruleY, P.line, 1.2));
        });
      }
      y += m.headH;
    }
    // 行見出しのグループ帯(2 階層の行見出し): 対象の行の高さを合計した 1 個の箱に、グループ名を縦横中央で置く
    if (m.groupW) {
      const groupX = r.x + m.axisW;
      let gy = y,
        gi = 0;
      for (const g of node.rowGroups) {
        const span = Math.max(1, Math.min(m.rows.length - gi, Number(g.span) || 1));
        const gh = m.rowHeights.slice(gi, gi + span).reduce((a, b) => a + b, 0);
        gi += span;
        if (gh <= 0) continue;
        if (node.rowHeadFill !== "none") c.prims.push(rect(groupX, gy + 2, m.groupW - 4, gh - 4, { fill: P.fillLight }));
        c.prims.push(textBox(groupX, gy, m.groupW - 4, gh, stripBold(g.text || ""), { fontSize: m.fs, bold: true, color: P.text, align: "center", valign: "middle", pad: m.cellPad, autofit: "none", role: "matrixcell" }));
        gy += gh;
      }
    }
    // 本文行
    m.rows.forEach((rw, i) => {
      const rh = m.rowHeights[i];
      const last = i === m.rows.length - 1;
      const rowHl = !!rw.highlight;
      if (rowHl) c.prims.push(rect(x0 + m.rowHeadW, y, m.gridW - m.rowHeadW, rh, { fill: P.accentTint }));
      if (m.hasRowHead) {
        const headText = (node._composition && m.chevron && node.numbered ? String(i + 1).padStart(2, "0") + "  " : "") + stripBold(rw.head || "");
        if (node.numbered && !m.chevron) {
          // 行番号「01」を行見出しの左に添える(参考デッキの判断軸の表)
          const numT = String(i + 1).padStart(2, "0");
          const nw = measure(numT, m.fs) + m.cellPad * 2 + 2;
          c.prims.push(textBox(x0, y, nw, rh, numT, { fontSize: m.fs, bold: true, color: rowHl ? P.accent : P.fillDark, align: "left", valign: "middle", pad: m.cellPad, autofit: "none", role: "matrixcell" }));
          c.prims.push(textBox(x0 + nw, y, m.rowHeadW - nw, rh, headText, { fontSize: m.fs, bold: true, color: P.text, align: "left", valign: "middle", pad: m.cellPad, autofit: "none", role: "matrixcell" }));
        } else if (m.chevron) {
          // 縦向きの矢羽(先頭は五角形): 図形を 90° 回転させ、文字は別テキストボックスで重ねる(文字は回転させない)
          const ov = SPACE.chevronOverlap;
          const sw = m.rowHeadW - 4,
            sh = rh - 3;
          const cx = x0 + m.rowHeadW / 2,
            cy = y + sh / 2;
          if (m.rowHeadW >= 110 && rh >= 30) c.prims.push(rect(cx - sh / 2, cy - sw / 2, sh, sw, { shape: "homePlate", rotation: 90, fill: node._composition ? P.fillLight : rowHl ? P.accent : P.fillDark }));
          else c.prims.push(rect(x0, y + 1, m.rowHeadW - 4, rh - 2, { fill: rowHl ? P.accent : P.fillDark })); // 小さいときは矢の形にせず長方形
          c.prims.push(textBox(x0, y, m.rowHeadW - 4, rh, headText, { fontSize: m.fs, bold: true, color: node._composition ? P.text : P.textOnDark, align: "center", valign: "middle", pad: 4, autofit: "none", role: "matrixcell" }));
        } else {
          if (node.rowHeadFill !== "none" && headText) c.prims.push(rect(x0, y + 2, m.rowHeadW - 4, rh - 4, { fill: P.fillLight })); // 行名は薄い箱
          c.prims.push(textBox(x0, y, m.rowHeadW - 4, rh, headText, { fontSize: m.fs, bold: true, color: P.text, align: "left", valign: "middle", pad: m.cellPad, autofit: "none", role: "matrixcell" }));
        }
      }
      for (let j = 0; j < m.ncols; j++) {
        const cell = (rw.cells || [])[j];
        const runs = parseRuns(cellText(cell));
        const cx = x0 + m.rowHeadW + m.colX(j);
        const cw = m.colWs[j];
        const cellHl = !!(cell && cell.highlight) || node.highlightCol === j;
        const fillKey = cell && typeof cell === "object" && cell.fill && cell.fill !== "none" ? cell.fill : null;
        const fill = fillKey ? { light: P.fillLight, dark: P.fillDark, tint: P.accentTint, accent: P.accent }[fillKey] : cellHl && !rowHl ? P.accentTint : null;
        const dark = fillKey === "dark" || fillKey === "accent";
        if (fill) c.prims.push(rect(cx, y + 2, cw, rh - 4, { fill })); // 塗りセルは左右の余白なしで隣と連結(ガントの帯)
        // セルは文章版(text)と箇条書き版(items)の両方を持てる。実際に割り当てられた列幅で文章版が
        // 折り返す行数を測り、狭くて読みにくくなるときだけ箇条書き版に切り替える(静的な文分割はしない。
        // LLM に両方書かせておき、どちらを採用するかは幾何が決まってから機械的に選ぶ)
        const itemsArr = cell && typeof cell === "object" && Array.isArray(cell.items) && cell.items.length ? cell.items.map(String) : null;
        const proseText = cellText(cell);
        let raw = proseText;
        if (itemsArr) {
          const proseLines = estimateLines(stripBold(proseText), cw - (fill ? 4 : m.cellPad * 2), m.fs);
          if (proseLines > 2) raw = itemsArr.map((s) => "・" + s).join("\n");
        }
        // 「—」は明示的なダッシュか、塗りセルの無い表の空セルだけ(ガントの空白セルには出さない)
        const na = !fill && isNA(runs.text) && !m.hasFills; // 塗りセルのある表(ガント等)では空欄・ダッシュとも空白のまま
        const numeric = colNumeric[j];
        const lines = raw.split("\n");
        const marked = lines.filter((l) => /^\s*[・•\-–]/.test(l)).length;
        // 複数行: 全行に記号があれば記号を外してネイティブ箇条書き、記号なしなら箇条書き。
        // 先頭行だけ記号なし・残り全行に記号(見出し+箇条書き)は、見出し行を除いた分だけ箇条書きにする
        const isList = lines.length > 1 && (marked === 0 || marked === lines.length);
        const headBullet = !isList && lines.length > 1 && marked === lines.length - 1 && !/^\s*[・•\-–]/.test(lines[0]);
        const cellPad = fill ? 2 : m.cellPad;
        const cellOpt = {
          fontSize: m.fs,
          color: dark ? P.textOnDark : na ? P.lineLight : P.text,
          bold: false,
          align: fill ? "center" : numeric ? "right" : "left",
          pad: cellPad,
          role: "matrixcell",
        };
        if (!na && headBullet) {
          // 見出し行と箇条書き行は別のシェイプに分ける。1 つのシェイプの中で見出し段落だけ点を外すことは
          // Office.js ではできない(getSubstring の paragraphFormat が段落単位に効かず、シェイプ全体に掛かる。
          // 実機で確認済み)。疑似マーカーは使わず、箇条書き側は PowerPoint 本来の bulletFormat のまま
          const headText = lines[0];
          const listText = lines.slice(1).map((l) => l.replace(/^\s*[・•\-–]\s*/, "")).join("\n");
          // 見出しの箱は上マージンぶんだけ高くする(下マージンまで数えると、見出しと 1 項目目の間が
          // 空行 1 つぶん空いて「箇条書きの前で改行している」ように見える。前置きは箇条書き側の上マージンだけで足りる)
          const vPad = Math.min(cellPad, 8);
          const headH = textHeight(stripBold(headText), cw - cellPad * 2, m.fs) + vPad;
          const listH = textHeight(listText, cw - cellPad * 2 - BULLET_INDENT, m.fs) + vPad * 2;
          const top = y + Math.max(0, (rh - headH - listH) / 2);
          c.prims.push(textBox(cx, top, cw, headH, headText, Object.assign({}, cellOpt, { valign: "top", bullets: false })));
          c.prims.push(textBox(cx, top + headH, cw, Math.max(listH, y + rh - top - headH), listText, Object.assign({}, cellOpt, { valign: "top", bullets: true })));
        } else {
          const text = na ? "—" : isList && marked ? lines.map((l) => l.replace(/^\s*[・•\-–]\s*/, "")).join("\n") : raw;
          c.prims.push(textBox(cx, y, cw, rh, text, Object.assign({}, cellOpt, { valign: "middle", bullets: isList })));
        }
        if (node.axes && j > 0) c.prims.push(line(cx, y + 2, cx, y + rh - 2, P.lineLight, RULE_THIN)); // 4 象限などは列の間にも薄い縦罫
      }
      if (!last && !m.chevron) c.prims.push(line(x0 + (m.hasRowHead ? m.rowHeadW : 0), y + rh, x0 + m.gridW, y + rh, P.lineLight, RULE_THIN));
      else if (!last && m.chevron) c.prims.push(line(x0 + m.rowHeadW + 4, y + rh, x0 + m.gridW, y + rh, P.lineLight, RULE_THIN));
      y += rh;
    });
    for (let i = matStart; i < c.prims.length; i++) if (c.prims[i].role === "matrixcell") c.prims[i].mid = mid;
    // 4 象限などの軸ラベル
    if (node.axes) {
      if (node.axes.x) c.prims.push(textBox(x0 + m.rowHeadW, y + 2, m.gridW - m.rowHeadW, m.axisH - 2, stripBold(node.axes.x) + " →", { fontSize: FONT.body, color: P.textMuted, align: "center", valign: "top", pad: 0, autofit: "none" }));
      if (node.axes.y) {
        const gh = y - r.y;
        // 縦書き相当: 幅=格子高さ の箱を 270° 回転させて左に置く
        c.prims.push(textBox(r.x - gh / 2 + m.axisW / 2, r.y + gh / 2 - m.axisW / 2, gh, m.axisW, stripBold(node.axes.y) + " →", { fontSize: FONT.body, color: P.textMuted, align: "center", valign: "middle", pad: 0, autofit: "none", rotation: 270 }));
      }
    }
  };

  /** ntable(Figure): PPT ネイティブ表(1.8+)。密な数表向け。ヘッダー塗りなし・太い強調色罫、行間は薄罫、最終行の下は罫なし */
  LEAF.ntable = function (node, r, c) {
    const colHeaders = (node.colHeaders || []).slice(0, 6);
    const rows = (node.rows || []).slice(0, 12);
    const firstColW = Math.min(r.w * 0.26, 220);
    const otherW = (r.w - firstColW) / Math.max(1, colHeaders.length);
    const colWidths = [firstColW].concat(colHeaders.map(() => otherW));
    const cellPad = node._composition ? 4 : 6;
    const headH = headerHeight();
    const availRows = r.h - headH;
    let fs = FONT.body;
    const rowNeed = (f) =>
      rows.map((rw) => {
        let h = textHeight(stripBold(rw.head || ""), colWidths[0] - cellPad * 2, f) + cellPad * 2;
        (rw.cells || []).slice(0, colHeaders.length).forEach((cell, j) => {
          h = Math.max(h, textHeight(cellText(cell), colWidths[j + 1] - cellPad * 2, f) + cellPad * 2);
        });
        return Math.max(h, f * FONT.lineHeight + cellPad * 2 + 4);
      });
    let needs = rowNeed(fs);
    while (fs > FONT.ntableMin && needs.reduce((a, b) => a + b, 0) > availRows) {
      fs -= 1;
      needs = rowNeed(fs);
    }
    const needSum = needs.reduce((a, b) => a + b, 0) || 1;
    if (needSum > availRows) c.warnings.push("ntable overflow at " + fs + "pt");
    const extra = Math.max(0, availRows - needSum);
    const rowHeights = needSum > availRows ? needs.map((h) => (h * Math.max(40, availRows)) / needSum) : needs.map((h) => h + extra / Math.max(1, rows.length));
    const heights = [headH].concat(rowHeights);
    const totalH = heights.reduce((a, b) => a + b, 0);
    const headFs = Math.min(FONT.head, fs + 2);

    const cells = [];
    const headRow = [{ text: stripBold(node.corner || node.rowHeader || ""), boldRanges: [], fill: null, color: c.P.text, bold: true, align: "left", fontSize: headFs, borderBottom: { color: c.P.accent, weight: RULE_THICK } }];
    colHeaders.forEach((h, j) => {
      const hl = node.highlightCol === j;
      headRow.push({ text: stripBold(h), boldRanges: [], fill: hl ? c.P.accentTint : null, color: hl ? c.P.accent : c.P.text, bold: true, align: "center", fontSize: headFs, borderBottom: { color: c.P.accent, weight: RULE_THICK } });
    });
    cells.push(headRow);
    rows.forEach((rw, i) => {
      const last = i === rows.length - 1;
      const bb = last ? null : { color: c.P.lineLight, weight: RULE_THIN };
      const row = [{ text: stripBold(rw.head || ""), boldRanges: [], fill: rw.highlight ? c.P.accentTint : null, color: c.P.text, bold: true, align: "left", fontSize: fs, borderBottom: bb }];
      for (let j = 0; j < colHeaders.length; j++) {
        const cell = (rw.cells || [])[j];
        const runs = parseRuns(cellText(cell));
        const hl = rw.highlight || node.highlightCol === j || (cell && cell.highlight);
        const na = isNA(runs.text);
        row.push({ text: na ? "—" : runs.text, boldRanges: runs.boldRanges, fill: hl ? c.P.accentTint : null, color: na ? c.P.lineLight : c.P.text, bold: false, align: isNumeric(runs.text) ? "right" : "left", fontSize: fs, borderBottom: bb });
      }
      cells.push(row);
    });
    const fallback = [];
    let y = r.y;
    cells.forEach((row, i) => {
      let x = r.x;
      row.forEach((cell, j) => {
        fallback.push(rect(x, y, colWidths[j], heights[i], { fill: cell.fill, text: cell.text, boldRanges: cell.boldRanges, fontSize: cell.fontSize, bold: cell.bold, color: cell.color, align: cell.align, valign: "middle", pad: cellPad, wrap: true, autofit: "shrink" }));
        x += colWidths[j];
      });
      if (row[0].borderBottom) fallback.push(line(r.x, y + heights[i], r.x + r.w, y + heights[i], row[0].borderBottom.color, row[0].borderBottom.weight));
      y += heights[i];
    });
    c.prims.push({ kind: "table", x: r.x, y: r.y, w: r.w, h: totalH, colWidths, rowHeights: heights, cells, fontSize: fs, headFontSize: headFs, fontName: c.fontName || null, fallback });
  };

  // ---------- gantt: 期間 × タスクの帯(割付は決定論。LLM は期間ラベルと開始・終了だけ出す) ----------
  function ganttMetrics(node, w) {
    const periods = node.periods || [];
    const tasks = node.tasks || [];
    const longest = Math.max.apply(null, [60].concat(tasks.map((t) => measure(stripBold(t.label || ""), FONT.body))));
    const labelW = Math.max(w * 0.16, Math.min(w * 0.3, longest + SPACE.cellPad * 2 + 8));
    const groupH = node.groups && node.groups.length ? FONT.body * FONT.lineHeight + 6 : 0;
    const headH = FONT.body * FONT.lineHeight + 8;
    const msN = (node.milestones || []).length ? 1 : 0;
    const rowNat = Math.max(22, FONT.body * FONT.lineHeight + 10);
    const natural = groupH + headH + (tasks.length + msN) * rowNat + (node.today ? 0 : 0);
    return { periods, tasks, labelW, groupH, headH, rowNat, msN, natural };
  }
  LEAF.gantt = function (node, r, c) {
    if (node._composition) return renderGanttComposition(node, r, c);
    const P = c.P;
    const m = ganttMetrics(node, r.w);
    const n = Math.max(1, m.periods.length);
    const x0 = r.x + m.labelW;
    const pw = (r.w - m.labelW) / n;
    const rowsN = m.tasks.length + m.msN;
    const availRows = r.h - m.groupH - m.headH;
    let rowH = Math.min(Math.max(m.rowNat, availRows / Math.max(1, rowsN)), m.rowNat * 1.6);
    let fs = FONT.body;
    while (rowH * rowsN > availRows && fs > FONT.tableMin) {
      fs -= 1;
      rowH = Math.max(16, fs * FONT.lineHeight + 8);
    }
    if (rowH * rowsN > availRows) {
      c.warnings.push("gantt overflow: " + rowsN + " rows");
      rowH = Math.max(12, availRows / rowsN);
    }
    let y = r.y;
    // 上位の期間(年度・四半期)
    if (m.groupH) {
      let gi = 0;
      for (const g of node.groups) {
        const span = Math.max(1, Math.min(n - gi, Number(g.span) || 1));
        const gx = x0 + pw * gi, gw = pw * span;
        c.prims.push(textBox(gx, y, gw, m.groupH - 2, stripBold(g.text || ""), { fontSize: fs, bold: true, color: P.text, align: "center", valign: "bottom", pad: 2, autofit: "none", role: "matrixcell" }));
        c.prims.push(line(gx + 3, y + m.groupH - 2, gx + gw - 3, y + m.groupH - 2, P.fillDark, RULE_THIN));
        gi += span;
      }
      y += m.groupH;
    }
    // 期間ラベル行
    m.periods.forEach((p, i) => {
      c.prims.push(textBox(x0 + pw * i, y, pw, m.headH - 2, stripBold(p), { fontSize: fs, bold: true, color: P.text, align: "center", valign: "bottom", pad: 2, autofit: "none", role: "matrixcell" }));
    });
    const ruleY = y + m.headH - 2;
    c.prims.push(line(r.x, ruleY, r.x + r.w, ruleY, node._inPanel ? P.line : P.fillDark, node._inPanel ? 1 : RULE_THICK));
    y += m.headH;
    const bodyTop = y;
    const bodyH = rowH * rowsN;
    // 期間の区切り(薄い縦罫)
    for (let i = 1; i < n; i++) c.prims.push(line(x0 + pw * i, bodyTop, x0 + pw * i, bodyTop + bodyH, P.lineLight, RULE_THIN));
    // タスク行
    m.tasks.forEach((t, i) => {
      const ty = bodyTop + rowH * i;
      c.prims.push(textBox(r.x, ty, m.labelW, rowH, stripBold(t.label || ""), { fontSize: fs, bold: false, color: P.text, align: "left", valign: "middle", pad: SPACE.cellPad, autofit: "none", role: "matrixcell" }));
      const s0 = Math.max(0, Math.min(n - 1, t.start)), e0 = Math.max(s0, Math.min(n - 1, t.end));
      const bh = Math.min(rowH * 0.56, 22);
      const inset = Math.min(3, pw * 0.08);
      const bw = pw * (e0 - s0 + 1) - inset * 2;
      // 帯はホームベース型の矢羽(方法論: ガントは細かい時間軸に対して矢羽で関係を示す)。短すぎる帯は長方形
      c.prims.push(rect(x0 + pw * s0 + inset, ty + (rowH - bh) / 2, bw, bh, { shape: bw >= 28 ? "homePlate" : "rect", fill: t.highlight && !P.mono ? P.accent : t.highlight ? P.fillDark : P.fillMid }));
      if (t.note) c.prims.push(textBox(x0 + pw * (e0 + 1), ty, Math.max(40, r.x + r.w - x0 - pw * (e0 + 1)), rowH, stripBold(t.note), { fontSize: Math.max(8, fs - 2), color: P.textMuted, align: "left", valign: "middle", pad: 3, autofit: "none", role: "note" }));
      if (i < m.tasks.length - 1 || m.msN) c.prims.push(line(r.x, ty + rowH, r.x + r.w, ty + rowH, P.lineLight, RULE_THIN));
    });
    // マイルストーン行(◆ と短いラベル)
    if (m.msN) {
      const ty = bodyTop + rowH * m.tasks.length;
      c.prims.push(textBox(r.x, ty, m.labelW, rowH, "マイルストーン", { fontSize: fs, bold: false, color: P.textMuted, align: "left", valign: "middle", pad: SPACE.cellPad, autofit: "none", role: "matrixcell" }));
      (node.milestones || []).forEach((ms) => {
        const at = Math.max(0, Math.min(n - 1, ms.at));
        const cx = x0 + pw * at + pw / 2;
        const d = Math.min(10, rowH * 0.4);
        c.prims.push(rect(cx - d / 2, ty + rowH / 2 - d / 2, d, d, { fill: P.fillDark, rotation: 45 }));
        c.prims.push(textBox(cx + d / 2 + 2, ty, Math.max(50, pw * 2), rowH, stripBold(ms.label || ""), { fontSize: Math.max(8, fs - 2), color: P.text, align: "left", valign: "middle", pad: 2, autofit: "none", role: "note" }));
      });
    }
    // 今日の線
    if (node.today != null && node.today >= 0 && node.today < n) {
      const tx = x0 + pw * node.today + pw / 2;
      // 今日の線: 期間見出しの上端から本文の下端まで。ラベルは線の右に小さく(月ラベルに重ねない)
      c.prims.push(line(tx, bodyTop - m.headH + 2, tx, bodyTop + bodyH, P.fillDark, 1));
      c.prims.push(textBox(tx + 2, bodyTop + bodyH - 14, 40, 14, "今日", { fontSize: Math.max(8, fs - 3), color: P.textMuted, align: "left", valign: "bottom", pad: 0, autofit: "none", role: "note" }));
    }
  };

  // ---------- org: 体制図(箱 + コネクタのツリー) ----------
  function orgBoxH() {
    return FONT.body * FONT.lineHeight + FONT.caption * FONT.lineHeight + 12;
  }
  function orgDepth(node) {
    return 1 + Math.max.apply(null, [0].concat((node.children || []).map(orgDepth)));
  }
  function orgLeaves(node) {
    return node.children && node.children.length ? node.children.reduce((a, k) => a + orgLeaves(k), 0) : 1;
  }
  LEAF.org = function (node, r, c) {
    const P = c.P;
    const root = node.root;
    if (!root) return;
    const depth = orgDepth(root);
    const leaves = orgLeaves(root);
    const hGap = 12;
    const boxH = node._composition ? Math.min(96, Math.max(orgBoxH(), (r.h - 36 * (depth - 1)) / depth)) : orgBoxH();
    const vGap = Math.max(18, Math.min(52, (r.h - depth * boxH) / Math.max(1, depth - 1)));
    const maxBoxW = 240;
    const labels = [];
    (function walkL(nd) {
      if (!nd) return;
      labels.push(stripBold(nd.label || ""));
      if (node._composition && nd.sub) labels.push(stripBold(nd.sub));
      (nd.children || []).forEach(walkL);
    })(root);
    const longest = Math.max.apply(null, [60].concat(labels.map((t) => measure(t, FONT.body) + 16)));
    let boxW = Math.min(maxBoxW, (r.w - hGap * (leaves - 1)) / leaves, Math.max(longest, 120));
    boxW = Math.max(boxW, Math.min(longest, (r.w - hGap * (leaves - 1)) / leaves)); // ラベルが入る幅を優先(葉が多いときは葉の幅が上限)
    let fs = FONT.body;
    if (boxW < 70) {
      c.warnings.push("org too wide: " + leaves + " leaves");
      boxW = Math.max(56, boxW);
      fs = FONT.tableMin;
    }
    // サブツリーの幅(葉の幅の合計)。親は子の中央に置く
    const widthOf = (nd) => (nd.children && nd.children.length ? nd.children.reduce((a, k) => a + widthOf(k), 0) + hGap * (nd.children.length - 1) : boxW);
    const totalW = widthOf(root);
    const totalH = depth * boxH + (depth - 1) * vGap;
    const ox = r.x + Math.max(0, (r.w - totalW) / 2);
    const oy = r.y + Math.max(0, (r.h - totalH) / 2);
    const place = (nd, left, level) => {
      const w = widthOf(nd);
      const cx = left + w / 2;
      const y = oy + level * (boxH + vGap);
      const dark = node._composition ? false : level === 0 || nd.highlight;
      c.prims.push(rect(cx - boxW / 2, y, boxW, boxH, { fill: dark ? P.fillDark : P.fillLight }));
      const label = stripBold(nd.label || "");
      const sub = stripBold(nd.sub || "");
      if (sub) {
        const lfs = fitFont(label, boxW - 8, boxH * 0.58, fs, FONT.tableMin); // 箱に入らないラベルは縮める(切らない)
        c.prims.push(textBox(cx - boxW / 2, y, boxW, boxH * 0.58, label, { fontSize: lfs, bold: true, color: dark ? P.textOnDark : P.text, align: "center", valign: "bottom", pad: 3, autofit: "none", role: "orglabel" }));
        const subFs = node._composition ? fitFont(sub, boxW - 12, boxH * 0.45 - 4, fs, FONT.caption) : FONT.caption;
        c.prims.push(textBox(cx - boxW / 2, y + boxH * 0.55, boxW, boxH * 0.45, sub, { fontSize: subFs, color: dark ? P.textOnDark : P.textMuted, align: "center", valign: "top", pad: 2, autofit: "none", role: "caption" }));
      } else c.prims.push(textBox(cx - boxW / 2, y, boxW, boxH, label, { fontSize: fitFont(label, boxW - 8, boxH, fs, FONT.tableMin), bold: true, color: dark ? P.textOnDark : P.text, align: "center", valign: "middle", pad: 3, autofit: "none", role: "orglabel" }));
      const kids = nd.children || [];
      if (kids.length) {
        const busY = y + boxH + vGap / 2;
        c.prims.push(line(cx, y + boxH, cx, busY, P.line, 1));
        let kx = left;
        const centers = [];
        kids.forEach((k) => {
          const kw = widthOf(k);
          centers.push(kx + kw / 2);
          place(k, kx, level + 1);
          kx += kw + hGap;
        });
        if (centers.length > 1) c.prims.push(line(centers[0], busY, centers[centers.length - 1], busY, P.line, 1));
        centers.forEach((kcx) => c.prims.push(line(kcx, busY, kcx, y + boxH + vGap, P.line, 1)));
      }
    };
    place(root, ox, 0);
  };

  /** 因果の三角: 列間中央に平たい塗り三角 1 つ(ブロック矢印は使わない) */
  /** 横に伸びる樹形図(論点ツリー)。体制図と同じデータで、根を左に置き右へ展開する。
   *  深さより葉の数が増えやすい題材(論点分解・施策の優先度)は、縦積みより横展開のほうが 16:9 に収まる。 */
  LEAF.tree = function (node, r, c) {
    const P = c.P;
    const root = node.root;
    if (!root) return;
    const depth = orgDepth(root), leaves = orgLeaves(root);
    const hGap = 16, vGap = 8;
    // 列幅は等分ではなく、その階層の一番長いラベルから決める。余りは末端(内容量が一番多い列)へ回す
    const perLevel = [];
    (function walkLv(nd, lv) {
      (perLevel[lv] = perLevel[lv] || []).push(nd);
      (nd.children || []).forEach((k) => walkLv(k, lv + 1));
    })(root, 0);
    const needW = perLevel.map((nds) => Math.max.apply(null, nds.map((nd) => Math.max(measure(stripBold(nd.label || ""), FONT.body), measure(stripBold(nd.sub || ""), FONT.body)))) + 24);
    // 末端に補足説明があるときは、右にもう 1 列取る(任意。LLM が note を出したときだけ)
    const noted = [];
    (function walkN(nd) {
      const kids = nd.children || [];
      if (!kids.length && nd.note) noted.push(nd);
      kids.forEach(walkN);
    })(root);
    const hasNotes = noted.length > 0;
    const noteGap = 26; // 末端と説明のあいだは線で結ばず、余白で分ける
    const availW = r.w - hGap * (depth - 1) - (hasNotes ? noteGap : 0);
    const natural = needW.map((n) => Math.min(300, Math.max(72, n)));
    const sumW = natural.reduce((a, b) => a + b, 0);
    // 説明列は木の自然幅を引いた残り(180〜380pt)。説明があるときは木の箱を広げず、幅を説明に回す
    const noteW = hasNotes ? Math.max(180, Math.min(420, availW - sumW)) : 0;
    // 余りは 1 列に寄せず全列を共通の倍率で広げる。1 列だけ極端に広い箱になるのを防ぐ。
    // それでも余ったら木ごと中央に置く。入らないときは比例縮小する
    // 説明があるときは木の箱を広げない(余った幅はすべて説明に回す)
    const k = Math.min(hasNotes ? 1.0 : 1.8, (availW - noteW) / sumW);
    const colWs = natural.map((w2) => w2 * k);
    const used = colWs.reduce((a, b) => a + b, 0);
    const colX = [];
    colWs.reduce((cx, w2) => (colX.push(cx), cx + w2 + hGap), hasNotes ? r.x : r.x + Math.max(0, (availW - used) / 2));
    const noteX = hasNotes ? colX[depth - 1] + colWs[depth - 1] + noteGap : 0;
    const rowH = (r.h - vGap * (leaves - 1)) / leaves;
    if (rowH < 24) c.warnings.push("tree overflow: " + leaves + " leaves");
    const boxH = Math.max(24, Math.min(rowH, 72));
    const noteFs = hasNotes ? Math.min.apply(null, noted.map((nd) => fitFont(stripBold(nd.note), noteW - 12, rowH - 6, FONT.body, FONT.tableMin))) : 0;
    const heightOf = (nd) => (nd.children && nd.children.length ? nd.children.reduce((a, k) => a + heightOf(k), 0) + vGap * (nd.children.length - 1) : rowH);
    const place = (nd, top, level) => {
      const h = heightOf(nd), cy = top + h / 2, x = colX[level], colW = colWs[level];
      const dark = level === 0 || nd.highlight;
      const label = stripBold(nd.label || ""), sub = stripBold(nd.sub || "");
      // 縮んだことを prims にも残す。密度判定(全体の縮小・2 枚分割)はこの印を見るので、
      // ツリーだけ黙って 11pt まで縮むと「収まっている」と誤判定される
      // 背景と文字を別のシェイプに割らず、1 つの塗り付きシェイプに入れる(掴んで動かせる・編集しやすい)。
      // ラベルと補足の書き分けは font の範囲指定で行う
      const lfs = fitFont(label, colW - 10, sub ? boxH * 0.58 : boxH, FONT.body, FONT.tableMin);
      const sfs = sub ? fitFont(sub, colW - 12, boxH * 0.45 - 4, FONT.body, FONT.caption) : 0;
      const text = sub ? label + "\n" + sub : label;
      const styleRanges = sub ? [{ start: label.length + 1, len: sub.length, size: sfs, color: dark ? P.textOnDark : P.textMuted, bold: false }] : [];
      c.prims.push(rect(x, cy - boxH / 2, colW, boxH, { fill: dark ? P.fillDark : P.fillLight, text, boldRanges: [], styleRanges, fontSize: lfs, bold: true, color: dark ? P.textOnDark : P.text, align: "center", valign: "middle", pad: 4, autofit: "none", shrunk: lfs < FONT.min, role: "treelabel" }));
      const kids = nd.children || [];
      if (!kids.length) {
        if (hasNotes && nd.note) {
          c.prims.push(textBox(noteX, top, noteW, h, stripBold(nd.note), { fontSize: noteFs, color: P.text, align: "left", valign: "middle", pad: 6, autofit: "none", shrunk: noteFs < FONT.min, role: "treenote" }));
        }
        return;
      }
      const busX = x + colW + hGap / 2;
      c.prims.push(line(x + colW, cy, busX, cy, P.line, 1)); // 親の右から幹まで
      let ky = top;
      const centers = [];
      kids.forEach((k) => {
        const kh = heightOf(k);
        centers.push(ky + kh / 2);
        place(k, ky, level + 1);
        ky += kh + vGap;
      });
      if (centers.length > 1) c.prims.push(line(busX, centers[0], busX, centers[centers.length - 1], P.line, 1)); // 縦の幹
      centers.forEach((kcy) => c.prims.push(line(busX, kcy, x + colW + hGap, kcy, P.line, 1))); // 幹から子の左へ
    };
    place(root, r.y, 0);
  };
  LEAF.arrow = function (node, r, c) {
    const dir = node.direction || (node._orientation === "cols" ? "right" : "down");
    // アイコン的に正方形へ収まる小さな三角(細長い矢印にしない)
    const d = Math.max(14, Math.min(22, r.w, r.h));
    const cx = r.x + r.w / 2,
      cy = r.y + r.h / 2;
    c.prims.push(rect(cx - d / 2, cy - d / 2, d, d, { shape: "triangle", rotation: dir === "right" ? 90 : 180, fill: c.P.fillMid }));
  };

  // ---------- Figure(ネイティブ図形で描く基本チャート。数値は入力にあるものだけ) ----------
  /** 横棒: ラベルは折り返して必ず読める幅、基準線 1 本、値は右 */
  LEAF.bars = function (node, r, c) {
    const items = capList(node.items, 10, c, "棒");
    const n = items.length || 1;
    const unit = node.unit || "";
    const max = Math.max.apply(null, items.map((it) => Math.abs(Number(it.value) || 0)).concat([1]));
    const noteH = node.note ? 24 : 0;
    const rowH = Math.min(72, (r.h - noteH) / n);
    const barH = Math.min(rowH * 0.55, 30);
    // ラベル幅: 概算より実グリフが広いので 12% の余裕を持たせる(折り返して 2 行目が切れるのを防ぐ)
    const longest = Math.max.apply(null, items.map((it) => measure(stripBold(it.label || ""), FONT.body)).concat([60]));
    const labelW = Math.min(r.w * 0.46, Math.max(100, longest * 1.12 + PAD * 2 + 6));
    const valueW = Math.min(110, r.w * 0.16);
    const barAreaW = Math.max(40, r.w - labelW - valueW - GAP * 1.5);
    const top = r.y; // 上端揃え(隣のパネル見出しと上端を合わせる。中央に浮かせない)
    const x0 = r.x + labelW + GAP / 2;
    items.forEach((it, i) => {
      const y = top + i * rowH;
      const v = Number(it.value) || 0;
      const w = Math.max(2, (Math.abs(v) / max) * barAreaW);
      const hl = !!it.highlight;
      const lfs = fitFont(stripBold(it.label || ""), labelW - PAD * 2, rowH - 4, FONT.body, FONT.tableMin);
      c.prims.push(textBox(r.x, y, labelW, rowH, stripBold(it.label || ""), { fontSize: lfs, color: c.P.text, align: "right", valign: "middle", autofit: "none", wrap: false, role: "barlabel" }));
      c.prims.push(rect(x0, y + (rowH - barH) / 2, w, barH, { fill: hl ? c.P.accent : c.P.fillMid }));
      const valText = it.valueLabel != null ? String(it.valueLabel) : formatNum(v) + unit;
      c.prims.push(textBox(x0 + w + 4, y, valueW, rowH, valText, { fontSize: FONT.body, bold: hl, color: hl ? c.P.accent : c.P.text, valign: "middle", pad: 2, autofit: "none" }));
    });
    c.prims.push(line(x0, top, x0, top + rowH * n, c.P.line, 1));
    if (node.note) c.prims.push(textBox(r.x, r.y + r.h - noteH, r.w, noteH, node.note, { fontSize: FONT.footnote + 1, color: c.P.textMuted, align: "right", valign: "bottom", pad: 0, autofit: "none", role: "note" }));
  };

  /** KPI: 大きな数値 + ラベル。ブロック内で垂直中央 */
  LEAF.kpi = function (node, r, c) {
    const hl = node.highlight !== false;
    const dark = node.fill === "dark" || node.fill === "accent";
    if (dark) c.prims.push(rect(r.x, r.y, r.w, r.h, { fill: node.fill === "accent" && !c.P.mono ? c.P.accent : c.P.fillDark }));
    const labelH = FONT.body * FONT.lineHeight + 4;
    const valH = Math.min(FONT.kpi * 1.4, r.h - labelH);
    const fs = fitFont(String(node.value || ""), r.w - PAD * 2, valH, FONT.kpi, 28);
    const total = valH + labelH;
    const y0 = r.y + Math.max(0, (r.h - total) / 2);
    c.prims.push(textBox(r.x, y0, r.w, valH, String(node.value || ""), { fontSize: fs, bold: true, color: dark ? c.P.textOnDark : hl ? c.P.accent : c.P.text, align: "center", valign: "bottom", autofit: "none", role: "kpi" }));
    c.prims.push(textBox(r.x, y0 + valH, r.w, labelH, stripBold(node.label || ""), { fontSize: FONT.body, color: dark ? c.P.textOnDark : c.P.textMuted, align: "center", valign: "top", pad: 2, autofit: "none" }));
  };

  /** 縦棒グラフ: 量の差・カテゴリ比較 */
  LEAF.column = function (node, r, c) {
    const items = capList(node.items, 8, c, "棒");
    const n = items.length || 1;
    const unit = node.unit || "";
    const max = Math.max.apply(null, items.map((it) => Math.abs(Number(it.value) || 0)).concat([1]));
    const labelH = FONT.body * FONT.lineHeight * 2 + 4;
    const valueH = FONT.body * FONT.lineHeight;
    const noteH = node.note ? 22 : 0;
    const plotTop = r.y + valueH;
    const plotH = r.h - valueH - labelH - noteH - 6;
    const slot = r.w / n;
    const bw = Math.min(slot * 0.55, 72);
    items.forEach((it, i) => {
      const v = Number(it.value) || 0;
      const h = Math.max(2, (Math.abs(v) / max) * plotH);
      const x = r.x + slot * i + (slot - bw) / 2;
      const y = plotTop + plotH - h;
      const hl = !!it.highlight;
      c.prims.push(rect(x, y, bw, h, { fill: hl ? c.P.accent : c.P.fillMid }));
      c.prims.push(textBox(r.x + slot * i, y - valueH, slot, valueH, formatNum(v) + unit, { fontSize: FONT.body, bold: hl, color: hl ? c.P.accent : c.P.text, align: "center", valign: "bottom", pad: 0, autofit: "none" }));
      c.prims.push(textBox(r.x + slot * i, plotTop + plotH + 4, slot, labelH, stripBold(it.label || ""), { fontSize: FONT.body, color: c.P.text, align: "center", valign: "top", pad: 2, autofit: "none", role: "barlabel" }));
    });
    c.prims.push(line(r.x, plotTop + plotH, r.x + r.w, plotTop + plotH, c.P.line, 1));
    if (node.note) c.prims.push(textBox(r.x, r.y + r.h - noteH, r.w, noteH, node.note, { fontSize: FONT.footnote + 1, color: c.P.textMuted, align: "right", valign: "bottom", pad: 0, autofit: "none", role: "note" }));
  };

  /** 折れ線グラフ: 推移(最大 2 系列)。線は回転した細い矩形、点は小円 */
  LEAF.line = function (node, r, c) {
    const labels = capList(node.labels, 12, c, "系列の点");
    let series = capList(node.series, 2, c, "系列");
    let droppedNote = "";
    if (series.length === 2) {
      const mx = (sv) => Math.max.apply(null, (sv.values || []).map(Number).filter((v) => !isNaN(v)).concat([0]));
      const a = mx(series[0]), b = mx(series[1]);
      const small = a > 0 && b > 0 && (a / b >= 10 || b / a >= 10) ? (a > b ? 1 : 0) : -1;
      if (small >= 0) {
        // 桁が違う系列は同じ軸に載せない。値は注記で残す(情報は落とさない)
        const sv = series[small];
        const vals = (sv.values || []).map(Number).filter((v) => !isNaN(v));
        droppedNote = stripBold(sv.label || "") + ": " + (vals.length ? vals[0] + " → " + vals[vals.length - 1] : "");
        series = [series[1 - small]];
      }
    }
    const unit = /\d/.test(node.unit || "") || (node.unit || "").length > 6 ? "" : node.unit || ""; // 「2019年比」のような単位もどきは使わない
    const n = Math.max(2, labels.length);
    const all = series.flatMap((s) => (s.values || []).map(Number)).filter((v) => !isNaN(v));
    const max = Math.max.apply(null, all.concat([1]));
    const dataMin = Math.min.apply(null, all.concat([max]));
    // 値が 0 から遠い(最大/最小 < 2.5)ときは 0 基線にせず、下に少し余白を取った基線にする(版面が空かないように)
    const min = dataMin > 0 && max / dataMin < 2.5 ? dataMin - Math.max(max - dataMin, max * 0.05) * 0.35 : Math.min(0, dataMin);
    const labelH = FONT.body * FONT.lineHeight + 6;
    const legendH = series.length > 1 ? FONT.body * FONT.lineHeight + 6 : 0;
    const valueH = FONT.body * FONT.lineHeight;
    const noteH = droppedNote ? FONT.body * FONT.lineHeight + 8 : 0; // 注記の行を先に確保する(最終値のラベルと重ねない)
    const plotTop = r.y + valueH + legendH + noteH;
    const plotH = r.h - valueH - legendH - labelH - noteH - 8;
    // 端のラベル(最初と最後)が枠外に出ないだけの余白を先に確保する。
    // 後からクランプすると、ラベルの中心が点からずれて「値と目盛りが合っていない」ように見える
    const labelHalf = labels.length ? Math.max.apply(null, labels.map((lb) => measure(stripBold(lb), FONT.body))) / 2 : 0;
    const padX = Math.max(Math.min(40, r.w * 0.06), Math.min(r.w * 0.16, labelHalf + 4));
    const stepX = (r.w - padX * 2) / (n - 1);
    const yOf = (v) => plotTop + plotH - ((v - min) / (max - min || 1)) * plotH;
    if (legendH) {
      let lx = r.x;
      series.forEach((s, si) => {
        const col = si === 0 ? c.P.accent : c.P.fillDark;
        c.prims.push(rect(lx, r.y + 6, 14, 14, { fill: col }));
        const lw = measure(stripBold(s.label || ""), FONT.body) + 24;
        c.prims.push(textBox(lx + 16, r.y, lw, legendH, stripBold(s.label || ""), { fontSize: FONT.body, color: c.P.text, valign: "middle", pad: 2, autofit: "none" }));
        lx += 16 + lw + 12;
      });
    }
    series.forEach((s, si) => {
      const col = series.length === 1 || s.highlight ? c.P.accent : si === 0 ? c.P.accent : c.P.fillDark;
      const vals = (s.values || []).slice(0, n).map(Number);
      for (let i = 0; i < vals.length - 1; i++) {
        const x1 = r.x + padX + stepX * i,
          y1 = yOf(vals[i]),
          x2 = r.x + padX + stepX * (i + 1),
          y2 = yOf(vals[i + 1]);
        const len = Math.hypot(x2 - x1, y2 - y1);
        const ang = (Math.atan2(y2 - y1, x2 - x1) * 180) / Math.PI;
        c.prims.push(rect((x1 + x2) / 2 - len / 2, (y1 + y2) / 2 - 1.25, len, 2.5, { fill: col, rotation: ang }));
      }
      vals.forEach((v, i) => {
        const x = r.x + padX + stepX * i,
          y = yOf(v);
        c.prims.push(rect(x - 4, y - 4, 8, 8, { shape: "ellipse", fill: col }));
        // 2 系列あるときは主系列だけに値を書く(副系列のラベルが軸ラベルと衝突するのを防ぐ)
        const showVal = series.length === 1 ? true : si === 0 && (i === 0 || i === vals.length - 1);
        if (showVal) {
          const vt = formatNum(v) + unit;
          const vw = Math.max(80, measure(vt, FONT.body) * 1.2 + 16); // 単位付き(例「2,160 件/月」)でも切れない幅
          c.prims.push(textBox(Math.max(r.x, Math.min(x - vw / 2, r.x + r.w - vw)), Math.max(r.y, y - valueH - 4), vw, valueH, vt, { fontSize: FONT.body, bold: i === vals.length - 1, color: col, align: "center", valign: "bottom", pad: 0, autofit: "none", role: "barlabel" }));
        }
      });
    });
    c.prims.push(line(r.x + padX - 10, plotTop + plotH, r.x + r.w - padX + 10, plotTop + plotH, c.P.line, 1));
    if (droppedNote) c.prims.push(textBox(r.x, r.y, r.w, FONT.body * FONT.lineHeight + 4, droppedNote, { fontSize: Math.max(8, FONT.body - 2), color: c.P.textMuted, align: "right", valign: "top", pad: 2, autofit: "none", role: "note" }));
    labels.forEach((lb, i) => {
      const x = r.x + padX + stepX * i; // 点と同じ x
      const lw = Math.min(stepX, padX * 2); // 隣と重ならず、かつ点を中心に置ける幅
      c.prims.push(textBox(x - lw / 2, plotTop + plotH + 4, lw, labelH, stripBold(lb), { fontSize: FONT.body, color: c.P.text, align: "center", valign: "top", pad: 0, autofit: "none", role: "barlabel" }));
    });
  };

  /** 100% 積み上げ横棒: 構成比(円グラフの代わり) */
  LEAF.stacked = function (node, r, c) {
    const items = capList(node.items, 6, c, "積み上げの項目");
    const n = items.length || 1;
    const parts = node.parts || Array.from(new Set(items.flatMap((it) => (it.values || []).map((_, i) => i)))).map((i) => "系列" + (i + 1));
    const cols = [c.P.accent, c.P.fillDark, c.P.line, c.P.fillMid, c.P.fillLight, c.P.lineLight];
    const legendH = FONT.body * FONT.lineHeight + 8;
    const labelW = Math.min(r.w * 0.3, Math.max(100, Math.max.apply(null, items.map((it) => measure(stripBold(it.label || ""), FONT.body))) + PAD * 2));
    const barX = r.x + labelW + GAP / 2;
    const barW = r.w - labelW - GAP / 2;
    const rowH = Math.min(56, (r.h - legendH) / n);
    const barH = Math.min(rowH * 0.6, 32);
    const top = r.y + legendH; // 上端揃え
    let lx = r.x;
    parts.forEach((p, pi) => {
      c.prims.push(rect(lx, r.y + 5, 14, 14, { fill: cols[pi % cols.length] }));
      const lw = measure(stripBold(p), FONT.body) + 20;
      c.prims.push(textBox(lx + 16, r.y, lw, legendH, stripBold(p), { fontSize: FONT.body, color: c.P.text, valign: "middle", pad: 2, autofit: "none" }));
      lx += 16 + lw + 14;
    });
    items.forEach((it, i) => {
      const y = top + i * rowH;
      const vals = (it.values || []).map((v) => Math.max(0, Number(v) || 0));
      const sum = vals.reduce((a, b) => a + b, 0) || 1;
      c.prims.push(textBox(r.x, y, labelW, rowH, stripBold(it.label || ""), { fontSize: FONT.body, color: c.P.text, align: "right", valign: "middle", autofit: "none", role: "barlabel" }));
      let x = barX;
      vals.forEach((v, pi) => {
        const w = (v / sum) * barW;
        c.prims.push(rect(x, y + (rowH - barH) / 2, Math.max(0.5, w), barH, { fill: cols[pi % cols.length] }));
        if (w > 44) {
          const dark = pi >= 2;
          c.prims.push(textBox(x, y + (rowH - barH) / 2, w, barH, Math.round((v / sum) * 100) + "%", { fontSize: 16, color: dark ? c.P.text : c.P.textOnDark, align: "center", valign: "middle", pad: 0, autofit: "none", role: "barlabel" }));
        }
        x += w;
      });
    });
  };

  // =====================================================================
  //  spec の正規化(LLM 出力のゆらぎ・旧形式を吸収し、cell/grid/table/figure/arrow に脱糖する)
  // =====================================================================
  // A layer shape and its explanation share a row; there is no nested panel.
  LEAF.pyramid = function (node, r, c) {
    const n = node.layers.length, gap = 10;
    const rowH = (r.h - gap * (n - 1)) / n;
    const figureW = r.w * 0.43, textX = r.x + r.w * 0.54, textW = r.w * 0.46;
    node.layers.forEach((layer, i) => {
      const y = r.y + i * (rowH + gap), w = figureW * (i + 1) / n;
      const cx = r.x + figureW / 2;
      c.prims.push(rect(cx - w / 2, y, w, rowH, { shape: i === 0 ? "triangle" : "trapezoid", fill: [c.P.fillLight, c.P.fillMid, c.P.line][Math.min(2, Math.floor(i * 3 / n))] }));
      const labelY = i === 0 ? y + rowH * 0.5 : y + rowH * 0.25;
      const labelH = i === 0 ? rowH * 0.48 : rowH * 0.5;
      const fs = fitFont(layer.head, Math.max(20, w * 0.72), labelH, FONT.head, FONT.tableMin);
      c.prims.push(textBox(cx - w * 0.36, labelY, w * 0.72, labelH, layer.head, { fontSize: fs, bold: true, color: c.P.text, align: "center", valign: "middle", pad: 0, role: "layerhead" }));
      c.prims.push(line(r.x + figureW + 12, y + rowH / 2, textX - 16, y + rowH / 2, c.P.line, 1));
      const bodyFs = fitFont(layer.text, textW - 8, rowH - 8, FONT.body, FONT.floor);
      if (textHeight(layer.text, textW - 8, bodyFs) > rowH - 8) c.warnings.push("pyramid overflow: split long layer descriptions");
      c.prims.push(textBox(textX, y, textW, rowH, layer.text, { fontSize: bodyFs, color: c.P.text, valign: "middle", pad: 4, role: "layertext" }));
    });
  };
  const LEAF_TYPES = ["cell", "table", "ntable", "bars", "column", "line", "stacked", "kpi", "arrow", "gantt", "org", "tree", "pyramid"];
  const FILLS = ["none", "light", "dark", "tint", "accent"];
  const WRAP_TYPES = ["cell", "table", "ntable", "matrix", "bars", "column", "line", "stacked", "kpi", "arrow", "chevrons", "figure", "card", "bullets", "callout", "label", "icon", "gantt", "org", "tree"];

  // Generation grammar: one visual, or 1–3 equal text panels. Legacy normalization
  // below is retained only for old saved specs; every /api/generate response uses this.
  /** 「**ラベル**：値」が 1 行に入らず、かつラベル列に割るには枠が狭いとき、語の途中ではなく
   *  ラベルの直後で折る。行数が増えない場合だけ採用する(「自社の運営要 / 因」のような切れ方を防ぐ) */
  function breakAtLabel(text, w, fs) {
    const t = String(text || "");
    if (!t || t.includes("\n")) return t;
    const before = estimateLines(t, w, fs);
    if (before < 2) return t;
    const m = t.match(/^(\s*(?:\*\*[^*]+\*\*|[^:：\n]{2,16})\s*[:：])(\s*)(\S[\s\S]*)$/);
    if (!m) return t;
    const head = m[1], rest = m[3];
    if (estimateLines(head, w, fs) > 1) return t;
    if (estimateLines(head, w, fs) + estimateLines(rest, w, fs) > before) return t;
    return head + "\n" + rest;
  }
  function renderNestedItems(node, r, c) {
    const rows = node.items.flatMap((v) => Array.isArray(v) ? v.map((t) => ({ text: itemToText(t), level: 1 })) : [{ text: itemToText(v), level: 0 }]);
    const labelledOutline =
      node._labelledOutline != null ? node._labelledOutline : !rows.some((v) => v.level) && rows.every((v) => /^\*\*[^*]+\*\*\s*[:：]/.test(v.text));
    const markerW = labelledOutline ? 0 : BULLET_INDENT;
    // ラベル付きの列挙(**ラベル**: 値)は、ラベル列を左に揃えて値だけを折り返す。
    // 1 つの文として流すと折り返した 2 行目がラベルの下に潜り、ラベルと値の境目が読めなくなる。
    // 幅が足りない(値が 1 行 12 字未満になる)ときは列に割らず、従来どおり 1 つの文として流す。
    const pairs = labelledOutline
      ? rows.map((row) => {
          const m = /^\*\*([^*]+)\*\*\s*[:：]?\s*([\s\S]*)$/.exec(row.text);
          const label = m ? m[1].trim() : "";
          const value = m ? m[2].trim() : "";
          return label && value ? { label, value } : null;
        })
      : null;
    const labelled = !!pairs && pairs.every(Boolean);
    // **箇条書きを 1 行ずつ別のシェイプにしない。** 生成後にユーザーが箱を選び、Enter で項目を足し引きする
    // のが実際の使い方で、行ごとに箱があるとその操作ができない(管理が破綻する)。
    // Office.js の ParagraphFormat は horizontalAlignment と indentLevel しか持たず(実機で確認)、
    // 段落間隔・行間・ぶら下げも、段落ごとの字下げも無い。1 シェイプでできることは限られるが、
    // それは編集容易性より優先しない。兄弟パネルの行揃えもこのために捨てた(パネルごとに 1 シェイプ)。
    // 余った高さは縦中央に置いて、下半分だけが空く見え方を避ける。
    if (!labelled) {
      const body = cellBody(node); // 入れ子は子をソフト改行で親の段落に入れて 1 シェイプに畳む
      const tw = r.w - PAD * 2 - (body.list ? BULLET_INDENT : 0);
      const fs = fitBody(c, body.text, tw, r.h - PAD * 2, Math.min(FONT.large, FONT.body + 4), 0, undefined, body.list);
      const txt = nestedText(body, tw, fs, r.h - PAD * 2); // 記号と塊の空行は、実際の幅・文字サイズ・高さで決める
      const need = textHeight(txt, tw, fs) + PAD * 2;
      c.prims.push(
        textBox(r.x, r.y, r.w, r.h, txt, { fontSize: fs, color: c.P.text, bullets: body.list, valign: "middle", pad: PAD, gid: node._gid, shrunk: fs < FONT.min })
      );
      if (need > r.h) c.warnings.push("nested list overflow");
      return;
    }
    const LABEL_GAP = 10;
    // 文字幅は近似なので、実フォントで 1 文字はみ出してラベルが折り返さないよう estimateLines と同じ安全率を戻す
    const labelColW = (fs) => Math.min((r.w - PAD * 2) * 0.4, Math.max.apply(null, pairs.map((p) => measure(p.label, fs))) / 0.95 + LABEL_GAP);
    const metrics = (fs) => {
      const lw = labelled ? labelColW(fs) : 0;
      const cols = labelled && r.w - PAD * 2 - lw >= fs * 12;
      return rows.map((row, i) => {
        const indent = row.level ? SPACE.levelIndent + 8 : 0;
        const fontSize = row.level ? Math.max(FONT.floor, fs - 2) : fs;
        const labelW = cols ? lw : 0;
        const w = r.w - PAD * 2 - indent - markerW - labelW;
        const text = cols ? pairs[i].value : breakAtLabel(row.text, w, fontSize);
        // 分けたときも 1 シェイプの箇条書きと同じ字送りにする(項目間に余白を足さない)。
        // 足せるのは分けたときだけなので、足すと 1 シェイプで描く箇条書きとの間で見た目が食い違う
        return { ...row, text, label: cols ? pairs[i].label : null, labelW, indent, fontSize, w, h: textHeight(text, w, fontSize), gap: i === rows.length - 1 ? 0 : rows[i + 1].level === 0 ? 0 : 3 };
      });
    };
    let fs = Math.min(FONT.large, FONT.body + 4), items = metrics(fs);
    const sum = (xs) => xs.reduce((n, x) => n + x.h + x.gap, 0);
    while (sum(items) > r.h - PAD * 2 && fs > FONT.floor) items = metrics(--fs);
    if (fs < FONT.min) c.warnings.push("shrunk nested list to " + fs + "pt");
    if (sum(items) > r.h) c.warnings.push("nested list overflow");
    // 余りが大きいときは等高の行の格子にし、文字は行の中で縦中央、行の下に薄い罫を引く。
    // 境界を作らずに間隔だけ広げると、見出し直下に穴が空いて版面のバランスが崩れる(実測で 56〜84pt の穴)。
    // 行の高さは兄弟パネルと共有するので、段数が違っても罫の y が揃う。
    const gridRows = Math.max(items.length, node._siblingRows || 0);
    // 行の高さは版面を割り切った値ではなく「自然な字送りを行数単位に丸めた値」を上限にする。
    // 割り切った値をそのまま使うと、疎なパネルで行が版面いっぱいに散り、1 シェイプで描く
    // 箇条書きより明らかに間延びして見える(同じスライドの中で間隔が食い違う)。
    // 行数単位に丸めるので、兄弟パネルの行数が同じなら左右の行は揃ったままになる
    const lineH = fs * FONT.lineHeight;
    const naturalRowH = Math.max.apply(null, items.map((row) => row.h).concat([lineH]));
    const gridRowH = Math.min((r.h - PAD * 2) / Math.max(1, gridRows), Math.ceil(naturalRowH / lineH) * lineH);
    // 兄弟パネルがあるときは段数が違っても必ず格子にする(片方だけ格子になると行が揃わず、かえって崩れて見える)。
    // 単独のパネルは格子にしない(そもそもここへは来ない)。余白のために行を広げると、
    // 1 シェイプで描く箇条書きとの間で間隔が食い違う。疎な版面は文字サイズと縦中央で埋める。
    const useGrid = node._siblingRows > 0;
    // 行の高さは自然な字送りそのものなので、収まり判定は等号まで許す(-4 のままだと必ず外れて
    // 兄弟パネルが揃わない流し込みに落ちる)
    if (items.length >= 2 && !rows.some((v) => v.level) && useGrid && items.every((row) => row.h <= gridRowH + 0.5)) {
      let gy = r.y + Math.max(PAD, (r.h - gridRowH * gridRows) / 2); // 縮めた分は上下に均等に返す(兄弟パネルで同じ値になる)
      items.forEach((row) => {
        const gx = r.x + PAD + row.indent;
        // 点は PowerPoint 本来の箇条書き(paragraphFormat.bulletFormat)に描かせる。
        // 「・」を別の図形として置くと、行の高さや折り返しのたびに点だけ位置がずれる。
        if (row.label) c.prims.push(textBox(gx, gy, row.labelW - 2, gridRowH, row.label, { fontSize: row.fontSize, bold: true, color: c.P.text, pad: 0, valign: "middle", role: "listlabel", shrunk: fs < FONT.min }));
        c.prims.push(textBox(gx + row.labelW, gy, row.w + markerW, gridRowH, row.text, { fontSize: row.fontSize, color: c.P.text, pad: 0, valign: "middle", bullets: markerW > 0, role: "listitem", shrunk: fs < FONT.min }));
        gy += gridRowH;
      });
      return;
    }
    let y = r.y + Math.max(PAD, (r.h - sum(items)) / 2);
    for (const row of items) {
      const x = r.x + PAD + row.indent;
      if (row.label) c.prims.push(textBox(x, y, row.labelW - 2, row.h, row.label, { fontSize: row.fontSize, bold: true, color: c.P.text, pad: 0, role: "listlabel", shrunk: fs < FONT.min }));
      c.prims.push(textBox(x + row.labelW, y, row.w + markerW, row.h, row.text, { fontSize: row.fontSize, color: c.P.text, pad: 0, bullets: markerW > 0, role: "listitem", shrunk: fs < FONT.min }));
      y += row.h + row.gap;
    }
  }
  function unnumberHead(value) {
    let text = String(value || "").trim(), prior;
    do {
      prior = text;
      text = text.replace(/^(?:(?:step|ステップ)\s*\d{1,2}\s*[:：.)\-]?\s*|第\s*\d{1,2}\s*(?:段階|工程|ステップ)\s*[:：.)\-]?\s*|\d{1,2}[.．、):：\-]\s*|\d{1,2}\s+|[①-⑳]\s*)/i, "").trim();
    } while (text !== prior);
    return text;
  }
  /** note が title / lead の言い換えなら true。完全一致・部分一致に加え、内容語の重なりでも判定する
   *  (「表の下にリード文と同じことが書いてある」状態は、言い換えられていても読み手には重複でしかない) */
  function repeatsSummary(note, summary) {
    const key = (v) => stripBold(String(v || "")).normalize("NFKC").replace(/[\s、。，,.：:！!？?]/g, "").replace(/(?:である|です|すること)$/u, "");
    const a = key(note), b = key(summary);
    if (!a) return false;
    if (a === b || (a.length >= 12 && b.includes(a))) return true;
    // 内容語(2 文字以上の漢字・カタカナ・英数の連なり)の重なりが 6 割を超えたら言い換えとみなす
    const words = (v) => new Set((String(v || "").match(/[一-龥ァ-ヶA-Za-z0-9][一-龥ァ-ヶーA-Za-z0-9.,%]+/g) || []).filter((w) => w.length >= 2));
    const wa = words(note), wb = words(summary);
    if (wa.size < 3) return false;
    let hit = 0;
    wa.forEach((w) => { if (wb.has(w)) hit++; });
    return hit / wa.size >= 0.6;
  }
  function compositionKey(specs) {
    return (Array.isArray(specs) ? specs : [specs]).map((s) => {
      if (s.panelCount > 1) return "panels:" + s.panelCount;
      const b = s.body || {};
      return b.type === "ntable" ? "table" : b.type || "legacy";
    }).join("/");
  }
  function normalizeGeneratedSpec(raw) {
    const s = raw || {};
    const plain = (v) => stripBold(str(v));
    const repairs = Array.isArray(s.compositionRepairs) ? s.compositionRepairs.slice() : [];
    const neutral = (v) => {
      if (typeof v === "string") return v;
      if (Array.isArray(v)) return v.map(neutral);
      if (!v || typeof v !== "object") return v;
      const o = {};
      for (const [k, x] of Object.entries(v)) {
        if (["highlight", "highlightCol", "fill", "bold", "style", "icon", "size", "align"].includes(k) || k.startsWith("_")) continue;
        o[k] = neutral(x);
      }
      return o;
    };
    // A structural repair must keep labels, quantities and relationships, even when
    // a legacy complex block has to become plain text inside a panel.
    const paragraphs = (n) => {
      if (!n) return [];
      if (typeof n !== "object") return [plain(n)].filter(Boolean);
      if (Array.isArray(n)) return n.flatMap(paragraphs);
      if (!n.type && (n.rows || n.cols)) return (n.rows || n.cols).flatMap(paragraphs);
      const prefix = [n.caption, n.head].filter(Boolean).map(plain).join(" · ");
      if (n.type === "table" || n.type === "ntable" || n.type === "matrix") {
        return (n.rows || []).map((r) => [plain(r.head), ...(r.cells || []).map((c, j) => [plain((n.colHeaders || [])[j]), plain(cellText(c))].filter(Boolean).join(": "))].filter(Boolean).join(" / "));
      }
      if (n.type === "gantt") return (n.tasks || []).map((t) => `${plain(t.label)}: ${typeof t.start === "number" ? n.periods[t.start] : t.start}–${typeof t.end === "number" ? n.periods[t.end] : t.end}${t.note ? " / " + plain(t.note) : ""}`).concat((n.milestones || []).map((m) => `${plain(m.label)}: ${typeof m.at === "number" ? n.periods[m.at] : m.at}`));
      if (n.type === "org" || n.type === "tree") {
        const tree = (v, path) => v ? [[...path, plain(v.label)].join(" → ") + (v.sub ? ": " + plain(v.sub) : ""), ...(v.children || []).flatMap((k) => tree(k, [...path, plain(v.label)]))] : [];
        return tree(n.root, []);
      }
      if (n.series) return n.series.map((v) => plain(v.label) + ": " + (v.values || []).map((x, i) => `${(n.labels || [])[i] || i + 1} ${x}${n.unit || ""}`).join(" / "));
      if (n.value != null) return [`${plain(n.label)}: ${plain(n.value)}${n.unit || ""}`];
      const entries = (n.items || []).flatMap((v) => Array.isArray(v) ? v.map(plain) : typeof v === "object" ? [plain(v.label) + ": " + (v.values || [v.value]).join(" / ") + (n.unit || "")] : [plain(v)]);
      return [prefix, plain(n.text), ...entries].filter(Boolean);
    };
    // 図 + 意味合いの 2 パネル。ガント・体制図・密な表は幅がいるので単独パネル専用にする
    const PANEL_FIGURES = ["line", "bars", "column", "stacked", "kpi", "table", "matrix", "pyramid"];
    const panel = (n) => {
      // 正規化は冪等にする。描画側(renderspec トリガ・案の差し替え)は正規化済みの spec をもう一度通すので、
      // ここで図のパネルを文章に畳んでしまうと図が消える
      if (n && n._figurePanel) return n;
      const fig = n && n.body && n.body.type && PANEL_FIGURES.includes(n.body.type) ? n.body : null;
      if (fig) {
        const v = normalizeVisual(fig);
        v.head = plain(n.head);
        v._figurePanel = true;
        return v;
      }
      n = neutral(n);
      const source = n && !n.type && n.rows ? n.rows : [n];
      const first = source[0] || {};
      const head = plain(n.head || first.head || first.caption);
      const rest = source.map((x, i) => i === 0 && x === first ? Object.assign({}, x, { head: "", caption: "" }) : x);
      const items = rest.flatMap((x) => x && (!x.type || x.type === "cell") && !x.rows && !x.cols
        ? [x.caption, x.head, x.text, ...(x.items || [])].filter(Boolean) : paragraphs(x));
      const normalizedItems = [];
      for (const src of items) {
        // {"label":…,"value":…} の構造を落とさずに「**ラベル**: 値」へ畳む(描画側でラベル列を揃える)
        const item = Array.isArray(src) ? src.map(labelValueText) : labelValueText(src);
        if (typeof item === "string" && /^\s+[-•–]\s+/.test(item) && normalizedItems.length) {
          if (!Array.isArray(normalizedItems[normalizedItems.length - 1])) normalizedItems.push([]);
          normalizedItems[normalizedItems.length - 1].push(item.replace(/^\s*[-•–]\s+/, ""));
        } else normalizedItems.push(Array.isArray(item) ? item.map((x) => str(x).replace(/^[-•–]\s+/, "")) : str(item).replace(/^[-•–]\s+/, ""));
      }
      return { type: "cell", head, items: normalizedItems, valign: "middle" };
    };
    const seq = (n) => {
      const steps = (n.steps || n.items || []).map((x) => ({ head: unnumberHead(plain(x.head || x.label || x.title)), ...(typeof x.icon === "string" && /^[a-z][a-z0-9-]{0,63}$/.test(x.icon) ? { icon: x.icon } : {}), text: [str(x.text || x.description), ...(x.items || x.bullets || []).flat().map(str)].filter(Boolean).join("\n") }));
      if (steps.length < 2 || steps.length > 8 || steps.some((x) => !x.head)) throw new Error("ステップは見出し付きで 2〜8 段にしてください。");
      return { type: "sequence", steps, ...(n.headShape === "chevron" ? { headShape: "chevron" } : {}) };
    };
    const normalizeVisual = (n) => {
      if (["sequence", "steps", "process", "chevrons"].includes(n.type)) return seq(n);
      if (n.type === "pyramid") {
        if (!Array.isArray(n.layers) || n.layers.length < 2 || n.layers.length > 5) throw new Error("ピラミッドは 2〜5 層にしてください。");
        const layers = n.layers.map((v) => ({ head: plain(v.head), text: str(v.text) }));
        if (layers.some((v) => !v.head || !v.text)) throw new Error("各層に見出しと説明を指定してください。");
        return { type: "pyramid", layers };
      }
      n = neutral(n);
      if (["table", "matrix", "ntable"].includes(n.type)) {
        const t = normTable(n, n.type === "ntable" ? "ntable" : "table");
        const nc = Math.max(t.colHeaders.length, ...t.rows.map((r) => r.cells.length), 0);
        if (!t.rows.length || !nc) throw new Error("マトリクスには行とセルが必要です。");
        if (t.rows.length > 12 || nc > 6) throw new Error("表が大きすぎます。12 行・6 列以内でスライドを分けてください。");
        if (t.type === "table" && !t.headShape && !t.axes && !t.colGroups && (t.rows.length >= 8 || t.rows.length * nc >= 36)) t.type = "ntable";
        if (t.type === "ntable" && t.colHeaders.length < nc) t.colHeaders = Array.from({ length: nc }, (_, j) => t.colHeaders[j] || "");
        if (t.type === "table") {
          // Suppress small-axis header bands, but retain supplied meaning inline.
          // 見出しをセル文字列へ畳み込むときも .items(箇条書き版)は残す(text だけに見出しを足す)
          const mergeLabel = (v, label) =>
            v && typeof v === "object" && Array.isArray(v.items) && v.items.length ? Object.assign({}, v, { text: label + cellText(v) }) : label + cellText(v);
          if (nc <= 2 && !t.colGroups && !t.axes) {
            t.rows.forEach((r) => { r.cells = r.cells.map((v, j) => (t.colHeaders[j] ? mergeLabel(v, t.colHeaders[j] + ": ") : v)); });
            t.colHeaders = [];
          }
          if (t.rows.length <= 2 && !t.headShape && !t.axes) {
            t.rows.forEach((r) => { if (r.head) r.cells = r.cells.map((v) => mergeLabel(v, r.head + ": ")); r.head = ""; });
            t.corner = "";
          }
          t.headFill = "light";
          if (t.headShape || t.numbered) {
            t.rows.forEach((r) => { r.head = unnumberHead(r.head); });
            if (t.headShape) t.numbered = true;
          }
        }
        return t;
      }
      const v = normNode(n, 0, { leaves: 0 });
      if (!v || isContainer(v)) throw new Error("1 パネルには主となる図を 1 つ指定してください。");
      if (v.type === "gantt" && v.groups) {
        let remaining = v.periods.length;
        v.groups = v.groups.flatMap((g) => {
          const span = Math.min(remaining, g.span);
          remaining -= span;
          return span > 0 ? [{ text: g.text, span }] : [];
        });
        if (remaining > 0) v.groups.push({ text: "", span: remaining });
      }
      return neutral(v);
    };
    const b = s.body || {};
    let candidates = s.panels || (b.cols && b.cols.filter((x) => x.type !== "arrow"));
    const stepRow = candidates && candidates.length && candidates.every((x) => x.type === "cell" && ["step", "home", "chevron"].includes(x.shape));
    const inferred = candidates && !stepRow && candidates.length >= 2 && candidates.length <= 3 ? candidates.length : 1;
    const count = s.panelCount == null ? inferred : Number(s.panelCount);
    if (![1, 2, 3].includes(count)) throw new Error("panelCount は 1・2・3 のいずれかです。");
    const out = { title: cleanTitle(plain(s.title || s.headline)), lead: plain(s.lead || s.message || s.subtitle), footnote: plain(s.footnote || s.source), kicker: plain(s.kicker), panelCount: count, compositionVersion: 1, body: null };
    if (["auto", "matrix", "outline", "relational"].includes(s.layoutIntent)) out.layoutIntent = s.layoutIntent;
    if (count > 1) {
      if (!candidates || candidates.length !== count) throw new Error("パネル数と panels の要素数が一致していません。");
      out.body = { cols: candidates.map(panel) };
      if (out.body.cols.filter((p) => p._figurePanel).length > 1) throw new Error("図は 1 枚に 1 つです。図を横に 2 つ並べず、片側は文章にしてください。");
      if (out.body.cols.some((p) => !p.head)) throw new Error("各パネルにはサブタイトルが必要です。");
      if (s.note) throw new Error("複数パネルの下に補足ブロックは置けません。該当パネルの本文に含めてください。");
    } else {
      let main = s.panels && candidates.length === 1 ? Object.assign({ type: "cell" }, candidates[0]) : b;
      let notes = s.note ? paragraphs(neutral(s.note)) : [];
      if (b.note && !["bars", "column", "line", "stacked"].includes(b.type)) notes.push(...paragraphs(neutral(b.note)));
      if (stepRow) main = { type: "sequence", steps: candidates };
      else if (isContainer(b)) {
        // Old free composition: promote one visual to full width, merge text
        // supplements into one unboxed note. Never silently drop a second visual.
        const leaves = [];
        const collect = (x) => {
          if (isContainer(x) && x.cols && x.cols.every((k) => k.type === "cell" && k.shape)) leaves.push({ type: "sequence", steps: x.cols });
          else if (isContainer(x)) (x.rows || x.cols).forEach(collect);
          else if (x && x.type !== "arrow") leaves.push(x);
        };
        collect(b);
        const visuals = leaves.filter((x) => x.type !== "cell");
        if (visuals.length > 1) throw new Error("1 スライドに独立した図が複数あります。図ごとにスライドを分けてください。");
        if (visuals.length) { main = visuals[0]; notes.push(...leaves.filter((x) => x.type === "cell").flatMap((x) => paragraphs(neutral(x)))); }
        else main = { type: "cell", items: leaves.flatMap((x) => paragraphs(neutral(x))) };
        repairs.push("free composition consolidated");
      }
      if (main.type === "cell" || !main.type) out.body = panel(main);
      else out.body = normalizeVisual(main);
      // A note is a single full-width text box; explicit identical repeats vanish.
      notes = [...new Set(notes)].filter((x) => !repeatsSummary(x, out.title) && !repeatsSummary(x, out.lead));
      if (notes.length) out.note = { type: "cell", text: notes.join("\n") };
    }
    if (repairs.length) out.compositionRepairs = repairs;
    return out;
  }

  /**
   * composition の後処理(方法論の決定論的な適用):
   * 1 パネルの箇条書きが「ラベル: 説明」の並列列挙(4 件以上)なら、行名の箱 + 横長の説明の格子に組み替える。
   */
  /**
   * 「外部要因 30%: …。内部要因 50%: …。その他 20%: …。」のように、
   * ラベル付きの文が 3 つ以上ひと塊で返ってきたときに箇条書きへ割る。
   * 同じ入力でも LLM が items で返したり text で返したりするので、見た目を揃えるために正規化する。
   */
  function splitLabelledSentences(t) {
    const raw = String(t || "").trim();
    if (raw.length < 40 || /\n/.test(raw)) return null;
    const parts = raw.split(/(?<=。)/).map((x) => x.trim()).filter(Boolean);
    if (parts.length < 3) return null;
    const labelled = parts.filter((x) => /^[^：:、。]{2,16}\s*[：:]/.test(x));
    if (labelled.length !== parts.length) return null;
    return parts.map((x) => x.replace(/。$/, ""));
  }
  /** パネル・セルの本文が上の形なら箇条書きに割る(情報は落とさない) */
  function splitBlobs(node) {
    if (!node || typeof node !== "object") return false;
    let hit = false;
    const one = (arr) => {
      if (!Array.isArray(arr) || arr.length !== 1 || typeof arr[0] !== "string") return null;
      return splitLabelledSentences(arr[0]);
    };
    const split = one(node.items);
    if (split) {
      node.items = split;
      hit = true;
    } else if (!node.items && node.text) {
      const s2 = splitLabelledSentences(node.text);
      if (s2) {
        node.items = s2;
        delete node.text;
        hit = true;
      }
    }
    return hit;
  }

  function refineComposition(out) {
    try {
      // Preserve the explicitly requested alternate layout and LLM text hierarchy.
      if (out.layoutIntent === "outline") return out;
      out = seriesToLine(out);
      // 箇条書きで返るか 1 文の塊で返るかは LLM 次第なので、ラベル付きの塊は割って揃える
      const targets = out && out.body ? (out.body.cols || [out.body]) : [];
      if (targets.filter((n) => !JSON.stringify(n).includes("**")).some(splitBlobs)) out.compositionRepairs = (out.compositionRepairs || []).concat(["labelled blob → list"]);
      const b0 = out && out.body;
      if (b0 && b0.type === "gantt" && Array.isArray(b0.periods) && Array.isArray(b0.tasks) && b0.tasks.length) {
        // 使われていない末尾の期間は落とす(LLM が 2 年分の月を並べても、最後の工程・節目までで切る)
        const used = Math.max.apply(null, b0.tasks.map((t) => Number(t.end) || 0).concat((b0.milestones || []).map((m) => Number(m.at) || 0)));
        if (used >= 3 && b0.periods.length > used + 1) {
          const keep = used + 1;
          b0.periods = b0.periods.slice(0, keep);
          if (Array.isArray(b0.groups) && b0.groups.length) {
            let remaining = keep;
            b0.groups = b0.groups.map((g) => {
              const span = Math.max(0, Math.min(remaining, Number(g.span) || 0));
              remaining -= span;
              return { text: g.text, span };
            }).filter((g) => g.span > 0);
          }
          out.compositionRepairs = (out.compositionRepairs || []).concat(["unused periods trimmed"]);
        }
      }
      if (b0 && b0.type === "gantt" && Array.isArray(b0.periods) && b0.periods.length >= 4 && b0.periods.every((x) => /^\d{1,2}月$/.test(String(x)))) {
        const qre = /^(?:(20\d\d)\s*(?:年度?|FY)?\s*)?(?:Q\s?\d|第\s?\d\s?四半期|\dQ)$/i;
        // 上位ラベルが「2026年度」のような年・年度だけの並びなら尊重し、それ以外(Q4/Q1 の取り違え、「四半期」「月」などの無意味なラベル)は年度四半期で組み直す
        const yearOnly = (g) => /^(20\d\d)\s*(年度?|FY)?$/.test(String(g.text || "").trim());
        const looksQuarter = !b0.groups || !b0.groups.length || !b0.groups.every(yearOnly);
        if (looksQuarter) {
          // 年度四半期(4〜6 月 = Q1)で組み直す。LLM のラベルに年があれば「2026年度 Q1」のように年度を付け、3 月→4 月で年度を進める
          const first = (b0.groups || []).map((g) => (String(g.text || "").match(/(20\d\d)/) || [])[1]).find(Boolean);
          const q = (m) => Math.floor(((m + 8) % 12) / 3) + 1;
          const m0 = Number(String(b0.periods[0]).replace("月", ""));
          let fy = first ? Number(first) - (m0 >= 4 ? 0 : 1) : null;
          const groups = [];
          let prev = null;
          b0.periods.forEach((lab) => {
            const m = Number(String(lab).replace("月", ""));
            if (prev != null && m === 4 && fy != null) fy++;
            prev = m;
            const t = (fy != null ? fy + "年度 " : "") + "Q" + q(m);
            if (groups.length && groups[groups.length - 1].text === t) groups[groups.length - 1].span++;
            else groups.push({ text: t, span: 1 });
          });
          b0.groups = groups;
          out.compositionRepairs = (out.compositionRepairs || []).concat(["quarters recomputed (FY)"]);
        }
      }
      // 箱フローの見出しに LLM が付けた「01 」などの番号は外す(エンジンが番号を振るので二重になる)
      if (b0 && b0.type === "sequence" && Array.isArray(b0.steps)) b0.steps.forEach((st) => (st.head = String(st.head || "").replace(/^\s*(?:\d{1,2}|[A-Z])[\.\)\s:：]+\s*/, "")));
      if (b0 && (b0.type === "table" || b0.type === "ntable") && Array.isArray(b0.rows)) {
        // 列の全セルが同じ「ラベル: 」で始まるなら、ラベルを列見出しに昇格させて本文から外す(4 回同じ前置きを読ませない)
        const ncol2 = Math.max((b0.colHeaders || []).length, ...b0.rows.map((rw) => (rw.cells || []).length));
        for (let j = 0; j < ncol2; j++) {
          const texts = b0.rows.map((rw) => cellText((rw.cells || [])[j]));
          if (texts.length < 2 || texts.some((t) => !t.trim())) continue;
          const pref = texts.map((t) => (t.match(/^\s*([^:：]{2,14})[:：]\s*/) || [])[1]).filter(Boolean);
          if (pref.length !== texts.length || new Set(pref).size !== 1) continue;
          const existing = (b0.colHeaders || [])[j];
          if (existing && existing.trim() && existing.trim() !== pref[0]) continue;
          b0.colHeaders = b0.colHeaders || [];
          while (b0.colHeaders.length < ncol2) b0.colHeaders.push("");
          const changed = b0.colHeaders[j] !== pref[0];
          b0.colHeaders[j] = pref[0];
          b0.rows.forEach((rw) => {
            const cl = (rw.cells || [])[j];
            const t = cellText(cl).replace(/^\s*[^:：]{2,14}[:：]\s*/, "");
            if (cl && typeof cl === "object") cl.text = t;
            else rw.cells[j] = t;
          });
          if (changed) out.compositionRepairs = (out.compositionRepairs || []).concat(["cell prefix → column header"]);
        }
        // 全行が空(— / 空文字)の列は落とす(空の観点列は情報ではない)
        const ncol = Math.max((b0.colHeaders || []).length, ...b0.rows.map((rw) => (rw.cells || []).length));
        const empty = [];
        for (let j = 0; j < ncol; j++) if (b0.rows.every((rw) => isNA(cellText((rw.cells || [])[j])) || !cellText((rw.cells || [])[j]).trim())) empty.push(j);
        if (empty.length && empty.length < ncol) {
          b0.rows.forEach((rw) => (rw.cells = (rw.cells || []).filter((_, j) => !empty.includes(j))));
          if (b0.colHeaders) b0.colHeaders = b0.colHeaders.filter((_, j) => !empty.includes(j));
          out.compositionRepairs = (out.compositionRepairs || []).concat(["empty column dropped"]);
        }
        // Keep the native Table threshold consistent with normalization and the prompt.
        if (b0.type === "ntable" && b0.rows.length < 8 && b0.rows.length * (b0.colHeaders || []).length < 36) {
          b0.type = "table";
          out.compositionRepairs = (out.compositionRepairs || []).concat(["ntable → matrix"]);
        }
      }
      const b = out && out.body;
      if (!out || out.panelCount !== 1 || !b || b.type !== "cell" || !Array.isArray(b.items) || b.items.length < 4) return out;
      const flat = b.items.filter((x) => typeof x === "string");
      if (flat.length !== b.items.length) return out; // 子階層がある列挙はそのまま
      if (flat.some((t) => t.includes("**"))) return out; // Editorial labels already express the hierarchy.
      const parsed = flat.map((t) => {
        const m = String(t).match(/^\s*([^:：]{2,16})[:：]\s*(.+)$/);
        return m ? { head: m[1].trim(), text: m[2].trim() } : null;
      });
      if (parsed.filter(Boolean).length < Math.ceil(flat.length * 0.75)) return out;
      const rows = parsed.map((p, i) => (p ? { head: p.head, cells: [p.text] } : { head: "", cells: [flat[i]] }));
      out.body = { type: "table", corner: "", colHeaders: [], rows, numbered: false };
      out.compositionRepairs = (out.compositionRepairs || []).concat(["labelled list → matrix"]);
    } catch (_) {}
    return out;
  }

  /** 「2023 年 1,200 件、2024 年 1,500 件 …」のような系列(3 点以上)を含む文章を折れ線 + note に組み替える */
  function seriesToLine(out) {
    if (!out || !out.body) return out;
    const textOfCell = (c) => [c.text].concat((c.items || []).flatMap((x) => (Array.isArray(x) ? x : [x]))).filter(Boolean).map(String).join("\n");
    const find = (t) => {
      const re = /(20\d\d)\s*年[のに]?\s*([\d,]+(?:\.\d+)?)\s*(件|円|人|%|万円|億円|社|台|回)?/g;
      const pts = [];
      let m;
      while ((m = re.exec(t))) pts.push({ year: m[1], value: Number(m[2].replace(/,/g, "")), unit: m[3] || "" });
      const uniq = [];
      pts.forEach((p) => {
        if (!uniq.some((q) => q.year === p.year)) uniq.push(p);
      });
      return uniq.length >= 3 && uniq.every((p) => !isNaN(p.value)) ? uniq : null;
    };
    // 図に載せた系列を含む文だけを落とし、残りの文は一字も変えずに残す。
    // (数値だけを正規表現で抜くと「月間問合せ件数はで増加している」のような壊れた文が残る)
    const strip = (t) => {
      const yearVal = /(20\d\d)\s*年[のに]?\s*[\d,]+(?:\.\d+)?/g;
      return String(t || "")
        .split(/(?<=[。\n])/)
        .filter((sent) => ((sent.match(yearVal) || []).length < 2))
        .join("")
        .replace(/^[、。\s]+|[、。\s]+$/g, "");
    };
    if (out.panelCount === 2 || out.panelCount === 3) {
      const cols = out.body.cols || [];
      const hits = cols.map((c) => find(textOfCell(c)));
      const idx = hits.findIndex(Boolean);
      if (idx < 0 || hits.filter(Boolean).length !== 1) return out;
      const pts = hits[idx];
      const panel = cols[idx];
      const notes = [];
      cols.forEach((c, i) => {
        const t = i === idx ? strip(textOfCell(c)) : textOfCell(c);
        const lines = t.split("\n").map((x) => x.trim()).filter((x) => x.length >= 6);
        if (lines.length) notes.push((c.head ? c.head + ": " : "") + lines.join("。").replace(/。。/g, "。"));
      });
      out.panelCount = 1;
      out.body = { type: "line", unit: pts[0].unit, labels: pts.map((p) => p.year), series: [{ label: stripBold(panel.head || ""), values: pts.map((p) => p.value) }] };
      if (notes.length) out.note = { type: "cell", text: notes.join("\n") };
      out.compositionRepairs = (out.compositionRepairs || []).concat(["series text → line"]);
      return out;
    }
    if (out.panelCount === 1 && out.body.type === "cell") {
      const t = textOfCell(out.body);
      const pts = find(t);
      if (!pts) return out;
      const rest = strip(t).split("\n").map((x) => x.trim()).filter((x) => x.length >= 6);
      out.body = { type: "line", unit: pts[0].unit, labels: pts.map((p) => p.year), series: [{ label: stripBold(out.body.head || ""), values: pts.map((p) => p.value) }] };
      const prev = out.note && out.note.text ? [out.note.text] : [];
      if (rest.length || prev.length) out.note = { type: "cell", text: prev.concat(rest).join("\n") };
      out.compositionRepairs = (out.compositionRepairs || []).concat(["series text → line"]);
    }
    return out;
  }

  function renderGanttComposition(node, r, c) {
    const P = c.P;
    const labelW = r.w * 0.19;
    const noteW = node.tasks.some((t) => t.note) ? r.w * 0.24 : 0;
    const plotX = r.x + labelW, plotW = r.w - labelW - noteW;
    const pw = plotW / node.periods.length;
    // マイルストーンは 1 行にまとめる(工程 10 行 + 節目 4 行で版面を潰さない)
    const ms = (node.milestones || []).slice().sort((a, b) => a.at - b.at);
    // 隣り合う期間に節目が並ぶとラベルが重なるので、そのときだけ 2 段(偶数番目 / 奇数番目)に分ける
    const crowded = ms.some((m, i) => i > 0 && m.at - ms[i - 1].at <= 1);
    const msRows = !ms.length ? [] : crowded ? [{ label: "マイルストーン", milestones: ms.filter((_, i) => i % 2 === 0) }, { label: "", milestones: ms.filter((_, i) => i % 2 === 1) }] : [{ label: "マイルストーン", milestones: ms }];
    const rows = node.tasks.map((t) => ({ label: t.label, note: t.note, task: t })).concat(msRows);
    // 行数が多いときは文字を段階的に縮めて 1 枚に収める(行を文字より低くして重ねない)
    let fs = FONT.body;
    const needsAt = (f) => rows.map((t) => Math.max(f * FONT.lineHeight + 6, textHeight(t.label, labelW - 12, f) + 6, t.note ? textHeight(t.note, noteW - 12, f - 2) + 6 : 0, t.milestones ? f * FONT.lineHeight + 6 : 0));
    const headAt = (f) => f * FONT.lineHeight + 12;
    let needs = needsAt(fs);
    let hh = headAt(fs);
    let groupH = node.groups && node.groups.length ? hh : 0;
    let available = r.h - hh - groupH;
    let need = needs.reduce((a, b) => a + b, 0);
    while (need > available && fs > FONT.tableMin) {
      fs -= 1;
      needs = needsAt(fs);
      hh = headAt(fs);
      groupH = node.groups && node.groups.length ? hh : 0;
      available = r.h - hh - groupH;
      need = needs.reduce((a, b) => a + b, 0);
    }
    if (need > available * 1.15) c.warnings.push("gantt overflow: row text needs " + Math.ceil(need) + "pt"); // 15% までの詰めは行の比例縮小で吸収できる
    // 余りは行に配分するが、行を自然高さの 1.6 倍より間延びさせない
    const heights = needs.map((h) => (need > available ? (h * available) / need : Math.min(h * 1.6, h + (available - need) / rows.length)));
    let y = r.y;
    if (groupH) {
      let col = 0;
      for (const g of node.groups) {
        c.prims.push(textBox(plotX + col * pw, y, g.span * pw, groupH, g.text, { fontSize: fs, fill: P.fillDark, color: P.textOnDark, bold: true, align: "center", valign: "middle", pad: 3, role: "matrixcell" }));
        col += g.span;
      }
      y += groupH;
    }
    // 時系列の見出し(年度四半期・月など)は既定で濃い帯。時間軸は表の一部ではなく骨格そのものなので、常に立てる
    node.periods.forEach((p, i) => c.prims.push(textBox(plotX + pw * i, y, pw, hh, p, { fontSize: fs, fill: P.fillDark, color: P.textOnDark, bold: true, align: "center", valign: "middle", pad: 3, role: "matrixcell" })));
    if (noteW) c.prims.push(textBox(plotX + plotW, y, noteW, hh, "補足", { fontSize: fs, bold: true, align: "left", valign: "middle", role: "matrixcell" }));
    y += hh;
    const bodyTop = y;
    c.prims.push(line(r.x, y, r.x + r.w, y, P.line, RULE_THIN));
    for (let j = 1; j < node.periods.length; j++) c.prims.push(line(plotX + pw * j, y, plotX + pw * j, r.y + r.h, P.lineLight, RULE_THIN));
    rows.forEach((row, i) => {
      const h = heights[i];
      c.prims.push(textBox(r.x, y, labelW, h, row.label, { fontSize: fs, color: P.text, valign: "middle", pad: 6, role: "matrixcell" }));
      if (row.task) {
        const t = row.task, bh = Math.min(24, h * 0.45);
        c.prims.push(rect(plotX + t.start * pw + 2, y + (h - bh) / 2, (t.end - t.start + 1) * pw - 4, bh, { shape: "homePlate", fill: P.fillMid }));
        if (t.note) c.prims.push(textBox(plotX + plotW, y, noteW, h, t.note, { fontSize: fs - 2, color: P.textMuted, valign: "middle", pad: 6, role: "note" }));
      } else {
        // 節目は 1 行に並べる(◆ + 短いラベル。右端で切れるなら左側に出す)
        (row.milestones || []).forEach((m) => {
          const cx = plotX + (m.at + 0.5) * pw;
          const tw = Math.min(pw * 2.2, measure(m.label, fs - 2) + 16);
          const right = cx + 10 + tw <= plotX + plotW;
          c.prims.push(rect(cx - 4, y + h / 2 - 4, 8, 8, { fill: P.fillDark, rotation: 45 }));
          const tx = right ? cx + 8 : Math.max(plotX, cx - tw - 8);
          c.prims.push(textBox(tx, y, tw, h, m.label, { fontSize: Math.max(FONT.tableMin, fs - 2), color: P.text, align: right ? "left" : "right", valign: "middle", pad: 2, autofit: "none", role: "note" }));
        });
      }
      y += h;
      if (i < rows.length - 1) c.prims.push(line(r.x, y, r.x + r.w, y, P.lineLight, RULE_THIN));
    });
    if (node.today != null && node.today >= 0 && node.today < node.periods.length) {
      const x = plotX + (node.today + .5) * pw;
      c.prims.push(line(x, bodyTop, x, r.y + r.h, P.fillDark, 1.2));
    }
  }

  function renderComposition(spec, r, c) {
    if (spec.panelCount > 1) {
      const kids = spec.body.cols;
      const figIdx = kids.findIndex((k) => k._figurePanel);
      // 図 + 補足説明の 2 パネルは 1.8 : 1(ゴールデン 8 枚の実測の中央値)。それ以外は等幅
      const units = kids.map((_, i) => (figIdx >= 0 && kids.length === 2 && i === figIdx ? 1.8 : 1));
      const unitSum = units.reduce((a, u) => a + u, 0);
      const avail = r.w - GAP * (kids.length - 1);
      const ws = units.map((u) => (avail * u) / unitSum);
      const xs = [];
      kids.reduce((cx, _, i) => (xs.push(cx), cx + ws[i] + GAP), r.x);
      const hh = Math.max(...kids.map((k, i) => textHeight(k.head, ws[i] - PAD * 2, FONT.head) + PAD * 2 + 2));
      // 兄弟パネルで行の高さを共有し、段数が違っても行の罫が同じ y に並ぶようにする
      const textKids = kids.filter((k) => !k._figurePanel);
      const rowsOf = (k) => (k.items || []).reduce((acc, v) => acc + (Array.isArray(v) ? v.length : 1), 0);
      const siblingRows = textKids.length >= 2 ? Math.max(...textKids.map(rowsOf)) : 0;
      // 「**ラベル**：」で始まる列挙は点を打たない。片方だけ点が付くと左右で体裁が変わるので、全パネルで揃える
      const labelled = textKids.every((k) => (k.items || []).every((v) => !Array.isArray(v) && /^\*\*[^*]+\*\*\s*[:：]/.test(itemToText(v))));
      kids.forEach((k, i) => {
        const w = ws[i], x = xs[i];
        c.prims.push(textBox(x, r.y, w, hh - 2, k.head, { fontSize: FONT.head, bold: true, color: c.P.text, align: "center", valign: "middle", role: "panelhead" }));
        c.prims.push(line(x, r.y + hh, x + w, r.y + hh, c.P.text, RULE_THICK));
        const br = { x, y: r.y + hh + GAP / 2, w, h: r.h - hh - GAP / 2 };
        if (k._figurePanel) (LEAF[k.type] || LEAF.cell)(Object.assign({}, k, { _composition: true, highlight: false }), br, c);
        else LEAF.cell({ type: "cell", items: k.items, valign: "middle", _inPanel: true, _gid: 1, _siblingRows: siblingRows, _labelledOutline: labelled }, br, c);
      });
      return;
    }
    let mr = Object.assign({}, r);
    if (spec.note) {
      const noteFs = Math.max(14, Math.min(20, FONT.body - 2));
      const nh = textHeight(spec.note.text, r.w - PAD * 2, noteFs) + PAD * 2;
      if (nh > r.h * 0.28) c.warnings.push("note too large: move detail to another slide");
      mr.h -= nh + GAP;
      if (mr.h < 100) c.warnings.push("body area too small after note");
      c.prims.push(textBox(r.x, r.y + mr.h + GAP, r.w, nh, spec.note.text, { fontSize: noteFs, color: c.P.text, valign: "middle", role: "note" }));
    }
    const b = spec.body;
    if (b.type === "sequence") {
      const steps = b.steps;
      const w = (mr.w - GAP * (steps.length - 1)) / steps.length;
      const vertical = steps.length > 5 || steps.some((s) => estimateLines(s.text, w - PAD * 2, FONT.body) > 3 || s.text.length > 48);
      if (vertical) {
        LEAF.table({ type: "table", colHeaders: [], headShape: b.headShape, numbered: true, _composition: true, rows: steps.map((s) => ({ head: s.head, cells: [s.text] })) }, mr, c);
      } else {
        const hh = Math.max(52, ...steps.map((s, i) => textHeight(String(i + 1).padStart(2, "0") + "  " + s.head, w - PAD * 2, FONT.head) + PAD * 2));
        const bh = Math.max(...steps.map((s) => textHeight(s.text, w - PAD * 2, FONT.body))) + PAD * 3;
        const iconSize = steps.every((s) => s.icon) && steps.every((s) => estimateLines(s.text, w - PAD * 2, FONT.body) <= 2) ? Math.max(0, Math.min(96, w * 0.42, mr.h - hh - bh - GAP * 2)) : 0;
        const iconArea = iconSize >= 48 ? iconSize + GAP : 0;
        const totalH = hh + GAP / 2 + bh + iconArea;
        const sy = mr.y + Math.max(0, (mr.h - totalH) * 0.35);
        if (totalH > mr.h) c.warnings.push("sequence overflow: text exceeds body");
        steps.forEach((s, i) => {
          const x = mr.x + i * (w + GAP);
          c.prims.push(textBox(x, sy, w, hh, String(i + 1).padStart(2, "0") + "  " + s.head, { shape: "homePlate", fill: c.P.fillLight, fontSize: FONT.head, bold: true, color: c.P.text, align: "center", valign: "middle", role: "sequencehead" }));
          if (iconArea) c.prims.push(image(s.icon, x + (w - iconSize) / 2, sy + hh + GAP, iconSize, iconSize, c.P.fillDark));
          LEAF.cell({ type: "cell", text: s.text, _inPanel: true, _gid: 1, valign: "middle" }, { x, y: sy + hh + GAP / 2 + iconArea, w, h: bh }, c);
        });
      }
    } else {
      const node = Object.assign({}, b, { _composition: true });
      node.highlight = false;
      if (b.type === "cell") { node.valign = "middle"; if (!b.head) node._inPanel = true; }
      (LEAF[b.type] || LEAF.cell)(node, mr, c);
    }
  }

  function normalizeSpec(raw) {
    if (raw && (raw.panelCount != null || raw.panels || raw.compositionVersion)) return refineComposition(normalizeGeneratedSpec(raw));
    const s = raw && typeof raw === "object" ? raw : {};
    const out = {
      title: cleanTitle(stripBold(str(s.title || s.headline))).slice(0, 120),
      lead: str(s.lead || s.message || s.subtitle).slice(0, 240),
      footnote: str(s.footnote || s.source).slice(0, 160),
      kicker: str(s.kicker || s.eyebrow || s.section).slice(0, 60),
      body: null,
    };
    const b = s.body != null ? s.body : s.blocks || s.content || {};
    out.body = normNode(b, 0, { leaves: 0 }) || { type: "cell", text: "" };
    applyMethodology(out.body, 888);
    // 1 スライド 1 デザインの強制は、生成応答では normalizeGeneratedSpec(composition v1)が担う。
    // 旧形式(モック・保存済み spec)にも同じ規律を当てたいときは enforceOneDesign(out.body) を通す
    enforceBudget(out.body, (out.title || "") + (out.lead || ""));
    return out;
  }
  /** タイトル規約: 「〜です/ます」を除く、「ラベル：」前置きを外す */
  function cleanTitle(t) {
    t = t.replace(/。$/u, "");
    const conv = [
      [/ています$/u, "ている"], [/できます$/u, "できる"], [/します$/u, "する"], [/なります$/u, "なる"], [/あります$/u, "ある"],
      [/います$/u, "いる"], [/れます$/u, "れる"], [/えます$/u, "える"], [/ります$/u, "る"], [/ます$/u, "る"], [/です$/u, ""], [/である$/u, ""],
    ];
    for (const [re, to] of conv) if (re.test(t)) { t = t.replace(re, to); break; }
    const m = t.match(/^[^：:]{1,12}[：:]\s*(.{8,})$/u);
    if (m) t = m[1];
    return t.trim();
  }

  /** 強調バジェット: highlight は 1 枚 2 箇所まで、tint 塗りは 1 つまで(先勝ち)。kpi と単系列の折れ線は数えない */
  /**
   * 1 スライド 1 デザイン: 図・格子・ガント・体制図・箱フロー・番号カードのうち 1 つだけを主デザインとして残し、
   * それ以外の内容は主デザインの下の全幅テキスト 1 つに畳む(情報は落とさない)。
   * 下線付きテキストパネルだけの構成(表現空間の分割)はデザインではないので、そのまま通す。
   */
  function enforceOneDesign(body) {
    if (!body || typeof body !== "object") return body;
    const isDesignNode = (n) => {
      if (!n || typeof n !== "object") return false;
      if (isContainer(n) && n.cols) {
        const ks = n.cols.filter((k) => k && k.type !== "arrow");
        return ks.length >= 2 && ks.every((k) => k.type === "cell" && (k.shape === "step" || k.shape === "chevron" || k.shape === "home" || (k.num && k.head)));
      }
      if (isContainer(n)) return false;
      return !["cell", "arrow"].includes(n.type);
    };
    const designs = [];
    const walk = (n) => {
      if (!n || typeof n !== "object") return;
      if (isDesignNode(n)) return designs.push(n);
      if (isContainer(n)) (n.rows || n.cols).forEach(walk);
    };
    walk(body);
    if (!designs.length) return body;
    if (designs.length === 1 && (isDesignNode(body) || (isContainer(body) && body.rows && body.rows.length === 2 && isDesignNode(body.rows[0]) && body.rows[1].type === "cell" && !body.rows[1].head))) return body;
    // 主デザイン: 面積の大きい種類を優先(格子・ガント・体制図 > 図・箱フロー・番号カード > kpi)。同順位なら先に出る方
    const rank = (n) => (isContainer(n) ? 2 : ["table", "ntable", "gantt", "org", "tree"].includes(n.type) ? 3 : n.type === "kpi" ? 0 : 2);
    let primary = designs[0];
    designs.forEach((d) => {
      if (rank(d) > rank(primary)) primary = d;
    });
    // 主デザインの直上にある見出しだけの cell は、その見出しとして残す
    let heading = null;
    const findHeading = (n) => {
      if (!n || !isContainer(n)) return;
      const ks = n.rows || n.cols;
      if (n.rows) {
        const i = ks.indexOf(primary);
        if (i > 0 && isHeadingCell(ks[i - 1])) heading = ks[i - 1];
      }
      ks.forEach(findHeading);
    };
    findHeading(body);
    const items = [];
    const push = (t) => {
      t = String(t || "").replace(/\s+/g, " ").trim();
      if (t) items.push(t);
    };
    const flatItems = (arr) => (arr || []).flatMap((it) => (Array.isArray(it) ? it : [it])).map((x) => String(x)).filter(Boolean);
    const bodyText = (c) => (c.text && String(c.text).trim()) || flatItems(c.items).join("、");
    const headOf = (c) => stripBold(String(c.head || "")).trim();
    const textOf = (n) => {
      if (!n || typeof n !== "object" || n === primary || n === heading) return;
      switch (n.type) {
        case "cell": {
          const h = headOf(n), b = bodyText(n);
          if (h && b) push("**" + h + "**: " + b);
          else if (b) push(b);
          else if (h) push(h);
          return;
        }
        case "arrow":
          return;
        case "table":
        case "ntable":
          (n.rows || []).forEach((rw) => {
            const cells = (rw.cells || []).map((cl, j) => ((n.colHeaders || [])[j] ? stripBold(n.colHeaders[j]) + " " : "") + cellText(cl).replace(/\n/g, "、")).filter((x) => x.trim());
            push((rw.head ? "**" + stripBold(rw.head) + "**: " : "") + cells.join(" / "));
          });
          return;
        case "bars":
        case "column":
          push((n.items || []).map((it) => stripBold(it.label || "") + " " + it.value + (n.unit || "")).join("、"));
          return;
        case "line":
          (n.series || []).forEach((sv) => push(stripBold(sv.label || "") + ": " + (sv.values || []).join(" → ") + (n.unit || "")));
          return;
        case "stacked":
          (n.items || []).forEach((it) => push(stripBold(it.label || "") + ": " + (it.values || []).map((v, i) => stripBold((n.parts || [])[i] || "") + " " + v).join("、")));
          return;
        case "kpi":
          push(stripBold(n.label || "") + " " + (n.value || ""));
          return;
        case "org": {
          const one = (nd) => stripBold(nd.label || "") + (nd.sub ? "(" + stripBold(nd.sub) + ")" : "");
          const rec = (nd) => {
            if (!nd) return;
            if (nd.children && nd.children.length) push("**" + one(nd) + "**: " + nd.children.map(one).join("、"));
            (nd.children || []).forEach(rec);
          };
          rec(n.root);
          return;
        }
        case "gantt": {
          const per = n.periods || [];
          push((n.tasks || []).map((t) => stripBold(t.label || "") + " " + (per[t.start] || "") + (t.end !== t.start ? "〜" + (per[t.end] || "") : "")).join("、"));
          (n.milestones || []).forEach((m) => push(stripBold(m.label || "") + " " + (per[m.at] || "")));
          return;
        }
        default:
          return;
      }
    };
    const collect = (n) => {
      if (!n || typeof n !== "object" || n === primary || n === heading) return;
      if (isContainer(n)) {
        const ks = n.rows || n.cols;
        for (let i = 0; i < ks.length; i++) {
          const k = ks[i];
          // 見出しだけの cell + 本文 cell(パネルの形)は 1 項目にまとめる
          if (n.rows && isHeadingCell(k) && ks[i + 1] && ks[i + 1].type === "cell" && !ks[i + 1].head && ks[i + 1] !== primary) {
            const b = bodyText(ks[i + 1]);
            push("**" + headOf(k) + "**" + (b ? ": " + b : ""));
            i++;
            continue;
          }
          collect(k);
        }
        return;
      }
      textOf(n);
    };
    collect(body);
    const rows = [];
    if (heading) rows.push(heading);
    rows.push(primary);
    if (items.length) rows.push({ type: "cell", items: items.slice(0, 10) });
    return rows.length === 1 ? primary : { rows };
  }

  function enforceBudget(body, mention) {
    // 強調は「結論(title / lead)に登場する対象」にだけ許す。LLM が付けた根拠のない highlight / dark は落とす
    const mentionText = String(mention || "").replace(/\s+/g, "");
    const mentioned = (t) => {
      const k = stripBold(String(t || "")).replace(/\s+/g, "").replace(/[（(].*?[)）]/g, "");
      return k.length >= 2 && mentionText.includes(k);
    };
    let hl = 0,
      tints = 0,
      darks = 0;
    const textLenOf = (n) => ((n.text || "") + (n.items || []).flat().join("")).replace(/\*\*/g, "").length;
    const takeHl = () => (hl < BUDGET.highlight ? (hl++, true) : false);
    const plainHeaded = (k) => !!k && k.type === "cell" && !!k.head && String(k.head).trim() && !(k.fill && k.fill !== "none") && !k.shape && !k.style;
    // 罫の見出しを持つパネル: 見出し付き cell(塗りなし)か、先頭がそれである rows
    const rulePanel = (k) => !!k && (plainHeaded(k) || (isContainer(k) && !!k.rows && (isHeadingCell(k.rows[0]) || plainHeaded(k.rows[0]))));
    const stripLight = (k) => {
      if (!k) return;
      if (k.type === "cell" && (k.fill === "light" || k.fill === "tint") && k.head && String(k.head).trim()) k.fill = "none";
      else if (isContainer(k) && k.rows && k.rows[0]) stripLight(k.rows[0]);
    };
    const walk = (n) => {
      if (!n || typeof n !== "object") return;
      if (isContainer(n)) {
        const kids = n.rows || n.cols;
        // 横並びのパネルで、罫の見出し(塗りなし)と薄い塗りの箱が混在すると片方だけ罫が抜けて見える → 罫に統一する(濃い箱は強調として残す)
        if (n.cols && kids.some(rulePanel)) kids.forEach(stripLight);
        return kids.forEach(walk);
      }
      switch (n.type) {
        case "cell":
          if (n.highlight && !(mentioned(n.head) || mentioned(n.text))) n.highlight = false;
          if (n.fill === "dark" && !mentioned(n.head)) n.fill = "none";
          if (n.highlight && !takeHl()) n.highlight = false;
          if (n.fill === "tint" && tints++ >= BUDGET.tint) n.fill = "none";
          // 濃い塗りは「一番立てたい 1 つ」の短い強調だけ。長文や 2 つ目は薄い塗りに落とす(黒い塊で版面を潰さない)
          if (n.fill === "dark" && (!(n.head && String(n.head).trim()) || darks++ >= 1)) n.fill = "light"; // 反転できる見出しが無い / 2 つ目 → 薄い塗り
          break;
        case "table":
        case "ntable":
          if (typeof n.highlightCol === "number" && !mentioned((n.colHeaders || [])[n.highlightCol])) n.highlightCol = undefined;
          if (typeof n.highlightCol === "number" && !takeHl()) n.highlightCol = undefined;
          (n.rows || []).forEach((rw) => {
            if (rw.highlight && !mentioned(rw.head)) rw.highlight = false;
            if (rw.highlight && !takeHl()) rw.highlight = false;
            (rw.cells || []).forEach((cl, j) => {
              if (cl && typeof cl === "object" && cl.highlight && !(mentioned(cl.text) || (mentioned(rw.head) && mentioned((n.colHeaders || [])[j])))) cl.highlight = false;
              if (cl && typeof cl === "object" && cl.highlight && !takeHl()) cl.highlight = false;
            });
          });
          break;
        case "bars":
        case "column":
          (n.items || []).forEach((it) => {
            if (it.highlight && !mentioned(it.label)) it.highlight = false;
            if (it.highlight && !takeHl()) it.highlight = false;
          });
          break;
        case "gantt":
          (n.tasks || []).forEach((t) => {
            if (t.highlight && !mentioned(t.label)) t.highlight = false;
          });
          break;
      }
    };
    walk(body);
    // dark → light の格下げ後に、罫の見出しと薄い箱の混在をもう一度ならす
    const unify = (n) => {
      if (!n || typeof n !== "object" || !isContainer(n)) return;
      const kids = n.rows || n.cols;
      if (n.cols && kids.some(rulePanel)) kids.forEach(stripLight);
      kids.forEach(unify);
    };
    unify(body);
  }

  function normNode(n, depth, state) {
    if (n == null) return null;
    if (typeof n === "string") {
      // {"type":"arrow"} が ["type","arrow"] のように崩れて届くことがある。スキーマの語だけの断片は捨てる
      if (/^(type|arrow|cell|table|ntable|rows|cols|items|head|text|figure|kpi|bars|line|column|stacked|highlight|fill|shape|weights|corner|colheaders)$/i.test(n.trim())) return null;
      return { type: "cell", text: n };
    }
    if (Array.isArray(n)) n = { rows: n };
    if (typeof n !== "object" || depth > BUDGET.depth) return null;

    // {"table":{…}} / {"cell":{…}} のように型名をキーにした包み方を吸収する(LLM のゆらぎ)
    if (!n.type && !n.rows && !n.cols) {
      const wrapKey = Object.keys(n).find((k) => WRAP_TYPES.includes(k.toLowerCase()) && n[k] && typeof n[k] === "object");
      if (wrapKey && Object.keys(n).length <= 2) {
        const inner = n[wrapKey];
        return normNode(Array.isArray(inner) ? { type: wrapKey, items: inner } : Object.assign({ type: wrapKey }, inner), depth, state);
      }
    }
    // {"type":"cols","cols":[…]} / {"type":"rows",…} のように容器に type を付けて返す LLM のゆらぎを吸収する
    const ctype = String(n.type || "").toLowerCase();
    if ((ctype === "cols" || ctype === "rows" || ctype === "grid" || ctype === "container" || ctype === "group" || ctype === "row" || ctype === "col") && (Array.isArray(n.cols) || Array.isArray(n.rows) || Array.isArray(n.children) || Array.isArray(n.items))) {
      const c = Object.assign({}, n);
      delete c.type;
      if (!c.rows && !c.cols) {
        if (ctype === "rows" || ctype === "row" || ctype === "container" || ctype === "group") c.rows = c.children || c.items;
        else c.cols = c.children || c.items;
      }
      return normNode(c, depth, state);
    }
    const rows = n.rows || n.children || n.blocks;
    const cols = n.cols || n.columns;
    if (Array.isArray(rows) && !n.type) {
      const kids = rows.map((k) => normNode(k, depth + 1, state)).filter(Boolean);
      return kids.length ? withWeights({ rows: kids }, n.weights, kids.length) : null;
    }
    if (Array.isArray(cols) && !n.type && !isLegacyColumns(n)) {
      const kids = cols.map((k) => normNode(k, depth + 1, state)).filter(Boolean);
      if (!kids.length) return null;
      return normCols(kids, n.weights);
    }

    const type = String(n.type || "").toLowerCase();
    // ---- 糖衣構文(旧型・別名)→ cell / grid ----
    if (type === "boxes") return normCols((n.items || n.boxes || []).map(normCard).slice(0, 8), null);
    if (type === "columns" || (isLegacyColumns(n) && !type)) {
      const cs = (n.columns || n.cols || n.items || []).map((col) => ({ type: "cell", head: str(col.head || col.title), items: (col.bullets || col.items || []).map(itemStr), highlight: !!col.highlight }));
      return normCols(cs.slice(0, 5), null);
    }
    if (type === "matrix") return normTable(n, "table");
    if (type === "chevrons" || type === "process" || type === "steps") return normChevrons(n);
    if (type === "figure") {
      const kinds = ["line", "bars", "column", "stacked", "kpi", "gantt", "org", "ntable", "table"];
      const k = String(n.kind || n.figure || "").toLowerCase();
      if (kinds.includes(k)) return normNode(Object.assign({}, n, { type: k }), depth, state);
      const key = Object.keys(n).find((x) => kinds.includes(x.toLowerCase()) && n[x] && typeof n[x] === "object");
      if (key) return normNode(Object.assign({ type: key.toLowerCase() }, n[key]), depth, state); // {"type":"figure","line":{…}}
      return normNode(Object.assign({}, n, { type: "bars" }), depth, state);
    }

    if (state.leaves >= BUDGET.leaves) return null;
    state.leaves++;
    switch (type) {
      case "cell":
      case "card":
      case "box":
        return dropEmptyCell(normCell(n));
      case "label":
      case "header":
        return { type: "cell", head: stripBold(str(n.text || n.head)), highlight: !!n.highlight };
      case "bullets":
      case "list":
        return { type: "cell", items: normItems(n.items || n.bullets), fill: n.style === "box" ? "light" : "none", highlight: !!n.highlight };
      case "callout":
      case "message":
        return { type: "cell", text: str(n.text), fill: "tint", align: n.align, highlight: !!n.highlight };
      case "icon":
      case "pictogram":
        return { type: "cell", icon: str(n.name || n.icon).toLowerCase(), text: str(n.label || n.text), align: "center", highlight: !!n.highlight };
      case "table":
        return degradeThinTable(normTable(n, "table"));
      case "ntable":
      case "nativetable":
        return normTable(n, "ntable");
      case "gantt":
      case "schedule":
      case "timeline":
        return normGantt(n);
      case "org":
      case "orgchart":
        return normOrg(n);
      case "tree":
      case "issuetree":
        // 論点ツリーは体制図と同じデータ構造だが、根を左に置いて右へ展開する(葉が増えても 16:9 に収まる)
        return Object.assign(normOrg(n), { type: "tree" });
      case "bars":
      case "bar":
      case "chart":
        return dropEmptyFigure({
          type: "bars",
          unit: str(n.unit),
          note: str(n.note),
          items: (n.items || n.data || []).slice(0, 10).map((it) => ({ label: str(it.label || it.name), value: Number(it.value) || 0, valueLabel: it.valueLabel != null ? str(it.valueLabel) : undefined, highlight: !!it.highlight })),
        });
      case "column":
      case "columns_chart":
      case "colchart":
        return dropEmptyFigure({ type: "column", unit: str(n.unit), note: str(n.note), items: (n.items || n.data || []).slice(0, 8).map((it) => ({ label: str(it.label || it.name), value: Number(it.value) || 0, highlight: !!it.highlight })) });
      case "line":
      case "linechart":
      case "trend":
        return dropEmptyFigure({
          type: "line",
          unit: str(n.unit),
          labels: (n.labels || n.x || []).map(str),
          series: (n.series || []).slice(0, 2).map((s) => ({ label: str(s.label || s.name), values: (s.values || s.data || []).map(Number), highlight: !!s.highlight })),
        });
      case "stacked":
      case "share":
      case "composition": {
        const its = (n.items || n.data || []).slice(0, 6);
        // 各項目の値が 1 つずつなら構成比ではなく量の比較。横棒に落とす(全項目 100% になるのを防ぐ)
        if (its.length && its.every((it) => ((it.values || it.parts || []).length || 0) <= 1)) {
          return dropEmptyFigure({ type: "bars", unit: str(n.unit), note: str(n.note), items: its.map((it) => ({ label: str(it.label || it.name), value: Number((it.values || it.parts || [])[0] != null ? (it.values || it.parts || [])[0] : it.value) || 0 })) });
        }
        return dropEmptyFigure({
          type: "stacked",
          parts: (n.parts || n.legend || []).map(str),
          items: its.map((it) => ({ label: str(it.label || it.name), values: (it.values || it.parts || []).map(Number) })),
        });
      }
      case "kpi":
      case "number":
        return dropEmptyFigure({ type: "kpi", value: str(n.value), label: str(n.label || n.text), highlight: n.highlight !== false, fill: n.fill === "dark" || n.fill === "accent" ? n.fill : undefined });
      case "arrow":
        return { type: "arrow", direction: n.direction === "right" || n.direction === "down" ? n.direction : undefined };
      case "text":
      case "paragraph":
      default: {
        const text = str(n.text || n.body || n.content || n.head || n.title);
        if (!text) return null;
        const fill = { fill: "light", box: "light", dark: "dark", tint: "tint" }[n.style] || "none";
        return { type: "cell", text, align: n.align, size: n.size === "large" ? "large" : undefined, fill, bold: !!n.bold, highlight: !!n.highlight };
      }
    }
  }
  /** cols の正規化: 2 パネル(+arrow)は左右等幅、葉が 5 個以上なら 2 行に折る(自動リフロー) */
  function normCols(kids, weights) {
    const content = kids.filter((k) => k.type !== "arrow");
    if (content.length <= 2) return { cols: kids, weights: kids.map((k) => (k.type === "arrow" ? 0.14 : 1)) };
    if (content.length >= 5 && content.length === kids.length && kids.every((k) => k.type === "cell" && !k.shape)) {
      const t = enumToTable(kids); // 方法論: 事象が多いときは横に割らず、行名の箱 + 横長の説明の格子で縦に並べる
      if (t) return t;
      const half = Math.ceil(kids.length / 2);
      return { rows: [{ cols: kids.slice(0, half) }, { cols: kids.slice(half) }] };
    }
    return withWeights({ cols: kids }, weights, kids.length);
  }
  /** 矢羽(順序): 横並びの帯 + 各段の要点。6 段以上は縦の格子(行見出し矢羽)に倒す */
  function normChevrons(n) {
    const steps = (n.steps || n.items || []).map(normCard).slice(0, 8);
    if (!steps.length) return null;
    const hasBody = steps.some((s) => (s.text && s.text.trim()) || (s.items && s.items.length));
    if (steps.length >= 6) {
      return { type: "table", corner: "", colHeaders: [], rows: steps.map((s) => ({ head: s.head, cells: [cellBody(s).text], highlight: !!s.highlight })), headShape: "chevron" };
    }
    const band = { cols: steps.map((s) => ({ type: "cell", head: s.head, shape: "home", highlight: !!s.highlight })), gap: 4 };
    if (!hasBody) return band;
    const bodies = { cols: steps.map((s) => ({ type: "cell", text: s.text, items: s.items })) };
    return { rows: [band, bodies], gap: GAP / 2 };
  }
  function isLegacyColumns(n) {
    const c = n.columns || n.cols;
    return Array.isArray(c) && c.length && c.every((x) => x && typeof x === "object" && (x.head || x.title) && (x.bullets || x.items) && !x.type && !x.rows && !x.cols);
  }
  function withWeights(node, weights, n) {
    if (Array.isArray(weights) && weights.length === n && weights.every((w) => typeof w === "number" && w > 0)) node.weights = weights;
    return node;
  }
  function normItems(items) {
    return (items || [])
      .slice(0, 10)
      .map((x) => (Array.isArray(x) ? x.slice(0, 6).map(itemStr) : itemStr(x)))
      .filter((x) => (Array.isArray(x) ? x.length : x));
  }
  /** 中身の無いセル(LLM の崩れで文字が全部落ちた等)は置かない */
  function dropEmptyCell(c) {
    if (!c || c.type !== "cell") return c;
    const has = (c.head && c.head.trim()) || (c.text && c.text.trim()) || (c.items && c.items.length) || c.icon || (c.caption && c.caption.trim());
    return has ? c : null;
  }
  function normCell(it) {
    const o = normCard(it);
    if ((it.style === "takeaway" || it.role === "takeaway") && o.items) {
      o.text = o.items.map((x) => (Array.isArray(x) ? x.join("、") : x)).join("。");
      delete o.items;
    }
    if (it.shape === "chevron" || it.shape === "home" || it.shape === "homeplate") o.shape = it.shape === "chevron" ? "chevron" : "home";
    if (it.shape === "step" || it.shape === "box") o.shape = "step";
    if (it.caption) o.caption = str(it.caption).slice(0, 40);
    if (it.num != null && String(it.num).trim()) o.num = String(it.num).trim().slice(0, 4);
    if (it.style === "takeaway" || it.role === "takeaway") o.style = "takeaway";
    const fill = String(it.fill || "").toLowerCase();
    if (FILLS.includes(fill) && fill !== "none") o.fill = fill;
    if (it.align === "center" || it.align === "right" || it.align === "left") o.align = it.align;
    if (it.size === "large") o.size = "large";
    if (it.level === 1 || it.level === 2) o.level = it.level;
    if (it.bold) o.bold = true;
    return o;
  }
  function normCard(it) {
    if (typeof it === "string") {
      const m = it.match(/^\*\*(.+?)\*\*[:：]?\s*(.*)$/);
      return m ? { type: "cell", head: m[1], text: m[2] } : { type: "cell", head: "", text: it };
    }
    const o = { type: "cell", head: stripBold(str(it.head || it.title || it.label)), text: str(it.text || it.body || it.description), highlight: !!it.highlight };
    const items = it.items || it.bullets;
    if (Array.isArray(items) && items.length) {
      o.items = normItems(items);
      if (!o.text && o.items.length === 1 && !Array.isArray(o.items[0])) {
        o.text = o.items[0];
        delete o.items;
      }
    }
    if (it.icon) o.icon = str(it.icon).toLowerCase();
    return o;
  }
  /** 数値の無い figure は描かない(空の軸だけが残るのを防ぐ)。LLM が図を宣言しつつデータを別の部品に書いたときに起きる */
  function dropEmptyFigure(f) {
    const num = (v) => typeof v === "number" && !isNaN(v);
    if (f.type === "bars" || f.type === "column") return (f.items || []).some((it) => num(it.value) && it.value !== 0) ? f : null;
    if (f.type === "line") return (f.series || []).some((s) => (s.values || []).some(num)) ? f : null;
    if (f.type === "stacked") return (f.items || []).some((it) => (it.values || []).some(num)) ? f : null;
    if (f.type === "kpi") return f.value ? f : null;
    return f;
  }
  /** 観点(列)が 1 つ以下の格子は「格子」ではないので箇条書き(cell.items)に落とす */
  function degradeThinTable(t) {
    if (!t || t.type !== "table" || t.headShape || t.axes || t.colGroups) return t;
    const cols = (t.colHeaders || []).length;
    const wide = (t.rows || []).some((r) => (r.cells || []).length > 1);
    if (cols > 1 || wide) return t;
    const items = (t.rows || [])
      .map((r) => {
        const head = str(r.head);
        const cell = cellText((r.cells || [])[0]);
        return head && cell ? "**" + head + "**: " + cell : head || cell;
      })
      .filter(Boolean);
    return items.length ? { type: "cell", items } : t;
  }
  /** ガント: periods(期間ラベル)と tasks(start/end は期間ラベルか index)を index に解決する */
  function normGantt(n) {
    let periods = (n.periods || n.columns || n.labels || []).map(str).filter(Boolean);
    const rawTasks = (n.tasks || n.rows || n.items || []).slice(0, 14);
    const idx = (v, from) => {
      if (v == null || v === "") return -1;
      if (typeof v === "number") return Math.round(v);
      const t = String(v).trim();
      const f = from || 0;
      let i = periods.findIndex((p, k) => k >= f && p === t);
      if (i < 0) i = periods.findIndex((p, k) => k >= f && p && (t.includes(p) || p.includes(t)));
      if (i < 0 && f > 0) return idx(v, 0);
      return i;
    };
    if (!periods.length) {
      // 期間が無ければタスクの開始・終了ラベルから順に作る
      const seen = [];
      rawTasks.forEach((t) => [t.start, t.end].forEach((v) => { if (v != null && v !== "" && typeof v !== "number" && !seen.includes(String(v))) seen.push(String(v)); }));
      periods = seen;
    }
    if (!periods.length) return null;
    const tasks = rawTasks
      .map((t) => {
        const label = str(t.label || t.name || t.head || t.task);
        let s0 = idx(t.start != null ? t.start : t.from);
        let e0 = idx(t.end != null ? t.end : t.to, Math.max(0, s0)); // 終了は開始以降で探す
        if (s0 < 0 && e0 >= 0) s0 = e0;
        if (e0 < 0 && s0 >= 0) e0 = s0;
        if (s0 < 0) return null;
        return { label, start: Math.min(s0, e0), end: Math.max(s0, e0), highlight: !!t.highlight, note: str(t.note) };
      })
      .filter(Boolean);
    if (!tasks.length) return null;
    const milestones = (n.milestones || []).slice(0, 6).map((m) => ({ label: str(m.label || m.name), at: idx(m.at != null ? m.at : m.when) })).filter((m) => m.at >= 0);
    const groups = Array.isArray(n.groups || n.periodGroups) ? (n.groups || n.periodGroups).map((g) => ({ text: str(g.text || g.label), span: Math.max(1, Number(g.span) || 1) })) : undefined;
    const today = n.today != null ? idx(n.today) : -1;
    return { type: "gantt", periods, tasks, milestones, groups, today: today >= 0 ? today : undefined };
  }
  /** 体制図: root{label, sub, children[]} または nodes[{label, sub, parent}] をツリーに */
  function normOrg(n) {
    const lim = { count: 0 };
    const build = (nd, depth) => {
      if (!nd || typeof nd !== "object" || depth > 4 || lim.count >= 18) return null;
      lim.count++;
      const o = { label: str(nd.label || nd.name || nd.head).slice(0, 24), sub: str(nd.sub || nd.text || nd.desc).slice(0, 40), highlight: !!nd.highlight };
      const note = str(nd.note || nd.detail).slice(0, 120); // 論点ツリーの末端に付ける補足説明(任意)
      if (note) o.note = note;
      const kids = (nd.children || nd.members || []).map((k) => (typeof k === "string" ? { label: k } : k)).map((k) => build(k, depth + 1)).filter(Boolean);
      if (kids.length) o.children = kids;
      return o.label ? o : null;
    };
    let root = n.root ? build(n.root, 0) : null;
    if (!root && Array.isArray(n.nodes)) {
      const byId = {};
      n.nodes.forEach((x, i) => (byId[str(x.id || x.label)] = Object.assign({ _i: i }, x, { children: [] })));
      let top = null;
      n.nodes.forEach((x) => {
        const me = byId[str(x.id || x.label)];
        const p = x.parent != null ? byId[str(x.parent)] : null;
        if (p) p.children.push(me);
        else if (!top) top = me;
      });
      root = top ? build(top, 0) : null;
    }
    return root ? { type: "org", root } : null;
  }
  function normTable(n, type) {
    const cellOf = (x) => {
      if (typeof x === "string" || typeof x === "number") return String(x);
      if (!x || typeof x !== "object") return "";
      const o = { text: str(x.text || x.value) };
      const fill = String(x.fill || "").toLowerCase();
      if (FILLS.includes(fill) && fill !== "none") o.fill = fill;
      if (x.highlight) o.highlight = true;
      if (Array.isArray(x.items) && x.items.length) o.items = x.items.map(str);
      return o.fill || o.highlight || o.items ? o : o.text;
    };
    const out = {
      type,
      corner: str(n.corner || n.rowHeader),
      colHeaders: (n.colHeaders || n.headers || n.columns || []).map(str),
      rows: (n.rows || []).map((r) => ({ head: str(r.head || r.label || r.title), cells: (r.cells || r.values || []).map(cellOf), highlight: !!r.highlight })),
      highlightCol: typeof n.highlightCol === "number" ? n.highlightCol : undefined,
    };
    if (type === "table") {
      if (n.headFill === "dark" || n.headStyle === "dark") out.headFill = "dark";
      else if (n.headFill === "none" || n.headFill === "light") out.headFill = n.headFill;
      if (n.rowHeadFill === "none") out.rowHeadFill = "none";
      if (n.numbered === true) out.numbered = true;
      if (n.headShape === "chevron" || n.headShape === "chevrons") out.headShape = "chevron";
      if (Array.isArray(n.colGroups) && n.colGroups.length) out.colGroups = n.colGroups.map((g) => ({ text: str(g.text || g.head), span: Math.max(1, Number(g.span) || 1) }));
      if (Array.isArray(n.rowGroups) && n.rowGroups.length) out.rowGroups = n.rowGroups.map((g) => ({ text: str(g.text || g.head), span: Math.max(1, Number(g.span) || 1) }));
      // 軸ラベル(axes)は 2×2 の格子にだけ意味がある。行や列が 2 でない格子に付いてきたら外す(回転ラベルが行名に重なる)
      if (n.axes && (n.axes.x || n.axes.y) && out.rows.length === 2 && (out.colHeaders || []).length <= 2 && out.rows.every((rw) => (rw.cells || []).length === 2)) out.axes = { x: str(n.axes.x), y: str(n.axes.y) };
    }
    return out;
  }
  function itemStr(x) {
    if (typeof x === "string") return x;
    if (x && typeof x === "object") return itemToText({ head: str(x.head || x.title || x.label), text: str(x.text || x.body || x.description) });
    return str(x);
  }
  function str(v) {
    if (v == null) return "";
    let t = String(v);
    // LLM の崩れ: JSON の断片(閉じ括弧の連続 / キー": [ / 内部フラグ風の英字列)が本文に混ざることがある
    t = t.replace(/(\s*[\]\}]){3,}[\s\S]*$/, "").replace(/\b[A-Za-z]{16,}\?[A-Za-z]+:(true|false)\}?/g, "");
    t = t.replace(/[A-Za-z_]*"\s*:\s*[\[\{][\s\S]*$/, "").replace(/[\[\]\{\}]{2,}[\s\S]*$/, "");
    if (/["\{\}\[\]]/.test(t) && !/[\u3040-\u30ff\u4e00-\u9fff]/.test(t)) return ""; // 記号だけの残骸は捨てる
    return t.trim();
  }

  const VERSION = "5.0";
  return { layout, normalizeSpec, normalizeGeneratedSpec, compositionKey, parseRuns, estimateLines, naturalHeight, tint, cleanTitle, DEFAULT_PALETTE, FONT: BASE_FONT, STYLE, LEAF_TYPES, VERSION };
});
