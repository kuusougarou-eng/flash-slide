/* global PowerPoint, Office */
/**
 * render.js — Office.js 描画層。
 *
 *  prepare()            … スライドサイズ・選択スライド・マスタ/レイアウト・参照スライドの解析(profile)を 1 回で取得
 *  renderSpecs(specs)   … spec 配列 → 選択スライド直後に順次挿入
 *
 * 挿入方式は 2 種類を自動選択:
 *  - clone : 選択スライドを複製(exportAsBase64 → insertSlidesFromBase64)し、タイトル/リード/出典は既存シェイプを再利用、
 *            フッター・ページ番号・ロゴ・罫線などの装飾は保持、本文だけ差し替える。コピペ運用デッキでも前後の見た目が揃う。
 *  - layout: 参照にできるスライドが無い(表紙・空)場合。同じマスタの白紙レイアウトでスライドを追加し、既定レイアウトで描く。
 *
 * sync 回数は最小限(準備 2〜3 回、描画 3〜4 回/枚)。
 */
(function (root) {
  "use strict";

  const VALIGN = { top: "Top", middle: "Middle", bottom: "Bottom" };
  const HALIGN = { left: "Left", center: "Center", right: "Right" };
  const TEXT_TYPES = ["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"];

  function supports(ver) {
    try {
      return Office.context.requirements.isSetSupported("PowerPointApi", ver);
    } catch (_) {
      return false;
    }
  }
  function isWindows() {
    try {
      return Office.context.platform === Office.PlatformType.PC;
    } catch (_) {
      return true;
    }
  }
  function pickLayoutId(layouts) {
    if (!layouts || !layouts.length) return undefined;
    const byName = (re) => layouts.find((l) => re.test(l.name || ""));
    const hit = byName(/^(blank|白紙|空白|leer|vide)/i) || byName(/blank|白紙/i) || byName(/title only|タイトルのみ/i) || layouts[layouts.length - 1];
    return hit.id;
  }
  const geoKey = (s) => [s.type, Math.round(s.left), Math.round(s.top), Math.round(s.width), Math.round(s.height)].join("|");

  // =====================================================================
  //  準備: サイズ・選択・マスタ・参照スライド解析
  // =====================================================================
  async function prepare() {
    return PowerPoint.run(async (context) => {
      const pres = context.presentation;
      const has110 = supports("1.10");
      const has18 = supports("1.8");
      const info = { width: 960, height: 540, selectedId: null, selectedIndex: null, masterId: null, layoutId: null, has18, has110, mode: "layout", profile: null };

      const masters = pres.slideMasters;
      masters.load("items/id,items/name,items/layouts/items/id,items/layouts/items/name");
      let ps = null;
      if (has110) {
        ps = pres.pageSetup;
        ps.load("slideWidth,slideHeight");
      }
      const slides = pres.slides;
      slides.load("items/id");
      const sel = pres.getSelectedSlides();
      sel.load("items/id");
      await context.sync();

      if (ps && ps.slideWidth) {
        info.width = ps.slideWidth;
        info.height = ps.slideHeight;
      }
      info.allIds = slides.items.map((s) => s.id);
      if (sel.items.length) {
        info.selectedId = sel.items[0].id;
        info.selectedIndex = info.allIds.indexOf(info.selectedId);
      }

      let master = masters.items[0];
      let refShapes = [];
      if (info.selectedId) {
        const s = slides.getItem(info.selectedId);
        s.slideMaster.load("id");
        s.layout.load("id,name");
        s.shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
        await context.sync();
        const m = masters.items.find((x) => x.id === s.slideMaster.id);
        if (m) master = m;
        info.refLayoutName = s.layout.name;
        refShapes = await loadShapeDetails(context, s.shapes.items, has18);
        // レイアウト側のプレースホルダ既定サイズ(タイトル/サブタイトル)。複製の連鎖で縮んだ実サイズを引きずらないための基準
        if (has18) {
          try {
            const ls = s.layout.shapes;
            ls.load("items/type,items/placeholderFormat/type,items/textFrame/textRange/font/size");
            await context.sync();
            const defaults = {};
            for (const l of ls.items) {
              try {
                if (l.type === "Placeholder" && l.placeholderFormat && l.placeholderFormat.type && l.textFrame.textRange.font.size) defaults[l.placeholderFormat.type] = l.textFrame.textRange.font.size;
              } catch (_) {}
            }
            for (const sh of refShapes) if (sh.placeholder && defaults[sh.placeholder]) sh.layoutFontSize = defaults[sh.placeholder];
            info.layoutDefaults = defaults;
          } catch (e) {
            info.layoutDefaultsError = String((e && e.message) || e);
          }
        }
      }
      if (master) {
        info.masterId = master.id;
        info.masterName = master.name;
        info.layouts = master.layouts.items.map((l) => ({ id: l.id, name: l.name }));
        info.layoutId = pickLayoutId(info.layouts);
      }
      info.refShapes = refShapes;
      const analysis = analyzeReference(refShapes, info.width, info.height);
      info.profile = analysis.profile;
      info.plan = analysis.plan; // clone 時の keep/reuse 判定
      info.mode = analysis.profile ? "clone" : "layout";
      info.analysisNote = analysis.note;
      return info;
    });
  }

  /**
   * 参照スライドのシェイプ詳細(テキスト・フォント・塗り)を読み込む。
   * foregroundColor は fill.type が Solid のときだけ読む(それ以外で読むと例外 → バッチ全体が失われるため)。
   * 各 sync は個別に try で保護し、1 つが失敗しても残りの情報は活かす。
   */
  async function loadShapeDetails(context, items, has18) {
    const out = items.map((sh) => ({ id: sh.id, name: sh.name, type: sh.type, left: sh.left, top: sh.top, width: sh.width, height: sh.height }));
    const textIdx = [];
    // --- sync 1: テキスト・フォント・段落・fill.type(色は読まない) ---
    items.forEach((sh, i) => {
      if (!TEXT_TYPES.includes(sh.type)) return;
      textIdx.push(i);
      sh.load(
        "textFrame/hasText,textFrame/textRange/text,textFrame/textRange/font/size,textFrame/textRange/font/bold,textFrame/textRange/font/color,textFrame/textRange/font/name,textFrame/textRange/paragraphFormat/horizontalAlignment,fill/type"
      );
      if (has18 && sh.type === "Placeholder") sh.load("placeholderFormat/type");
    });
    try {
      await context.sync();
    } catch (_) {
      // 失敗時はシェイプ単位で最小限を取り直す
      for (const i of textIdx) {
        try {
          items[i].load("textFrame/hasText,textFrame/textRange/text,textFrame/textRange/font/size,textFrame/textRange/font/bold,textFrame/textRange/font/color,textFrame/textRange/font/name,fill/type");
          await context.sync();
        } catch (_) {}
      }
    }
    const solidIdx = [];
    for (const i of textIdx) {
      const sh = items[i],
        o = out[i];
      try {
        o.hasText = sh.textFrame.hasText;
        o.text = o.hasText ? sh.textFrame.textRange.text : "";
        o.fontSize = sh.textFrame.textRange.font.size;
        o.bold = sh.textFrame.textRange.font.bold;
        o.color = sh.textFrame.textRange.font.color;
        o.fontName = sh.textFrame.textRange.font.name;
      } catch (_) {}
      try {
        o.align = sh.textFrame.textRange.paragraphFormat.horizontalAlignment;
      } catch (_) {}
      try {
        o.fillType = sh.fill.type;
        if (o.fillType === "Solid") solidIdx.push(i);
      } catch (_) {}
      try {
        if (sh.type === "Placeholder") o.placeholder = sh.placeholderFormat.type;
      } catch (_) {}
    }
    // --- sync 2: Solid 塗りの色だけ読む ---
    if (solidIdx.length) {
      solidIdx.forEach((i) => items[i].fill.load("foregroundColor"));
      try {
        await context.sync();
        for (const i of solidIdx) {
          try {
            out[i].fillColor = items[i].fill.foregroundColor;
          } catch (_) {}
        }
      } catch (_) {}
    }
    return out;
  }

  // =====================================================================
  //  参照スライド解析 → profile(レイアウト規約) と plan(複製時の扱い)
  // =====================================================================
  function analyzeReference(shapes, W, H) {
    const none = { profile: null, plan: null, note: "参照なし → 既定レイアウト" };
    if (!shapes || !shapes.length) return none;
    // Flash Slide 自身が描いた本文(FS_*)は「前スライドの内容」なので、参照解析の対象から外す(直前の生成物を参照にしても崩れない)
    const own = (s) => /^FS_/.test(s.name || "");
    const inSlide = shapes.filter((s) => !own(s) && s.left + s.width > 0 && s.top + s.height > 0 && s.left < W && s.top < H);
    const texts = inSlide.filter((s) => s.hasText && (s.text || "").trim());
    const fs = (s) => s.fontSize || Math.min(40, Math.max(10, s.height * 0.5));

    // --- タイトル候補 ---
    const phTitle = texts.find((s) => s.placeholder === "Title");
    const centerTitle = inSlide.find((s) => s.placeholder === "CenterTitle");
    let title = phTitle || null;
    if (!title) {
      const cands = texts.filter((s) => s.top < H * 0.25 && s.width > W * 0.4 && (s.text || "").length < 140);
      cands.sort((a, b) => fs(b) - fs(a) || a.top - b.top);
      title = cands[0] || null;
    }
    if (!title) return { profile: null, plan: null, note: "タイトルが見つからない → 既定レイアウト" };
    // 表紙: CenterTitle / 下寄せタイトル / 「タイトル + サブタイトルだけ」で本文が無い(Subtitle プレースホルダ自体はリード文として使うテンプレートが多い)
    const bodyish = texts.filter((s) => s !== title && s.placeholder !== "Subtitle" && s.placeholder !== "Footer" && s.placeholder !== "SlideNumber" && s.placeholder !== "Date" && !isFixedFooterText(s.text) && s.top < H * 0.85);
    const hasOwn = shapes.some(own); // 自分が生成したスライド(本文は FS_*)は表紙ではない
    const onlySubtitle = !hasOwn && inSlide.some((s) => s.placeholder === "Subtitle") && bodyish.length === 0;
    if (centerTitle || title.top > H * 0.3 || onlySubtitle) return { profile: null, plan: null, note: "表紙/中表紙と判定 → 既定レイアウト" };

    const titleBottom = title.top + title.height;
    // --- リード文候補: タイトル直下・幅広・短文 ---
    const leadCands = texts
      .filter((s) => s !== title && s.top >= titleBottom - 6 && s.top < H * 0.45 && s.width > W * 0.4 && s.height < H * 0.22 && (s.text || "").length <= 160 && fs(s) <= fs(title))
      .sort((a, b) => a.top - b.top);
    let lead = leadCands[0] || null;
    // 箇条書き本文プレースホルダ(Body/Content)はリードではない
    if (lead && (lead.placeholder === "Body" || lead.placeholder === "Content" || /\n/.test(lead.text || ""))) lead = null;
    // 空の Subtitle プレースホルダ(前の生成がリード無しだった等)もリード枠として保持する。消すと、それ以降の複製でリード文の置き場が無くなる
    if (!lead) {
      const emptySub = inSlide.find((s) => s.placeholder === "Subtitle" && !(s.text || "").trim() && s.top >= titleBottom - 6 && s.top < H * 0.45 && s.width > W * 0.4 && s.height < H * 0.22);
      if (emptySub) lead = emptySub;
    }

    // --- フッター帯: 下端 12% ---
    const footerZoneTop = H * 0.88;
    const footerShapes = inSlide.filter((s) => s.top >= footerZoneTop - 2 || ["Footer", "SlideNumber", "Date"].includes(s.placeholder));    // 出典行候補: 下部 25% にある小さめ文字の横長テキスト(フッター定型文は除く)。最も下にあるものを出典とする
    const srcCands = texts
      .filter((s) => s !== title && s !== lead && s.top > H * 0.72 && fs(s) <= 13 && s.width > W * 0.3 && !isFixedFooterText(s.text) && !["Footer", "SlideNumber", "Date"].includes(s.placeholder))
      .sort((a, b) => b.top - a.top);
    const footnote = srcCands[0] || null;

    // --- 保持する装飾: ヘッダ帯(タイトル/リードの下端まで)・フッタ帯・左右の余白帯にあるものだけ。
    //     本文領域内の細線・画像・バーは前スライドの内容なので保持しない ---
    const headerBottom = (lead ? lead.top + lead.height : titleBottom) + 12;
    const marginL = Math.min(title.left, lead ? lead.left : title.left) - 4;
    const marginR = Math.max(title.left + title.width, lead ? lead.left + lead.width : 0) + 4;
    const keep = new Set();
    inSlide.forEach((s) => {
      const area = (s.width * s.height) / (W * H);
      const inHeader = s.top + s.height <= headerBottom;
      const inFooter = s.top >= footerZoneTop - 2;
      const inMargin = s.left + s.width <= marginL || s.left >= marginR;
      const zone = inHeader || inFooter || inMargin;
      // フッタ帯のテキストは、定型文(©・Confidential・ページ番号)かプレースホルダだけ保持。前スライドの注記などは消す
      const footerKeep =
        footerShapes.includes(s) &&
        (!s.hasText ||
          isFixedFooterText(s.text) ||
          ["Footer", "SlideNumber", "Date"].includes(s.placeholder) ||
          s === footnote ||
          (s.fillType === "Solid" && (s.text || "").length <= 10)); // ロゴ・チップ風の小さな塗り図形
      const decor =
        footerKeep ||
        (zone && (s.type === "Image" || s.type === "Graphic") && area < 0.12) ||
        (zone && s.type === "Line") ||
        (zone && !s.hasText && (s.height <= 6 || s.width <= 6)) || // 罫線的な細い矩形
        (inHeader && !s.hasText) || // タイトル帯の背景
        (inHeader && s.hasText && s !== title && s !== lead && fs(s) <= 12); // 章番号・ラベルなど
      if (decor) keep.add(s.id);
    });
    keep.add(title.id);
    if (lead) keep.add(lead.id);
    if (footnote) keep.add(footnote.id);

    // --- 本文領域 ---
    const bodyLeft = Math.min(title.left, lead ? lead.left : title.left);
    const bodyRight = Math.max(title.left + title.width, lead ? lead.left + lead.width : 0);
    const topAnchor = lead ? lead.top + lead.height : titleBottom;
    let bottomLimit = H - Math.max(24, H * 0.06);
    const lowerKept = inSlide.filter((s) => keep.has(s.id) && s !== title && s !== lead && s.top > H * 0.6);
    if (lowerKept.length) bottomLimit = Math.min(bottomLimit, Math.min.apply(null, lowerKept.map((s) => s.top)) - 6);
    if (footnote) bottomLimit = Math.min(bottomLimit, footnote.top - 6);
    const body = { x: bodyLeft, y: topAnchor + 10, w: bodyRight - bodyLeft, h: bottomLimit - (topAnchor + 10) };
    if (body.h < 150 || body.w < W * 0.4) return { profile: null, plan: null, note: "本文領域が狭すぎる → 既定レイアウト" };

    // --- 強調色: 参照スライドで使われている有彩色の中で最頻 ---
    const colorCount = {};
    inSlide.forEach((s) => {
      const cands = [s.fillType === "Solid" ? s.fillColor : null, s.hasText ? s.color : null];
      cands.forEach((c) => {
        if (c && isChromatic(c)) colorCount[c.toUpperCase()] = (colorCount[c.toUpperCase()] || 0) + 1;
      });
    });
    const accent = Object.keys(colorCount).sort((a, b) => colorCount[b] - colorCount[a])[0] || null;
    // 本文の文字サイズ: 本文領域にある通常テキストの最頻値(同数なら小さい方)。見出しは太字の最頻値
    const bodyCands = texts.filter(
      (s) => s !== title && s !== lead && s !== footnote && !["Footer", "SlideNumber", "Date"].includes(s.placeholder) && !isFixedFooterText(s.text) && s.fontSize >= 8 && s.fontSize <= 28 && s.top >= topAnchor - 20 && s.top < bottomLimit + 20
    );
    const mode = (list) => {
      const cnt = {};
      list.forEach((s) => (cnt[Math.round(s.fontSize)] = (cnt[Math.round(s.fontSize)] || 0) + 1));
      const keys = Object.keys(cnt).map(Number).sort((a, b) => cnt[b] - cnt[a] || a - b);
      return keys.length ? keys[0] : null;
    };
    const bodyFontSize = mode(bodyCands.filter((s) => !s.bold)) || mode(bodyCands);
    const headFontSize = mode(bodyCands.filter((s) => s.bold)) || null;
    const bodyText = texts.find((s) => s !== title && s !== lead && s.fontName && !/^\+/.test(s.fontName));
    const fontName = (bodyText && bodyText.fontName) || (title.fontName && !/^\+/.test(title.fontName) ? title.fontName : null);

    const profile = {
      title: { x: title.left, y: title.top, w: title.width, h: title.height, fontSize: title.layoutFontSize || title.fontSize || undefined, bold: title.bold !== false, color: title.color, align: alignOf(title.align), reuse: true },
      lead: lead ? { x: lead.left, y: lead.top, w: lead.width, h: lead.height, fontSize: lead.layoutFontSize || lead.fontSize || undefined, color: lead.color, align: alignOf(lead.align), reuse: true } : null,
      body,
      footnote: footnote ? { x: footnote.left, y: footnote.top, w: footnote.width, h: footnote.height, fontSize: footnote.fontSize || 10, reuse: true } : null,
      accent,
      fontName,
      bodyFontSize,
      headFontSize,
      footnoteFontSize: footnote ? Math.round(footnote.fontSize || 0) || null : null,
    };
    const plan = { keepKeys: {}, titleKey: geoKey(title), leadKey: lead ? geoKey(lead) : null, footnoteKey: footnote ? geoKey(footnote) : null };
    inSlide.forEach((s) => {
      if (keep.has(s.id)) plan.keepKeys[geoKey(s)] = true;
    });
    const note = `参照: タイトル ${Math.round(title.fontSize || 0)}pt${lead ? " / リード有" : " / リード無"}${footnote ? " / 出典行有" : ""} / 本文 ${Math.round(body.w)}×${Math.round(body.h)}pt${bodyFontSize ? " / 本文 " + bodyFontSize + "pt" : ""}${headFontSize ? " / 見出し " + headFontSize + "pt" : ""} / 保持 ${keep.size} 形${fontName ? " / " + fontName : ""}${accent ? " / 色 " + accent : ""}`;
    return { profile, plan, note };
  }
  /** paragraphFormat.horizontalAlignment は段落が混在すると文字列以外(null / オブジェクト)になりうる */
  function alignOf(a) {
    const v = typeof a === "string" ? a.toLowerCase() : "";
    return ["left", "center", "right", "justify"].includes(v) ? v : "left";
  }
  function isFixedFooterText(t) {
    return /©|\(c\)|confidential|all rights reserved|社外秘|禁複製|^\s*\d+\s*(\/\s*\d+)?\s*$/i.test(t || "");
  }
  function isChromatic(hex) {
    const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex || "");
    if (!m) return false;
    const r = parseInt(m[1], 16), g = parseInt(m[2], 16), b = parseInt(m[3], 16);
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    return max - min > 40 && max > 60; // 彩度があり黒すぎない
  }

  // =====================================================================
  //  描画プリミティブ → シェイプ
  // =====================================================================
  function applyRect(shapes, p, ctx) {
    const G = PowerPoint.GeometricShapeType;
    const type =
      { chevron: G.chevron, homePlate: G.homePlate, rightArrow: G.rightArrow, downArrow: G.downArrow, roundRect: G.roundRectangle, ellipse: G.ellipse, triangle: G.triangle }[p.shape] ||
      G.rectangle;
    const shape = shapes.addGeometricShape(type, { left: p.x, top: p.y, width: Math.max(1, p.w), height: Math.max(1, p.h) });
    shape.name = "FS_" + (p.text ? "text" : "rect");
    if (p.fill) shape.fill.setSolidColor(p.fill);
    else shape.fill.clear();
    if (p.rotation) {
      try {
        shape.rotation = p.rotation;
      } catch (_) {}
    }
    if (p.line) {
      shape.lineFormat.visible = true;
      shape.lineFormat.color = p.line;
      shape.lineFormat.weight = p.lineWeight || 0.75;
    } else shape.lineFormat.visible = false;

    if (p.text) {
      const tf = shape.textFrame;
      tf.wordWrap = p.wrap !== false;
      const pad = p.pad == null ? 10 : p.pad;
      tf.leftMargin = pad;
      tf.rightMargin = pad;
      tf.topMargin = Math.min(pad, 8);
      tf.bottomMargin = Math.min(pad, 8);
      tf.verticalAlignment = VALIGN[p.valign] || "Top";
      tf.autoSizeSetting = p.autofit === "shrink" ? "AutoSizeTextToFitShape" : "AutoSizeNone";
      const tr = tf.textRange;
      if (p.fontName) tr.font.name = p.fontName; // フォントは text 設定前に指定(アイコンフォントの豆腐化を防ぐ)
      tr.text = p.text;
      tr.font.size = p.fontSize || 18;
      tr.font.bold = !!p.bold;
      if (p.color) tr.font.color = p.color;
      if (p.fontName) tr.font.name = p.fontName;
      tr.paragraphFormat.horizontalAlignment = HALIGN[p.align] || "Left";
      tr.paragraphFormat.bulletFormat.visible = !!p.bullets;
      if (!p.bold && p.boldRanges && p.boldRanges.length) {
        for (const [s, l] of p.boldRanges) if (l > 0 && s + l <= p.text.length) tr.getSubstring(s, l).font.bold = true;
      }
    }
    return shape;
  }

  /** 罫線は細い塗り矩形で描く */
  /** 罫線は本物の線分(addLine)。使えない環境では細い矩形にフォールバック */
  function applyLine(shapes, p) {
    const w = Math.abs(p.x2 - p.x1), h = Math.abs(p.y2 - p.y1), weight = p.weight || 1;
    try {
      // 幅/高さに 0 を渡すと Office の既定サイズ(75pt)に置き換えられて斜線になるため、0 は 0.1 にする
      const ln = shapes.addLine(PowerPoint.ConnectorType.straight, { left: Math.min(p.x1, p.x2), top: Math.min(p.y1, p.y2), width: w || 0.1, height: h || 0.1 });
      ln.name = "FS_rule";
      ln.lineFormat.color = p.color || "#7F7F7F";
      ln.lineFormat.weight = weight;
      ln.lineFormat.visible = true;
      return ln;
    } catch (_) {
      const shape = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: Math.min(p.x1, p.x2),
        top: Math.min(p.y1, p.y2) - (h === 0 ? weight / 2 : 0),
        width: w === 0 ? weight : w,
        height: h === 0 ? weight : h,
      });
      shape.name = "FS_rule";
      shape.fill.setSolidColor(p.color || "#7F7F7F");
      shape.lineFormat.visible = false;
      return shape;
    }
  }

  /** ネイティブ表(1.8+)。ヘッダー塗りなし+太い強調色罫、行間は薄罫、最終行の下は罫なし、縦罫なし */
  function applyTable(shapes, p) {
    const nRows = p.cells.length, nCols = p.cells[0].length;
    // 縦罫線は白の細線で「見えない」扱い(透明指定は環境により InvalidArgument になるため)
    const none = { color: "#FFFFFF", dashStyle: "Solid", weight: 0.5 };
    const border = (b) => (b ? { color: b.color, dashStyle: "Solid", weight: Math.max(0.5, b.weight) } : none);
    const values = p.cells.map((row) => row.map((cell) => cell.text || ""));
    const specific = p.cells.map((row, i) =>
      row.map((cell, j) => {
        const o = {
          font: { size: Math.round(cell.fontSize || (i === 0 ? p.headFontSize : p.fontSize)), color: cell.color || "#1A1A1A", bold: !!cell.bold },
          horizontalAlignment: HALIGN[cell.align] || "Left",
          verticalAlignment: "Middle",
          borders: { top: border(i > 0 ? p.cells[i - 1][j].borderBottom : null), bottom: border(cell.borderBottom), left: none, right: none },
        };
        if (cell.fill) o.fill = { color: cell.fill };
        if (p.fontName) o.font.name = p.fontName;
        return o;
      })
    );
    const table = shapes.addTable(nRows, nCols, {
      left: p.x,
      top: p.y,
      values,
      columns: p.colWidths.map((w) => ({ columnWidth: Math.round(w) })),
      rows: p.rowHeights.map((h) => ({ rowHeight: Math.round(h) })),
      uniformCellProperties: {},
      specificCellProperties: specific,
    });
    table.name = "FS_table";
    // セル内の **太字** ラン(getSubstring は表セルに使えないため、セル全体を太字にする代替は行わない)
    return table;
  }

  /** シェイプ系プリミティブを描く。画像(アイコン)は後段で挿入するため返す */
  function drawPrims(shapes, prims, info) {
    const images = [];
    for (const p of prims) {
      if (p.kind === "rect") applyRect(shapes, p, info);
      else if (p.kind === "line") applyLine(shapes, p);
      else if (p.kind === "table") {
        if (info.has18) applyTable(shapes, p);
        else for (const f of p.fallback) applyRect(shapes, f, info);
      } else if (p.kind === "image") images.push(p);
    }
    return images;
  }

  /**
   * ピクトグラム画像を挿入する。サーバ(/api/icons)で Tabler Icons を色付き PNG に描画し、
   * Office 共通 API の setSelectedDataAsync(image) で「現在のスライド」に位置指定で貼る。
   * 失敗しても本文の描画は完了しているので握りつぶす。
   */
  async function insertImages(images) {
    if (!images.length) return 0;
    let map = {};
    try {
      const res = await fetch("/api/icons", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ items: images.map((p) => ({ name: p.name, color: p.color })), size: 256 }),
      });
      map = (await res.json()).icons || {};
    } catch (_) {
      return 0;
    }
    let n = 0;
    for (const p of images) {
      const key = `${p.name}|${String(p.color || "#404040").toUpperCase()}`;
      const b64 = map[key];
      if (!b64) continue;
      const ok = await new Promise((resolve) => {
        try {
          Office.context.document.setSelectedDataAsync(
            b64,
            { coercionType: Office.CoercionType.Image, imageLeft: p.x, imageTop: p.y, imageWidth: p.w, imageHeight: p.h },
            (r) => resolve(r.status === Office.AsyncResultStatus.Succeeded)
          );
        } catch (_) {
          resolve(false);
        }
      });
      if (ok) n++;
    }
    return n;
  }

  // =====================================================================
  //  スライド生成
  // =====================================================================
  /**
   * @param specs  normalize 済み spec の配列(1〜2 枚)
   * @param opts   { prepared, palette, preview, layoutId }
   * @returns {Promise<Array<{slideId, ms, png, warnings, width, height, mode}>>}
   */
  /**
   * 差し替え用: このアドインが挿入したスライドを ID で削除する。存在しない ID(ユーザーが消した等)は無視。
   * @returns {Promise<number>} 削除できた枚数
   */
  async function deleteSlides(ids) {
    ids = (ids || []).filter(Boolean);
    if (!ids.length) return 0;
    let n = 0;
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();
      const have = new Set(slides.items.map((s) => s.id));
      for (const id of ids) {
        if (!have.has(id)) continue;
        slides.getItem(id).delete();
        n++;
      }
      await context.sync();
    });
    return n;
  }

  /** 生成/差し替え後に、挿入したスライドを表示・選択する(PowerPointApi 1.5)。失敗しても無視 */
  async function selectSlides(ids) {
    ids = (ids || []).filter(Boolean);
    if (!ids.length || !supports("1.5")) return false;
    try {
      await PowerPoint.run(async (context) => {
        context.presentation.setSelectedSlides(ids);
        await context.sync();
      });
      return true;
    } catch (_) {
      return false;
    }
  }

  async function renderSpecs(specs, opts) {
    opts = opts || {};
    const info = opts.prepared || (await prepare());
    const results = [];
    let afterId = info.selectedId;
    let cloneB64 = null;

    for (const spec of specs) {
      const t0 = performance.now();
      const useClone = info.mode === "clone" && info.selectedId && info.has18;
      const lay = root.SlideLayout.layout(spec, { width: info.width, height: info.height, palette: opts.palette, profile: useClone ? info.profile : null });

      const r = await PowerPoint.run(async (context) => {
        const pres = context.presentation;
        const slides = pres.slides;
        let newSlide, newId;
        let preDelete = null;

        if (useClone) {
          if (!cloneB64) {
            const ref = slides.getItem(info.selectedId);
            const b = ref.exportAsBase64();
            await context.sync();
            cloneB64 = b.value;
          }
          pres.insertSlidesFromBase64(cloneB64, { formatting: "UseDestinationTheme", targetSlideId: afterId });
          slides.load("items/id");
          await context.sync();
          const ids = slides.items.map((s) => s.id);
          newId = ids[ids.indexOf(afterId) + 1];
          newSlide = slides.getItem(newId); // ID で再取得(インデックス由来のプロキシを使わない)
          newSlide.shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
          await context.sync();
          // 本文を消し、タイトル/リード/出典は文字だけ差し替える
          for (const sh of newSlide.shapes.items) {
            const k = geoKey(sh);
            if (k === info.plan.titleKey) {
              sh.textFrame.textRange.text = lay.header.titleText;
              if (lay.header.titleFontSize) sh.textFrame.textRange.font.size = lay.header.titleFontSize;
            } else if (k === info.plan.leadKey) {
              sh.textFrame.textRange.text = lay.header.leadText || "";
              if (lay.header.leadFontSize && lay.header.leadText) sh.textFrame.textRange.font.size = lay.header.leadFontSize;
            } else if (k === info.plan.footnoteKey) sh.textFrame.textRange.text = lay.header.footnoteText || "";
            else if (!info.plan.keepKeys[k]) sh.delete();
          }
        } else {
          const addOpts = {};
          if (info.masterId) addOpts.slideMasterId = info.masterId;
          // ユーザー選択のレイアウト(現在のマスタに存在するときだけ)を自動選択より優先
          const userLayout = opts.layoutId && (info.layouts || []).some((l) => l.id === opts.layoutId) ? opts.layoutId : null;
          if (userLayout || info.layoutId) addOpts.layoutId = userLayout || info.layoutId;
          slides.add(addOpts);
          slides.load("items/id");
          await context.sync();
          // opts.ignoreIds: 差し替え中の旧スライド(参照解析時には無かった ID)を既知扱いにし、新規スライドと取り違えない
          const known = new Set((info.allIds || []).concat(results.map((x) => x.slideId)).concat(opts.ignoreIds || []));
          const found = slides.items.find((s) => !known.has(s.id)) || slides.items[slides.items.length - 1];
          newId = found.id;
          newSlide = slides.getItem(newId); // ID で再取得
          newSlide.shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
          await context.sync();
          preDelete = newSlide.shapes.items.map((s) => `${s.name}@${Math.round(s.left)},${Math.round(s.top)} ${Math.round(s.width)}x${Math.round(s.height)}`);
          newSlide.shapes.items.forEach((s) => s.delete());
          if (info.has18 && afterId) {
            const ids = slides.items.map((s) => s.id);
            newSlide.moveTo(ids.indexOf(afterId) + 1);
          }
        }

        const images = drawPrims(newSlide.shapes, lay.prims, info);
        pres.setSelectedSlides([newId]);
        await context.sync();
        if (images.length) {
          await insertImages(images); // 選択中スライド(=新規)に画像を貼る
          pres.setSelectedSlides([newId]); // 画像選択状態を解除
          await context.sync();
        }
        const msDraw = performance.now() - t0;

        let png = null;
        if (opts.preview !== false && info.has18) {
          try {
            const img = newSlide.getImageAsBase64({ width: 640 });
            await context.sync();
            png = img.value;
          } catch (_) {}
        }
        const kinds = {};
        for (const p of lay.prims) kinds[p.kind] = (kinds[p.kind] || 0) + 1;
        return {
          slideId: newId,
          ms: Math.round(msDraw),
          png,
          warnings: lay.warnings,
          width: info.width,
          height: info.height,
          mode: useClone ? "clone" : "layout",
          debug: { version: root.SlideLayout.VERSION, prims: lay.prims.length, kinds, preDelete: preDelete ? preDelete.length : null, layoutId: useClone ? null : (opts.layoutId || info.layoutId || null) },
        };
      });
      results.push(r);
      afterId = r.slideId;
    }
    return results;
  }

  /** 互換: 1 枚 */
  async function renderSpec(spec, opts) {
    return (await renderSpecs([spec], opts))[0];
  }

  /**
   * 開発用: 現在のプレゼンの選択スライドを「既存デッキ」風(ネイビー Meiryo・リード・罫線・出典・ページ番号)に
   * 仕立てる。clone 検証のリファレンス作成に使う(同一文書内 Office.js なので確実)。
   */
  async function devBuildReference() {
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      slide.shapes.load("items/id");
      await context.sync();
      slide.shapes.items.forEach((s) => s.delete());
      const navy = "#1F3A5F",
        gray = "#595959",
        lineC = "#BFBFBF";
      const add = (x, y, w, h, text, size, bold, color, align) => {
        const sh = slide.shapes.addTextBox(text, { left: x, top: y, width: w, height: h });
        const tf = sh.textFrame;
        tf.leftMargin = 4;
        tf.rightMargin = 4;
        tf.topMargin = 2;
        tf.bottomMargin = 2;
        tf.verticalAlignment = "Middle";
        const tr = tf.textRange;
        tr.font.size = size;
        tr.font.bold = bold;
        tr.font.name = "Meiryo UI";
        tr.font.color = color;
        tr.paragraphFormat.horizontalAlignment = align;
      };
      const rule = (x, y, w, h, color) => {
        const s = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, { left: x, top: y, width: w, height: h });
        s.fill.setSolidColor(color);
        s.lineFormat.visible = false;
      };
      add(48, 28, 864, 44, "既存スライド: 主力製品の粗利率は 3 年で 6pt 低下", 24, true, navy, "Left");
      rule(48, 76, 864, 2, navy);
      add(48, 84, 864, 40, "原材料高と値引き競争が主因。価格改定と SKU 整理で来期 2pt の回復を狙う。", 16, false, gray, "Left");
      const box = (x) => rule(x, 140, 270, 300, "#F2F2F2");
      box(48);
      rule(345, 140, 567, 300, "#F2F2F2");
      add(60, 150, 246, 30, "本文ダミー: 粗利率推移", 14, false, gray, "Left");
      add(357, 150, 543, 30, "本文ダミー: 要因分解と打ち手", 14, false, gray, "Left");
      add(48, 470, 600, 20, "出典: 社内管理会計データ(2026年3月期)", 10, false, gray, "Left");
      rule(48, 500, 864, 1, lineC);
      add(48, 506, 400, 20, "© 2026 Sample Consulting Inc.", 9, false, gray, "Left");
      add(880, 506, 32, 20, "1", 9, false, gray, "Right");
      await context.sync();
      return { ok: true };
    });
  }

  root.SlideRender = { prepare, renderSpec, renderSpecs, deleteSlides, selectSlides, supports, analyzeReference, devBuildReference };
})(typeof self !== "undefined" ? self : this);
