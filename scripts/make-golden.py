# -*- coding: utf-8 -*-
"""
ゴールデンスライド v2 — ユーザーのマスタ(プロジェクトアプローチ.pptx)の上に python-pptx で直接組む品質の正解。
エンジン(layout.js)の制約は使わない。人が作るのと同じ自由度で、コンサル資料の典型パターンを網羅する。

  python scripts/make-golden.py [--master <path>] [--out docs/golden/golden-v2.pptx] [--only 3,7,11]

v2 での方針(v1 のレビュー反映):
- 本文域(y 128〜488)を使い切る。行高は「残り高さ ÷ 行数」で計算し、下を空けない。
- 文字サイズに階層を付ける: 見出し 13 太字 / 小見出し 11.5 太字 / 本文 10.5 / 補足 9 / 注記 8.5 / KPI 28。
- 矢羽(ステップ)の下には、段ごとの内容だけを行で揃えて置く。別の可視化を同居させない。
- 黒下線 + 中央揃え見出しのパネル(2〜4 枚)、n×m テキスト格子の複数パターン、縦/横の列挙を用意する。
"""
import argparse, os, sys
sys.stdout.reconfigure(encoding="utf-8")
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Pt, Emu
from lxml import etree

ap = argparse.ArgumentParser()
ap.add_argument("--master", default=r"C:\Users\nakaj\Documents\プロジェクトアプローチ.pptx")
ap.add_argument("--out", default="docs/golden/golden-v2.pptx")
ap.add_argument("--only", default="")
A = ap.parse_args()

# ---------------- トークン ----------------
ACCENT = RGBColor(0xD0, 0x4A, 0x02)
INK = RGBColor(0x1A, 0x1A, 0x1A)
DARK = RGBColor(0x40, 0x40, 0x40)
MUTED = RGBColor(0x6E, 0x6E, 0x6E)
MID = RGBColor(0xA6, 0xA6, 0xA6)
LIGHT = RGBColor(0xD9, 0xD9, 0xD9)
PALE = RGBColor(0xF2, 0xF2, 0xF2)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
AMBER = RGBColor(0xE0, 0x9A, 0x00)
RED = RGBColor(0xC6, 0x28, 0x28)
FONT = "Arial"
LAYOUT_NAME = "Title and Full Content - Subtitle"
BX, BY, BW = 32, 128, 896
BB = 488            # 本文域の下端
BH = BB - BY        # 360
SRC_Y = 494
# 文字階層
T_H1, T_H2, T_BODY, T_SMALL, T_NOTE, T_KPI = 13, 11.5, 10.5, 9, 8.5, 28

def pt(v):
    return Emu(int(round(v * 12700)))

# ---------------- 低レベル ----------------
def _font(run, size, bold=False, color=INK, name=FONT, italic=False):
    run.font.size = Pt(size); run.font.bold = bold; run.font.italic = italic
    run.font.name = name; run.font.color.rgb = color

def _bullet(p, level=0, char=None):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(160000 + level * 180000)); pPr.set("indent", str(-160000))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    bu = etree.SubElement(pPr, qn("a:buChar")); bu.set("char", char or ("•" if level == 0 else "–"))

def _nobullet(p):
    pPr = p._p.get_or_add_pPr(); pPr.set("marL", "0"); pPr.set("indent", "0")
    etree.SubElement(pPr, qn("a:buNone"))

def _rich(p, s, size, bold, color, name=FONT):
    for i, seg in enumerate(s.split("**")):
        if seg:
            r = p.add_run(); r.text = seg; _font(r, size, bold or (i % 2 == 1), color, name)

def text(s, x, y, w, h, content, size=T_BODY, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         margins=(3, 2, 3, 2), ls=1.15, name=FONT):
    """content: str | list[str | (str, opts)]。**…** は太字。opts: level/size/bold/color/align/sb(space_before)"""
    tb = s.shapes.add_textbox(pt(x), pt(y), pt(w), pt(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [pt(m) for m in margins]
    items = content if isinstance(content, list) else [content]
    for i, it in enumerate(items):
        o = {}
        if isinstance(it, tuple):
            it, o = it
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = o.get("align", align); p.line_spacing = o.get("ls", ls)
        if o.get("sb"):
            p.space_before = Pt(o["sb"])
        if o.get("level") is not None:
            _bullet(p, o["level"], o.get("char"))
        else:
            _nobullet(p)
        _rich(p, it, o.get("size", size), o.get("bold", bold), o.get("color", color), name)
    return tb

def rect(s, x, y, w, h, fill=None, line=None, lw=0.75, shape=MSO_SHAPE.RECTANGLE):
    sp = s.shapes.add_shape(shape, pt(x), pt(y), pt(w), pt(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(lw)
    sp.shadow.inherit = False
    if sp.has_text_frame:
        sp.text_frame.text = ""
    return sp

def shape_text(sp, content, size, bold=False, color=INK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margins=(4, 1, 4, 1)):
    tf = sp.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [pt(m) for m in margins]
    items = content if isinstance(content, list) else [content]
    for i, it in enumerate(items):
        o = {}
        if isinstance(it, tuple):
            it, o = it
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = o.get("align", align); p.line_spacing = 1.1
        _nobullet(p)
        _rich(p, it, o.get("size", size), o.get("bold", bold), o.get("color", color))
    return sp

def line(s, x1, y1, x2, y2, color=LIGHT, w=0.75, dash=False, arrow=False, elbow=False):
    c = s.shapes.add_connector(MSO_CONNECTOR.ELBOW if elbow else MSO_CONNECTOR.STRAIGHT, pt(x1), pt(y1), pt(x2), pt(y2))
    c.line.color.rgb = color; c.line.width = Pt(w)
    ln = c.line._get_or_add_ln()
    if dash:
        d = etree.SubElement(ln, qn("a:prstDash")); d.set("val", "dash")
    if arrow:
        t = etree.SubElement(ln, qn("a:tailEnd")); t.set("type", "triangle"); t.set("w", "med"); t.set("len", "med")
    return c

def band(s, x, y, w, h, label, fill=DARK, color=WHITE, size=T_H2, align=PP_ALIGN.CENTER, bold=True):
    return shape_text(rect(s, x, y, w, h, fill=fill), label, size, bold, color, align)

def panel_head(s, x, y, w, label, size=T_H1, lw=1.5):
    """黒下線 + 中央揃えの見出し。戻り値は下線の y"""
    text(s, x, y, w, 22, label, size=size, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM, margins=(0, 0, 0, 2))
    line(s, x, y + 24, x + w, y + 24, color=INK, w=lw)
    return y + 24

def rowhead(s, x, y, w, h, label, fill=LIGHT, size=T_BODY, color=INK, align=PP_ALIGN.LEFT):
    return shape_text(rect(s, x, y, w, h, fill=fill), label, size, True, color, align, margins=(6, 1, 6, 1))

def vchevron(s, x, y, w, h, label, fill=LIGHT, size=T_BODY):
    """縦の矢羽(マスタの VerticalHeader): 下向きの五角形 + 回転しない文字"""
    sp = rect(s, 0, 0, h, w, fill=fill, shape=MSO_SHAPE.PENTAGON)
    sp.rotation = 90
    sp.left, sp.top = pt(x + (w - h) / 2), pt(y + (h - w) / 2)
    text(s, x, y, w, h, label, size=size, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return sp

def harvey(s, cx, cy, d, level):
    x, y = cx - d / 2, cy - d / 2
    if level >= 4:
        rect(s, x, y, d, d, fill=DARK, line=DARK, shape=MSO_SHAPE.OVAL); return
    rect(s, x, y, d, d, fill=WHITE, line=DARK, lw=1, shape=MSO_SHAPE.OVAL)
    if level <= 0:
        return
    sp = s.shapes.add_shape(MSO_SHAPE.PIE, pt(x), pt(y), pt(d), pt(d))
    sp.fill.solid(); sp.fill.fore_color.rgb = DARK; sp.line.fill.background(); sp.shadow.inherit = False
    sp.adjustments[0] = 270 * 0.6; sp.adjustments[1] = ((270 + 90 * level) % 360) * 0.6

def dot(s, cx, cy, d, color):
    return rect(s, cx - d / 2, cy - d / 2, d, d, fill=color, shape=MSO_SHAPE.OVAL)

def source(s, t):
    text(s, BX, SRC_Y, BW, 12, t, size=8, color=MUTED, margins=(0, 0, 0, 0))

def rows_fill(y0, y1, n, gap=4):
    h = (y1 - y0 - gap * (n - 1)) / n
    return [(y0 + i * (h + gap), h) for i in range(n)]

def cols_fill(x0, w, n, gap=12):
    cw = (w - gap * (n - 1)) / n
    return [(x0 + i * (cw + gap), cw) for i in range(n)]

def style_chart(chart, size=9, legend=None, gridlines=False):
    chart.font.size = Pt(size); chart.font.name = FONT; chart.has_title = False
    if legend:
        chart.has_legend = True; chart.legend.position = legend; chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(size)
    else:
        chart.has_legend = False
    try:
        va = chart.value_axis
        va.has_major_gridlines = gridlines
        if gridlines:
            va.major_gridlines.format.line.color.rgb = PALE
        va.format.line.fill.background(); va.tick_labels.font.size = Pt(size); va.tick_labels.font.color.rgb = MUTED
        ca = chart.category_axis
        ca.format.line.color.rgb = LIGHT; ca.tick_labels.font.size = Pt(size); ca.tick_labels.font.color.rgb = INK
        ca.has_major_gridlines = False
    except Exception:
        pass

def line_series(ser, col, width=2.25, labels=True, pos=XL_LABEL_POSITION.ABOVE, size=9, dash=False):
    ser.format.line.color.rgb = col; ser.format.line.width = Pt(width); ser.smooth = False
    if dash:
        ser.format.line.dash_style = 4
    ser.marker.style = 8; ser.marker.size = 6
    ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = col; ser.marker.format.line.color.rgb = col
    if labels:
        dl = ser.data_labels; dl.show_value = True; dl.number_format = "0.0"; dl.number_format_is_linked = False
        dl.font.size = Pt(size); dl.font.color.rgb = col; dl.position = pos

# ---------------- 骨組み ----------------
prs = Presentation(A.master)
layout = next(l for l in prs.slide_layouts if l.name == LAYOUT_NAME)
for sldId in list(prs.slides._sldIdLst):
    prs.part.drop_rel(sldId.rId); prs.slides._sldIdLst.remove(sldId)
page = [0]
SLIDES = []

def slide(fn):
    SLIDES.append(fn); return fn

def new_slide(title, lead, title_size=20, lead_size=13):
    page[0] += 1
    s = prs.slides.add_slide(layout)
    for ph in list(s.placeholders):
        if ph.placeholder_format.idx == 1:
            ph._element.getparent().remove(ph._element)
    t = s.shapes.title
    t.left, t.top, t.width, t.height = pt(32), pt(30), pt(896), pt(58)
    t.text_frame.text = title
    for p in t.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(title_size)
    sub = next(ph for ph in s.placeholders if ph.placeholder_format.idx == 16)
    sub.left, sub.top, sub.width, sub.height = pt(32), pt(92), pt(896), pt(32)
    sub.text_frame.text = lead; sub.text_frame.word_wrap = True
    for p in sub.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(lead_size); r.font.color.rgb = DARK; r.font.name = FONT
    text(s, 919, 513, 9, 9, str(page[0]), size=7.5, margins=(0, 0, 0, 0), align=PP_ALIGN.RIGHT)
    return s

def bullets(items, size=T_BODY):
    """[(text, level)] or [text] → text() 用のリスト。"#…" は太字の小見出し(段落前に余白)"""
    out = []
    for it in items:
        if isinstance(it, tuple):
            t, lv = it
            out.append((t, {"level": lv, "size": size - (0.5 if lv else 0)}))
        elif it.startswith("#"):
            out.append((it[1:], {"bold": True, "size": T_H2, "sb": 6}))
        else:
            out.append((it, {"level": 0, "size": size}))
    return out

# =====================================================================
@slide
def s01_exec_summary():
    s = new_slide("収益悪化の主因は値引きと物流費であり、価格統制と拠点集約で営業利益率を 2.1% から 6.5% に戻せる",
                  "粗利率は 3 年で 6.1pt 低下したが、市場要因(原材料高)は 1.8pt にとどまる。残り 4.3pt は自社の運営に起因し、投資 4.8 億円・回収 14 か月で是正できる")
    cols = [("#", 40), ("主張", 250), ("根拠となる事実", 396), ("利益インパクト(FY27)", 210)]
    x = BX
    for lab, w in cols:
        band(s, x, BY, w - 4, 22, lab, align=PP_ALIGN.CENTER if lab == "#" else PP_ALIGN.LEFT); x += w
    rows = [
        ("1", "値引きを本部承認制にし、平均値引き率を 9.7% → 5% 以下へ",
         ["平均値引き率は 4.8%(FY23)→ 9.7%(FY25)に拡大。粗利率低下 6.1pt のうち **2.7pt** を占める",
          "値引き総額 15.6 億円の 73% が上位 20 顧客に集中。承認制の対象を絞れば統制できる",
          "同業 B 社は承認制導入後 8 か月で値引き率を 4.1pt 圧縮(公開資料)"], "+2.4pt", "+11.5 億円"),
        ("2", "6 拠点を 3 拠点に集約し、共同配送で物流費率 9.4% → 8.3% へ",
         ["拠点稼働率は平均 58%、小口配送(1 便 20 万円未満)が件数の 41%",
          "翌日配送は上位 50 社 + 半径 150km で維持でき、顧客離反リスクは限定的(31 社ヒアリング)",
          "投資 3.1 億円、年間削減 5.3 億円、回収 7 か月"], "+1.1pt", "+5.3 億円"),
        ("3", "下位 40% の 1,540 SKU を 2 段階で廃番し、段取り時間を 33% 削減",
         ["3,860 SKU のうち下位 40% は売上の 3.2%、粗利率 15% 未満",
          "段取り替え月 1,900 回・1 回 42 分。SKU 集約で月 1,270 回・28 分へ",
          "廃番基準(年間売上 300 万円未満かつ粗利率 15% 未満)を新設"], "+0.8pt", "+3.8 億円"),
    ]
    tot_h = 40
    for (y, rh), (num, claim, facts, pt_, yen) in zip(rows_fill(BY + 28, BB - tot_h - 8, 3, 6), rows):
        x = BX
        text(s, x, y, 36, rh, num, size=22, bold=True, color=ACCENT, align=PP_ALIGN.CENTER); x += 40
        text(s, x, y, 246, rh, claim, size=T_H2, bold=True); x += 250
        text(s, x, y, 392, rh, bullets(facts), size=T_BODY); x += 396
        text(s, x, y, 206, rh, [(pt_, {"size": 18, "bold": True, "align": PP_ALIGN.CENTER}), (yen, {"size": T_BODY, "color": DARK, "align": PP_ALIGN.CENTER})], anchor=MSO_ANCHOR.MIDDLE)
        line(s, BX, y + rh + 3, BX + BW, y + rh + 3, color=LIGHT)
    y = BB - tot_h
    rect(s, BX, y, BW, tot_h, fill=PALE)
    text(s, BX + 8, y, 640, tot_h, "3 施策の合計: 営業利益率 **2.1% → 6.5%**(FY27 通期)。投資 4.8 億円、回収 14 か月。市場要因(原材料高 1.8pt)は医療機器向け参入で 3 年かけて相殺する", size=T_BODY, anchor=MSO_ANCHOR.MIDDLE)
    text(s, BX + BW - 220, y, 212, tot_h, [("+4.4pt", {"size": 18, "bold": True, "color": ACCENT, "align": PP_ALIGN.RIGHT}), ("+20.6 億円", {"size": T_BODY, "color": DARK, "align": PP_ALIGN.RIGHT})], anchor=MSO_ANCHOR.MIDDLE)
    source(s, "出典: X 社 財務データ(FY2023〜FY2025)、営業日報 12,400 件、拠点別原価集計、顧客ヒアリング 31 社。インパクトは FY2027 通期の営業利益への寄与")

# =====================================================================
@slide
def s02_two_panel():
    s = new_slide("現状の課題は「値引きの統制不在」と「物流の固定費過多」に集約でき、打ち手は仕組みの変更と拠点の再編である",
                  "課題は営業日報 12,400 件と拠点別原価の分析から特定した。打ち手は他社事例と自社の顧客構成に照らして 3 か月以内に着手できるものに絞っている")
    (x1, w), (x2, _) = cols_fill(BX, BW, 2, 24)
    y0 = panel_head(s, x1, BY, w, "現状の課題")
    panel_head(s, x2, BY, w, "打ち手の方向性")
    foot_h = 92
    text(s, x1, y0 + 8, w, BB - foot_h - y0 - 16, bullets([
        "#値引きが担当者の裁量で決まり、上限も承認もない",
        "平均値引き率 4.8%(FY23)→ 9.7%(FY25)。上位 20 顧客に総額の 73% が集中",
        ("下位顧客にも 10% 超の値引きが 18% ある(ランクと連動していない)", 1),
        ("案件粗利が申請時に見えず、担当者が影響を把握できない", 1),
        "#物流拠点 6 か所の稼働率が 58% で、固定費が重い",
        "小口配送(1 便 20 万円未満)が件数の 41%、便あたり原価は大口の 2.3 倍",
        ("拠点間の在庫移動が月 340 便。集約すれば 6 割が不要", 1),
        "#SKU が 3,860 点に増え、段取り替えが月 1,900 回",
        "下位 40% の SKU は売上の 3.2% しかなく、粗利率 15% 未満",
        ("段取り替え 1 回 42 分、月 1,330 時間が段取りに消える", 1),
    ], size=12.5), size=12.5)
    text(s, x2, y0 + 8, w, BB - foot_h - y0 - 16, bullets([
        "#値引きを本部承認制にし、ランク別の上限で自動判定する",
        "5% 以下は即時承認、超過は本部が 1 営業日で判断。申請時に案件粗利を表示",
        ("顧客ランク A/B/C の上限を 8% / 5% / 3% に設定", 1),
        ("FY27 から粗利額と値引き率を営業の評価指標に加える", 1),
        "#拠点を 3 か所に集約し、共同配送で小口を束ねる",
        "翌日配送は上位 50 社 + 半径 150km 圏で維持(31 社ヒアリングで離反懸念は 2 社)",
        ("投資 3.1 億円、年間削減 5.3 億円、回収 7 か月", 1),
        "#下位 SKU を 2 段階で廃番し、段取り時間を 33% 短縮",
        "廃番基準(年間売上 300 万円未満かつ粗利率 15% 未満)で 1,540 点を対象化",
        ("代替品の対応表を付けて主要顧客 12 社へ事前照会し、失う売上を 0.6 億円以下に抑える", 1),
    ], size=12.5), size=12.5)
    y = BB - foot_h
    band(s, BX, y, BW, 22, "課題と打ち手の対応", align=PP_ALIGN.LEFT)
    text(s, BX, y + 26, BW, foot_h - 26, bullets([
        "**3 つの課題はいずれも「基準がなく担当者任せ」に帰着する。** 打ち手は人を増やすのではなく、基準(上限・廃番・配送ルール)を明文化して仕組みに載せる",
        "**効果の合計は営業利益率 +4.3pt(FY27)。** 投資 4.8 億円のうち 3.1 億円は拠点集約で、6 月の投資判断を経て着手する",
        "**打ち手の順序は値引き → SKU → 拠点。** 前の 2 つは投資が小さく 3 か月で効果が出るため、拠点集約の投資判断の材料(粗利の回復実績)になる",
    ], size=11), size=11)
    source(s, "出典: 営業日報 12,400 件(FY2025)、拠点別原価集計、顧客ヒアリング 31 社(2026 年 5〜6 月)")

# =====================================================================
@slide
def s03_trend():
    s = new_slide("粗利率は 3 年連続で低下し、FY25 に業界平均を 4.6pt 下回った。原材料高だけでは説明できない",
                  "低下幅 6.1pt のうち原材料高は 1.8pt。残る 4.3pt は値引きと物流費という自社の運営要因で、同業他社は同じ原材料高の下で 30% 前後を維持している")
    cw = 540
    band(s, BX, BY, cw, 22, "粗利率の推移(%)", align=PP_ALIGN.LEFT)
    cd = CategoryChartData(); cd.categories = ["FY21", "FY22", "FY23", "FY24", "FY25"]
    cd.add_series("X 社", (32.4, 31.8, 31.2, 28.0, 25.1)); cd.add_series("業界平均(8 社)", (30.1, 30.4, 30.2, 29.9, 29.7))
    ch = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, pt(BX), pt(BY + 26), pt(cw), pt(BB - BY - 26), cd).chart
    style_chart(ch, legend=XL_LEGEND_POSITION.BOTTOM, gridlines=True)
    ch.value_axis.minimum_scale = 20; ch.value_axis.maximum_scale = 35; ch.value_axis.major_unit = 5
    line_series(ch.series[0], ACCENT); line_series(ch.series[1], MID, 1.5, pos=XL_LABEL_POSITION.BELOW)
    for k in (3, 4):
        ch.series[0].points[k].data_label.position = XL_LABEL_POSITION.BELOW
        ch.series[0].points[k].data_label.font.size = Pt(9); ch.series[0].points[k].data_label.font.color.rgb = ACCENT
        ch.series[1].points[k].data_label.position = XL_LABEL_POSITION.ABOVE
        ch.series[1].points[k].data_label.font.size = Pt(9); ch.series[1].points[k].data_label.font.color.rgb = MID
    text(s, BX + 250, BY + 36, 230, 40, "FY24: 主要顧客 A 社との年次交渉で **7% の値引き**に応じ、単年で −3.2pt", size=T_SMALL, color=DARK)
    line(s, BX + 360, BY + 76, BX + 392, BY + 130, color=DARK)
    rx = BX + cw + 24; rw = BX + BW - rx
    band(s, rx, BY, rw, 22, "意味合い", align=PP_ALIGN.LEFT)
    text(s, rx, BY + 28, rw, 150, bullets([
        "#低下の 7 割は自社要因",
        "値引き率の拡大 2.7pt、物流費の増加 1.5pt、SKU 増による段取り 0.6pt",
        "原材料高(特殊鋼 +12%)は 1.8pt で、同業 8 社も同じ条件",
        "#業界平均との差 4.6pt は「取り戻せる差」",
        "同業 B 社は FY24 に承認制を導入し、値引き率を 4.1pt 圧縮した",
        "差 4.6pt のうち 4.3pt は 3 施策で FY27 までに回収できる",
    ]))
    ty = BY + 196
    text(s, rx, ty, rw, 16, "FY25 の同業比較", size=T_H2, bold=True, margins=(0, 0, 0, 0))
    line(s, rx, ty + 18, rx + rw, ty + 18, color=INK, w=0.75)
    rows = [("", "X 社", "同業平均", "差"), ("粗利率", "25.1%", "29.7%", "−4.6pt"), ("平均値引き率", "9.7%", "5.2%", "+4.5pt"), ("物流費率", "9.4%", "8.1%", "+1.3pt"), ("原材料費率", "41.8%", "41.2%", "+0.6pt")]
    cw_ = [rw - 210, 70, 70, 70]; yy = ty + 20
    for i, r in enumerate(rows):
        xx = rx
        for j, v in enumerate(r):
            text(s, xx, yy, cw_[j], 17, v, size=T_SMALL, bold=(i == 0 or j == 3), color=DARK if i == 0 else INK, align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT, margins=(0, 0, 0, 0), anchor=MSO_ANCHOR.MIDDLE)
            xx += cw_[j]
        line(s, rx, yy + 18, rx + rw, yy + 18, color=PALE); yy += 19
    rect(s, rx, BB - 44, rw, 44, fill=PALE)
    text(s, rx + 6, BB - 44, rw - 12, 44, "**So what:** 原材料高への対応(価格転嫁)よりも先に、値引きと物流の統制を打つべき", size=T_BODY, anchor=MSO_ANCHOR.MIDDLE)
    source(s, "出典: X 社 決算資料(FY2021〜FY2025)、業界平均は同業 8 社の公表値の単純平均。原材料高の影響は原価差異分析による")

# =====================================================================
@slide
def s04_waterfall():
    s = new_slide("粗利率 6.1pt の低下は、値引き 2.7pt・原材料 1.8pt・物流 1.5pt・SKU 0.6pt に分解でき、7 割が運営要因",
                  "FY23 → FY25 の原価差異を要因別に積み上げた。為替と数量は +0.5pt の押し上げで、自社要因が無ければ低下は 1.3pt にとどまっていた")
    cw = 560
    band(s, BX, BY, cw, 22, "粗利率の変化要因(FY23 → FY25、pt)", align=PP_ALIGN.LEFT)
    cats = ["FY23", "値引き拡大", "原材料高", "物流費増", "SKU 段取り", "為替・数量", "FY25"]
    vals = [31.2, -2.7, -1.8, -1.5, -0.6, 0.5, 25.1]
    base, up, down, total = [], [], [], []; run = 0.0
    for i, v in enumerate(vals):
        if i in (0, len(vals) - 1):
            base.append(0); total.append(v); up.append(0); down.append(0); run = v
        elif v < 0:
            base.append(run + v); down.append(-v); up.append(0); total.append(0); run += v
        else:
            base.append(run); up.append(v); down.append(0); total.append(0); run += v
    cd = CategoryChartData(); cd.categories = cats
    for n_, d_ in (("base", base), ("total", total), ("減少", down), ("増加", up)):
        cd.add_series(n_, d_)
    chh = BB - BY - 26
    ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, pt(BX), pt(BY + 26), pt(cw), pt(chh), cd).chart
    style_chart(ch)
    ch.value_axis.minimum_scale = 20; ch.value_axis.maximum_scale = 34; ch.value_axis.visible = False
    ch.plots[0].gap_width = 45; ch.plots[0].overlap = 100
    for ser, col in zip(ch.series, [None, DARK, MID, LIGHT]):
        if col is None:
            ser.format.fill.background()
        else:
            ser.format.fill.solid(); ser.format.fill.fore_color.rgb = col
        ser.format.line.fill.background()
    ch.series[2].points[1].format.fill.solid(); ch.series[2].points[1].format.fill.fore_color.rgb = ACCENT
    slot = (cw - 50) / len(cats)
    for i, v in enumerate(vals):
        lab = f"{v:.1f}" if i in (0, len(vals) - 1) else f"{v:+.1f}"
        top = max(base[i] + up[i] + down[i], total[i])
        yy = BY + 26 + chh - 28 - (top - 20) / 14 * (chh - 40)
        text(s, BX + 30 + slot * i, yy - 15, slot, 14, lab, size=T_SMALL, bold=True, align=PP_ALIGN.CENTER, margins=(0, 0, 0, 0))
    rx = BX + cw + 24; rw = BX + BW - rx
    band(s, rx, BY, rw, 22, "読み取り", align=PP_ALIGN.LEFT)
    text(s, rx, BY + 28, rw, 170, bullets([
        "#運営要因 4.8pt / 市場要因 1.8pt",
        "値引き・物流・SKU は自社の意思決定で戻せる",
        "原材料高は同業共通。価格転嫁は交渉力の回復後",
        "#値引き 2.7pt の内訳",
        "上位 20 顧客への個別値引き 1.9pt、下位顧客への一律値引き 0.8pt",
        "#物流費 1.5pt の内訳",
        "小口配送の増加 0.9pt、拠点固定費 0.6pt",
    ]))
    ty = BY + 218
    text(s, rx, ty, rw, 16, "施策による回収見込み(FY27)", size=T_H2, bold=True, margins=(0, 0, 0, 0))
    line(s, rx, ty + 18, rx + rw, ty + 18, color=INK, w=0.75)
    rows = [("要因", "低下", "回収", "施策"), ("値引き", "2.7", "2.4", "承認制"), ("物流", "1.5", "1.1", "拠点集約"), ("SKU 段取り", "0.6", "0.8", "廃番 + 段取り短縮"), ("原材料", "1.8", "0.9*", "価格転嫁(未検証)")]
    cw_ = [74, 40, 40, rw - 154]; yy = ty + 20
    for i, r in enumerate(rows):
        xx = rx
        for j, v in enumerate(r):
            text(s, xx, yy, cw_[j], 17, v, size=T_SMALL, bold=(i == 0), color=DARK if i == 0 else INK, align=PP_ALIGN.RIGHT if j in (1, 2) else PP_ALIGN.LEFT, margins=(6 if j == 3 else 0, 0, 4 if j in (1, 2) else 0, 0), anchor=MSO_ANCHOR.MIDDLE)
            xx += cw_[j]
        line(s, rx, yy + 18, rx + rw, yy + 18, color=PALE); yy += 19
    text(s, rx, yy + 2, rw, 12, "* 転嫁できた場合。9 月の顧客交渉で確認", size=8, color=MUTED, margins=(0, 0, 0, 0))
    source(s, "出典: X 社 原価差異分析(FY2023 → FY2025)。要因の帰属は標準原価との差異を、価格差異・数量差異・費目別に配分して算出")

# =====================================================================
@slide
def s05_three_panel():
    s = new_slide("3 つの論点それぞれに仮説を置き、10 月までの 6 週間で検証する。データは既存の日報と原価集計で足りる",
                  "論点は「なぜ利益が減ったか」の分解から導き、仮説は同業事例と一次分析で立てた。検証は追加のシステム投資なしで進める")
    cols = cols_fill(BX, BW, 3, 20)
    heads = ["論点", "仮説", "検証アプローチ"]
    y0 = None
    for (x, w), h in zip(cols, heads):
        y0 = panel_head(s, x, BY, w, h)
    items = [
        ("値引きは統制できるか", "上位 20 顧客の個別値引きを承認制にすれば、値引き率を 4pt 圧縮できる",
         ["顧客別の値引き率と粗利を FY23〜25 で分解", "上位 20 社の契約条件と交渉履歴を確認", "B 社の導入前後の値引き率を公開資料で比較"]),
        ("拠点集約で納期は守れるか", "3 拠点でも上位 50 社 + 150km 圏は翌日配送を維持でき、離反は限定的",
         ["配送先の座標と納期要件を 12 か月分で地図化", "3 拠点案のシミュレーション(物流会社 2 社)", "上位 50 社ヒアリングで許容納期を確認"]),
        ("SKU 廃番で売上を失わないか", "下位 40% の SKU は代替品で 8 割を吸収でき、失う売上は 0.6 億円以下",
         ["SKU 別の売上・粗利・顧客数を集計", "代替品の対応表を商品企画が作成", "主要顧客 12 社に廃番候補を事前照会"]),
    ]
    owners = [("営業本部長 + コンサル 2 名", "10/3", "承認制の上限値(8/5/3%)を確定"), ("SCM 部長 + 物流設計 2 名", "9/26", "3 拠点案の投資判断資料に反映"), ("商品企画部長 + コンサル 1 名", "9/19", "第 2 弾 640 点の対象を確定")]
    foot_h = 70
    rows = rows_fill(y0 + 10, BB - foot_h - 8, 3, 8)
    for k, ((y, rh), (q, hyp, steps), (own, due, dec)) in enumerate(zip(rows, items, owners)):
        x, w = cols[0]
        shape_text(rect(s, x, y, 32, 32, fill=DARK), f"{k + 1}", 14, True, WHITE)
        text(s, x + 40, y, w - 40, rh, [(q, {"size": 13, "bold": True}), (f"担当: {own}", {"size": T_SMALL + 0.5, "color": DARK, "sb": 6}), (f"結論の期限: {due}", {"size": T_SMALL + 0.5, "color": DARK})], anchor=MSO_ANCHOR.TOP)
        x, w = cols[1]
        text(s, x, y, w, rh, [(hyp, {"size": 12}), (f"検証後の使い方: {dec}", {"size": T_SMALL + 0.5, "color": DARK, "sb": 6})], anchor=MSO_ANCHOR.TOP)
        x, w = cols[2]
        text(s, x, y, w, rh, bullets(steps, size=11.5), size=11.5, anchor=MSO_ANCHOR.TOP)
        if k < 2:
            line(s, BX, y + rh + 4, BX + BW, y + rh + 4, color=LIGHT)
    y = BB - foot_h
    band(s, BX, y, BW, 22, "検証の進め方", align=PP_ALIGN.LEFT)
    text(s, BX, y + 26, BW, foot_h - 26, bullets([
        "**3 つの検証は並行で進め、週次の PMO で中間結果を共有する。** 仮説が棄却された場合は代替案(値引きは段階的上限、拠点は 4 拠点案、SKU は第 1 弾のみ)に切り替える",
        "**追加データの取得は不要。** 営業日報・原価集計・配送実績はすべて既存システムから抽出でき、顧客ヒアリングは 9 月中旬までに終える",
    ], size=T_BODY), size=T_BODY)
    source(s, "出典: プロジェクト計画書(2026 年 8 月版)。検証期間は 9 月第 1 週〜10 月第 2 週")

# =====================================================================
@slide
def s06_options():
    s = new_slide("物流拠点は 3 拠点集約(案 B)が投資回収と顧客影響の両面で最も優れ、年 5.3 億円を 7 か月で回収する",
                  "案 A は効果が 1.2 億円と小さく、案 C は削減 6.8 億円と最大だが上位 20 社しか翌日配送を維持できず、推定 9 億円/年の受注を失う")
    hw = 190; cw3 = (BW - hw) / 3
    for j, (lab, f, c) in enumerate([("案 A: 現状維持 + 効率化", LIGHT, INK), ("案 B: 3 拠点に集約(推奨)", ACCENT, WHITE), ("案 C: 2 拠点に集約", LIGHT, INK)]):
        band(s, BX + hw + cw3 * j, BY, cw3 - 4, 24, lab, fill=f, color=c)
    crit = [("年間削減額", ["1.2 億円", "**5.3 億円**", "6.8 億円"], None), ("初期投資", ["0.4 億円", "3.1 億円", "5.6 億円"], None),
            ("回収期間", ["4 か月", "**7 か月**", "10 か月"], None), ("翌日配送の維持", ["全顧客", "上位 50 社 + 150km 圏", "上位 20 社のみ"], [4, 3, 1]),
            ("実施の容易さ", ["3 か月・投資小", "9 か月・並行稼働", "14 か月・大規模移転"], [4, 3, 1]),
            ("在庫・品質リスク", ["変化なし", "移行期に +0.8 億円", "移行期に +1.9 億円"], [4, 3, 2]),
            ("顧客離反リスク", ["なし", "限定的(31 社中 2 社が懸念)", "推定 −9 億円/年"], [4, 3, 1])]
    tot_h = 34; leg_h = 14
    rows = rows_fill(BY + 28, BB - tot_h - leg_h - 6, len(crit), 4)
    for (y, rh), (lab, cells, balls) in zip(rows, crit):
        rowhead(s, BX, y, hw - 4, rh, lab, size=T_BODY)
        for j, c in enumerate(cells):
            cx0 = BX + hw + cw3 * j
            if j == 1:
                rect(s, cx0, y, cw3 - 4, rh, fill=PALE)
            if balls:
                harvey(s, cx0 + 16, y + rh / 2, 14, balls[j])
                text(s, cx0 + 28, y, cw3 - 34, rh, c, size=T_BODY, anchor=MSO_ANCHOR.MIDDLE)
            else:
                text(s, cx0 + 6, y, cw3 - 14, rh, c, size=T_H2, anchor=MSO_ANCHOR.MIDDLE)
        line(s, BX + hw, y + rh + 2, BX + BW, y + rh + 2, color=LIGHT)
    y = BB - tot_h - leg_h - 2
    rowhead(s, BX, y, hw - 4, tot_h, "総合評価", fill=DARK, color=WHITE, size=T_BODY)
    for j, (c, col) in enumerate([("効果不足", INK), ("推奨: 効果・回収・顧客影響のバランス", WHITE), ("顧客影響が大きい", INK)]):
        cx0 = BX + hw + cw3 * j
        shape_text(rect(s, cx0, y, cw3 - 4, tot_h, fill=ACCENT if j == 1 else PALE), c, T_BODY, True, col, PP_ALIGN.LEFT, margins=(6, 1, 6, 1))
    text(s, BX + 500, BB - leg_h, 396, leg_h, "● 良い   ◕ やや良い   ◑ 中   ◔ やや悪い   ○ 悪い", size=8, color=MUTED, align=PP_ALIGN.RIGHT, margins=(0, 0, 0, 0))
    source(s, "出典: 拠点別原価集計(FY2025)、顧客ヒアリング 31 社(2026 年 5〜6 月)、物流会社 3 社見積。顧客離反は上位 20 社以外の翌日配送依存案件を集計")

# =====================================================================
@slide
def s07_2x2():
    s = new_slide("効果が大きく実行しやすい値引き統制と SKU 整理を先行し、拠点集約は投資判断を Q3 に置いて並行で準備する",
                  "12 施策を利益効果と実行難易度で整理した。左上 3 施策で効果の 76%(15.6 億円)を占め、6 か月以内に着手できる")
    px, py, pw, ph = BX + 36, BY + 6, 560, BB - BY - 26
    rect(s, px, py, pw, ph, line=LIGHT)
    line(s, px + pw / 2, py, px + pw / 2, py + ph, color=LIGHT, dash=True)
    line(s, px, py + ph / 2, px + pw, py + ph / 2, color=LIGHT, dash=True)
    text(s, px, py + ph + 4, pw, 14, "実行難易度  →  高(投資額・関係部門数・期間で評価)", size=T_SMALL, color=MUTED, align=PP_ALIGN.CENTER, margins=(0, 0, 0, 0))
    ax = text(s, px - 30 - ph / 2 + 8, py + ph / 2 - 8, ph, 16, "利益効果  →  大(FY27 営業利益への寄与、pt)", size=T_SMALL, color=MUTED, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0))
    ax.rotation = 270
    for lab, xx, yy in [("先行して実行", px + 6, py + 4), ("投資判断のうえ実行", px + pw / 2 + 6, py + 4), ("手が空けば実行", px + 6, py + ph - 18), ("見送り", px + pw / 2 + 6, py + ph - 18)]:
        text(s, xx, yy, 200, 14, lab, size=T_SMALL, bold=True, color=MUTED, margins=(0, 0, 0, 0))
    pts = [("値引き承認制", 0.12, 0.90, 2.4), ("SKU 廃番", 0.25, 0.72, 0.8), ("小口配送の統合", 0.30, 0.55, 0.4), ("3 拠点集約", 0.70, 0.80, 1.1), ("医療機器向け参入", 0.88, 0.85, 3.0),
           ("営業日報の標準化", 0.10, 0.22, 0.1), ("見積テンプレート統一", 0.18, 0.30, 0.15), ("在庫基準の見直し", 0.35, 0.18, 0.2),
           ("基幹システム刷新", 0.82, 0.30, 0.3), ("工場自動化(第 2 期)", 0.75, 0.20, 0.4), ("海外拠点の再編", 0.90, 0.12, 0.2), ("人事制度改定", 0.62, 0.10, 0.1)]
    for lab, fx, fy, imp in pts:
        d = 10 + imp * 9; cx, cy = px + fx * pw, py + (1 - fy) * ph
        rect(s, cx - d / 2, cy - d / 2, d, d, fill=ACCENT if fy > 0.5 and fx < 0.5 else (DARK if fy > 0.5 else MID), shape=MSO_SHAPE.OVAL)
        lab_s = f"{lab} ({imp:+.1f}pt)" if imp >= 0.3 else lab
        if fx > 0.8:
            text(s, cx - d / 2 - 132, cy - 8, 130, 16, lab_s, size=T_SMALL, margins=(0, 0, 0, 0), anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)
        else:
            text(s, cx + d / 2 + 2, cy - 8, 130, 16, lab_s, size=T_SMALL, margins=(0, 0, 0, 0), anchor=MSO_ANCHOR.MIDDLE)
    rx = px + pw + 24; rw = BX + BW - rx
    band(s, rx, BY, rw, 22, "判断", align=PP_ALIGN.LEFT)
    text(s, rx, BY + 28, rw, 200, bullets([
        "#先行 3 施策(6 か月以内)", "値引き承認制 +2.4pt", "SKU 廃番 +0.8pt", "小口配送の統合 +0.4pt",
        "#投資判断(Q3)", "3 拠点集約 +1.1pt、投資 3.1 億円", "医療機器向け参入 +3pt/3 年、投資 1.2 億円",
        "#見送り・後回し", "基幹刷新・自動化は拠点集約の完了後に再評価",
    ]))
    ty = BY + 236
    text(s, rx, ty, rw, 16, "象限ごとの集計", size=T_H2, bold=True, margins=(0, 0, 0, 0))
    line(s, rx, ty + 18, rx + rw, ty + 18, color=INK, w=0.75)
    rows = [("象限", "施策数", "効果 pt", "投資 億円"), ("先行", "3", "3.6", "0.4"), ("投資判断", "2", "4.1", "4.3"), ("手が空けば", "3", "0.5", "0.3"), ("見送り", "4", "1.0", "6.8")]
    cw_ = [rw - 156, 46, 55, 55]; yy = ty + 20
    for i, r in enumerate(rows):
        xx = rx
        for j, v in enumerate(r):
            text(s, xx, yy, cw_[j], 17, v, size=T_SMALL, bold=(i == 0), color=DARK if i == 0 else INK, align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT, margins=(0, 0, 0, 0), anchor=MSO_ANCHOR.MIDDLE)
            xx += cw_[j]
        line(s, rx, yy + 18, rx + rw, yy + 18, color=PALE); yy += 19
    text(s, rx, yy + 4, rw, 12, "円の大きさ = 利益効果(pt)", size=8, color=MUTED, margins=(0, 0, 0, 0))
    source(s, "出典: 施策別の効果試算(FY2027 営業利益ベース)。実行難易度は投資額・関係部門数・実施期間を 3 段階で評価し合成")

# =====================================================================
@slide
def s08_matrix_sentences():
    s = new_slide("部門ごとに課題の現れ方は異なるが、原因は「基準がなく担当者任せ」で共通しており、打ち手は基準の明文化に集約される",
                  "営業・商品企画・SCM・生産の 4 部門について、課題 → 原因 → 打ち手 → 効果を一つの格子で整理した。効果は FY27 の営業利益への寄与")
    hw = 96; cols = [("課題(現状)", 230), ("原因", 210), ("打ち手", 230), ("効果", BW - hw - 230 - 210 - 230)]
    x = BX + hw
    for lab, w in cols:
        band(s, x, BY, w - 4, 22, lab, align=PP_ALIGN.LEFT); x += w
    rows = [
        ("営業", "値引きが担当者の裁量で決まり、平均 9.7%。上位 20 顧客に総額の 73% が集中", "値引き上限も承認基準もない。案件粗利が申請時に見えず、売上高だけで評価される",
         "本部承認制と顧客ランク別の上限(8/5/3%)。申請時に案件粗利を表示し、FY27 から粗利を評価に加える", "**+2.4pt**\n値引き率 5.0% へ"),
        ("商品企画", "SKU が 3,860 点に増え、下位 40% は売上の 3.2%。段取り替えが月 1,900 回", "廃番基準がなく、顧客要望で追加した専用品が残り続ける",
         "廃番基準(年間売上 300 万円未満かつ粗利率 15% 未満)を新設し、2 段階で 1,540 点を廃番。代替品の対応表を整備", "**+0.8pt**\nSKU 2,320 へ"),
        ("SCM", "6 拠点の稼働率 58%。小口配送が件数の 41% で、便あたり原価は大口の 2.3 倍", "拠点は買収時のまま。配送は顧客ごとに個別対応し、便の統合ルールがない",
         "3 拠点に集約し、共同配送で小口を束ねる。翌日配送は上位 50 社 + 150km 圏で維持", "**+1.1pt**\n物流費率 8.3% へ"),
        ("生産", "稼働率 71%(−4pt)。段取り替え 1 回 42 分で、月 1,330 時間が段取りに消える", "SKU 増で小ロット化。段取りの標準手順がなく、熟練者に依存",
         "SKU 集約に合わせて段取り手順を標準化し、1 回 28 分へ。外段取り化で稼働率 76% を目標", "**+0.3pt**\n段取り 33% 減"),
    ]
    for (y, rh), (dept, *cells) in zip(rows_fill(BY + 28, BB, 4, 6), rows):
        rowhead(s, BX, y, hw - 4, rh, dept, size=T_H2, align=PP_ALIGN.CENTER)
        x = BX + hw
        for (lab, w), c in zip(cols, cells):
            if lab == "効果":
                rect(s, x, y, w - 4, rh, fill=PALE)
                text(s, x, y, w - 4, rh, c, size=T_BODY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            else:
                text(s, x, y, w - 4, rh, c, size=T_BODY, anchor=MSO_ANCHOR.MIDDLE, margins=(6, 2, 6, 2))
            x += w
        line(s, BX + hw, y + rh + 3, BX + BW, y + rh + 3, color=LIGHT)
    source(s, "出典: 部門ヒアリング 24 名(2026 年 6〜7 月)、営業日報 12,400 件、拠点別原価集計、生産実績(6 拠点)")

# =====================================================================
@slide
def s09_vertical_enum():
    s = new_slide("第 1 波の 5 施策は投資 0.6 億円で年 8.2 億円の効果があり、いずれも 9 月までに着手できる",
                  "施策ごとに内容・担当・期限・KPI を並べた。上から効果の大きい順で、上位 2 つで第 1 波の効果の 85% を占める")
    items = [
        ("値引き承認制の導入", "営業が案件・値引き率・理由を申請し、5% 以下は自動承認、超過は本部が 1 営業日で判断する。申請時に案件粗利を表示し、担当者が影響を把握できるようにする。見積には承認番号を付与し、承認のない値引きを無くす", "営業本部 / 6 月", "値引き率 9.7% → 5.0%", "5.8"),
        ("顧客ランク別の値引き上限", "年間売上でランク A/B/C を定義し、自動承認の上限を 5/3/0%、本部承認の上限を 8/5/3% とする。特注品と新規顧客の初回取引は営業部長判断の例外枠を置く", "営業本部 / 6 月", "C ランクの 10% 超 18% → 0%", "1.2"),
        ("下位 SKU 廃番(第 1 弾 900 点)", "年間売上 300 万円未満かつ粗利率 15% 未満の SKU を対象に、代替品の対応表を付けて主要顧客 12 社へ事前照会し、8 月に廃番する", "商品企画 / 8 月", "SKU 3,860 → 2,960", "1.9"),
        ("小口配送の統合", "1 便 20 万円未満の配送を曜日別に束ね、物流会社 2 社と統合便の料率で契約する。上位 50 社は翌日配送を維持", "SCM 部 / 9 月", "小口比率 41% → 25%", "1.0"),
        ("見積テンプレートの統一", "見積書式を 1 種類に統一し、原価・値引き・粗利を自動計算する。承認制の申請データを見積から自動生成する", "営業本部 / 5 月", "見積作成 90 分 → 40 分", "0.3"),
    ]
    x_num, x_t, w_t, x_d, x_o, w_o, x_k, w_k, x_e, w_e = BX, BX + 40, 170, BX + 218, BX + 610, 90, BX + 706, 130, BX + 842, 54
    for lab, xx, ww, al in [("施策", x_t, w_t, PP_ALIGN.LEFT), ("内容", x_d, x_o - x_d - 8, PP_ALIGN.LEFT), ("担当 / 期限", x_o, w_o, PP_ALIGN.LEFT), ("KPI(目標)", x_k, w_k, PP_ALIGN.LEFT), ("効果\n億円/年", x_e, w_e, PP_ALIGN.RIGHT)]:
        text(s, xx, BY, ww, 24, lab, size=T_SMALL, bold=True, color=DARK, align=al, anchor=MSO_ANCHOR.BOTTOM, margins=(0, 0, 0, 2))
    line(s, BX, BY + 26, BX + BW, BY + 26, color=INK, w=1)
    for k, ((y, rh), (t, d, o, kpi, e)) in enumerate(zip(rows_fill(BY + 32, BB, 5, 6), items)):
        shape_text(rect(s, x_num, y, 30, 30, fill=DARK if k < 2 else LIGHT), f"{k + 1:02d}", 12, True, WHITE if k < 2 else INK)
        text(s, x_t, y, w_t, rh, t, size=13, bold=True, margins=(0, 0, 4, 0))
        text(s, x_d, y, x_o - x_d - 8, rh, d, size=11, margins=(0, 0, 0, 0))
        text(s, x_o, y, w_o, rh, o, size=T_BODY, color=DARK, margins=(0, 0, 0, 0))
        text(s, x_k, y, w_k, rh, kpi, size=T_BODY, margins=(0, 0, 0, 0))
        text(s, x_e, y, w_e, rh, e, size=16, bold=True, align=PP_ALIGN.RIGHT, margins=(0, 0, 0, 0))
        line(s, BX, y + rh + 3, BX + BW, y + rh + 3, color=LIGHT)
    source(s, "出典: 分科会別の実行計画(2026 年 8 月版)。効果は FY2027 通期の営業利益への寄与")

# =====================================================================
@slide
def s10_horizontal_enum():
    s = new_slide("施策の実行にあたり、4 つの原則を全分科会に共通で適用し、部分最適と先送りを防ぐ",
                  "原則は過去 2 回の改善活動が途中で止まった原因(効果の未測定・例外の増殖・決定の先送り・現場の巻き込み不足)から逆算した")
    cols = cols_fill(BX, BW, 4, 16)
    items = [
        ("効果は月次で測る", ["施策ごとに KPI と計測方法を着手前に決める", "PMO が月次で実績と計画差を一覧にする", "3 か月連続で未達なら委員会で見直す", "効果の帰属が曖昧な施策は「効果なし」と数える"]),
        ("例外を増やさない", ["例外は営業部長の判断枠に限定し、件数を公表する", "例外の合計が対象の 10% を超えたら基準を見直す", "「今回だけ」の口頭承認を認めない", "例外の理由を 3 分類(特注・新規・戦略)に限定する"]),
        ("決定は委員会で", ["投資判断・中止・KPI 変更は委員会のみが決める", "分科会は判断材料を揃え、期限までに上げる", "先送りは「決めない」という決定として記録する", "委員会の議題は開催 3 営業日前に PMO が確定する"]),
        ("現場を先に巻き込む", ["基準の設計に営業・生産の担当者が参加する", "導入 1 か月前に説明会と Q&A を全拠点で行う", "初月は本部が拠点に常駐して運用を支える", "現場からの改善提案は翌月の委員会で必ず扱う"]),
    ]
    foot_h = 118
    metrics = [("確認指標", "未達 KPI の数(3 か月連続)", "目標: 0 件"), ("確認指標", "例外の件数 / 対象件数", "目標: 10% 未満"), ("確認指標", "先送りした決定の数", "目標: 月 1 件以下"), ("確認指標", "説明会の実施率(拠点)", "目標: 100%")]
    for k, ((x, w), (h, bl), (ml, mv, mt)) in enumerate(zip(cols, items, metrics)):
        y0 = panel_head(s, x, BY, w, f"{k + 1}. {h}")
        mb_h = 54
        text(s, x, y0 + 8, w, BB - foot_h - y0 - 16 - mb_h - 8, bullets(bl, size=12.5), size=12.5)
        my = BB - foot_h - 8 - mb_h
        rect(s, x, my, w, mb_h, fill=PALE)
        text(s, x + 8, my, w - 16, mb_h, [(ml, {"size": T_SMALL, "color": MUTED}), (mv, {"size": T_H2, "bold": True}), (mt, {"size": T_SMALL, "color": DARK})], anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0))
    y = BB - foot_h
    band(s, BX, y, BW, 22, "適用の仕方", fill=DARK, align=PP_ALIGN.LEFT)
    text(s, BX, y + 26, BW, foot_h - 26, bullets([
        "**分科会のキックオフ(9 月第 1 週)で 4 原則を合意し、各分科会の運営ルールに書き込む。** 原則に反する判断が出たときは PMO が委員会に上げる",
        "**原則の運用状況を四半期ごとに委員会で確認する。** 例外の件数、未達 KPI の数、先送りした決定の数を指標にする",
        "**過去の改善活動(2022 年・2024 年)が止まった原因は、いずれも原則 1 と 3 の欠如だった。** 今回は最初の 3 か月でこの 2 つを重点的に守る",
    ], size=11.5), size=11.5)
    source(s, "出典: 過去の改善活動の振り返り(2024 年 12 月、経営企画)、分科会運営ルール案")

# =====================================================================
@slide
def s11_approach_steps():
    s = new_slide("承認制は「申請 → 自動判定 → 本部審査 → 回答」の 4 段で回し、申請から回答まで 2 営業日以内で完了させる",
                  "5% 以下は自動承認で営業の機動性を保ち、超過分だけを本部が判断する。各段の入力・処理・出力と所要時間を揃えて設計した")
    steps = [("申請", "営業担当", "10 分"), ("自動判定", "システム", "即時"), ("本部審査", "営業本部", "1 営業日"), ("回答・実行", "営業担当", "1 営業日")]
    n = len(steps); gap = 6; sw = (BW - 110 - gap * (n - 1)) / n; x0 = BX + 110
    for i, (h, who, t) in enumerate(steps):
        x = x0 + i * (sw + gap)
        sp = rect(s, x, BY, sw, 40, fill=ACCENT if i == 2 else DARK, shape=MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON)
        shape_text(sp, [(f"{i + 1:02d}", {"size": 9, "color": WHITE}), (h, {"size": T_H1, "bold": True, "color": WHITE})], T_H1, True, WHITE, margins=(14, 0, 8, 0))
    rowsdef = [
        ("担当 / 所要", [f"{w}\n{t}" for _, w, t in steps]),
        ("入力", ["案件・顧客・値引き率\n理由(競合・数量・関係)", "申請データ\n顧客ランク・案件粗利", "超過申請\n粗利・競合状況・交渉履歴", "承認結果\n承認番号・条件"]),
        ("処理", ["見積から申請を自動生成\n案件粗利を表示して確認", "ランク別上限(5/3/0%)と照合\n以下なら即時承認", "粗利・ランク・競合で判断\n条件付き承認・却下も可", "顧客へ回答\n見積に承認番号を付与"]),
        ("出力", ["申請(承認番号の仮採番)", "承認 / 本部へ回付", "承認 / 条件付き / 却下", "確定見積・受注"]),
        ("KPI", ["入力率 100%", "即時承認率 85%", "回答 1 営業日以内 95%", "承認番号のない受注 0 件"]),
    ]
    rows = rows_fill(BY + 48, BB, len(rowsdef), 4)
    for (y, rh), (lab, cells) in zip(rows, rowsdef):
        rowhead(s, BX, y, 104, rh, lab, size=T_BODY, align=PP_ALIGN.CENTER)
        for i, c in enumerate(cells):
            x = x0 + i * (sw + gap)
            if lab == "KPI":
                rect(s, x, y, sw, rh, fill=PALE)
                text(s, x, y, sw, rh, c, size=T_BODY, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            elif lab == "担当 / 所要":
                text(s, x, y, sw, rh, c, size=T_BODY, color=DARK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            else:
                text(s, x, y, sw, rh, c, size=T_BODY, anchor=MSO_ANCHOR.MIDDLE, margins=(6, 2, 6, 2))
        line(s, x0, y + rh + 2, BX + BW, y + rh + 2, color=LIGHT)
    source(s, "出典: 承認制の設計書(2026 年 8 月版)。所要時間は同業 B 社の導入実績(公開資料)を参考に設定")

# =====================================================================
@slide
def s12_swimlane():
    s = new_slide("値引き承認の業務フローは営業・システム・本部・顧客の 4 レーンで完結し、人の判断は本部審査の 1 か所に絞る",
                  "分岐は「5% 以下か」「承認するか」の 2 つだけにし、差し戻しは条件付き承認で吸収する。例外(特注品・新規顧客)は営業部長のレーン外判断")
    lanes = ["営業担当", "システム", "営業本部", "顧客"]
    lw = 80; lx = BX + lw; lwid = BW - lw
    lrows = rows_fill(BY, BB, 4, 0)
    for (y, h), lab in zip(lrows, lanes):
        rowhead(s, BX, y + 1, lw - 6, h - 2, lab, size=T_BODY, align=PP_ALIGN.CENTER, fill=LIGHT)
        line(s, lx, y + h, BX + BW, y + h, color=LIGHT)
    rect(s, lx, BY, lwid, BH, line=LIGHT)
    bw_, bh_ = 104, 36
    def ly(i):
        return lrows[i][0] + lrows[i][1] / 2
    def box(cx, lane, label, fill=WHITE, shape=MSO_SHAPE.RECTANGLE, w=bw_, h=bh_, size=T_SMALL + 0.5):
        sp = rect(s, cx - w / 2, ly(lane) - h / 2, w, h, fill=fill, line=DARK if fill == WHITE else None, lw=0.75, shape=shape)
        shape_text(sp, label, size, False, INK if fill != DARK else WHITE, margins=(2, 0, 2, 0))
        return sp
    xs = [lx + 70, lx + 200, lx + 330, lx + 460, lx + 590, lx + 720]
    box(xs[0], 0, "見積作成\n(粗利を自動表示)", fill=PALE)
    box(xs[1], 0, "値引き申請\n(理由・率を入力)")
    box(xs[1], 1, "ランク別上限\nと照合")
    box(xs[2], 1, "5%以下?", fill=WHITE, shape=MSO_SHAPE.DIAMOND, w=96, h=52)
    box(xs[3], 2, "本部審査\n(粗利・競合・履歴)")
    box(xs[4], 2, "承認?", fill=WHITE, shape=MSO_SHAPE.DIAMOND, w=70, h=48)
    box(xs[4], 1, "承認番号を付与\n見積を確定", fill=PALE)
    box(xs[5], 0, "顧客へ回答\n(確定見積)")
    box(xs[5], 3, "受注 / 発注", fill=DARK)
    box(xs[3], 0, "再申請 or 取り下げ\n(条件付きの場合)")
    def A_(x1, y1, x2, y2):
        return line(s, x1, y1, x2, y2, color=DARK, w=1, arrow=True)
    A_(xs[0] + bw_ / 2, ly(0), xs[1] - bw_ / 2, ly(0))
    A_(xs[1], ly(0) + bh_ / 2, xs[1], ly(1) - bh_ / 2)
    A_(xs[1] + bw_ / 2, ly(1), xs[2] - 48, ly(1))
    A_(xs[2] + 48, ly(1), xs[4] - bw_ / 2, ly(1))
    text(s, xs[2] + 50, ly(1) - 16, 60, 12, "Yes(85%)", size=8, color=DARK, margins=(0, 0, 0, 0))
    line(s, xs[2], ly(1) + 26, xs[2], ly(2), color=DARK, w=1)
    A_(xs[2], ly(2), xs[3] - bw_ / 2, ly(2))
    text(s, xs[2] + 4, ly(1) + 30, 60, 12, "No(15%)", size=8, color=DARK, margins=(0, 0, 0, 0))
    A_(xs[3] + bw_ / 2, ly(2), xs[4] - 35, ly(2))
    A_(xs[4], ly(2) - 24, xs[4], ly(1) + bh_ / 2)
    text(s, xs[4] + 4, ly(2) - 40, 60, 12, "承認", size=8, color=DARK, margins=(0, 0, 0, 0))
    line(s, xs[4] + 35, ly(2), xs[5], ly(2), color=DARK, w=1)
    A_(xs[5], ly(2), xs[5], ly(0) + bh_ / 2)
    text(s, xs[4] + 38, ly(2) - 16, 90, 12, "条件付き / 却下", size=8, color=DARK, margins=(0, 0, 0, 0))
    line(s, xs[4] + bw_ / 2, ly(1), xs[5], ly(1), color=DARK, w=1)
    A_(xs[5] - bw_ / 2, ly(0), xs[3] + bw_ / 2, ly(0))
    A_(xs[5], ly(0) + bh_ / 2, xs[5], ly(3) - bh_ / 2)
    text(s, lx + 6, BB - 16, lwid - 12, 14, "所要: 申請〜回答 2 営業日以内(自動承認は即時)。フロー外: 特注品・新規顧客の初回取引は営業部長判断(案件粗利 20% 以上)", size=8, color=MUTED, margins=(0, 0, 0, 0))
    source(s, "出典: 承認制の業務設計(2026 年 8 月版)。分岐の比率は FY2025 の申請実績で試算")

# =====================================================================
@slide
def s13_gantt():
    s = new_slide("値引き統制を 4 月に開始し、拠点集約を翌年 3 月までに完了させる。効果は FY26 下期から段階的に発現する",
                  "第 1 波(値引き・SKU・小口統合)は投資 0.6 億円で年 8.2 億円。拠点集約は 6 月の投資判断後、並行稼働 6 か月を経て 3 月に完了する")
    lw, ow = 170, 96; gx = BX + lw; gw = BW - lw - ow
    months = ["4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", "1月", "2月", "3月", "4月", "5月", "6月"]
    pw_ = gw / len(months)
    band(s, gx, BY, pw_ * 12 - 2, 16, "2026 年度", fill=LIGHT, color=INK, size=9)
    band(s, gx + pw_ * 12, BY, pw_ * 3 - 2, 16, "2027 年度", fill=LIGHT, color=INK, size=9)
    for i, m in enumerate(months):
        text(s, gx + pw_ * i, BY + 18, pw_, 14, m, size=T_NOTE, bold=True, align=PP_ALIGN.CENTER, margins=(0, 0, 0, 0))
    text(s, gx + gw, BY + 18, ow, 14, "担当", size=T_NOTE, bold=True, margins=(4, 0, 0, 0))
    line(s, BX, BY + 34, BX + BW, BY + 34, color=DARK, w=1)
    tasks = [("値引き承認制の導入", 0, 2, "営業本部", True), ("顧客ランク別の上限設定", 1, 2, "営業本部", False), ("SKU 廃番 第 1 弾(900 点)", 1, 4, "商品企画", False),
             ("SKU 廃番 第 2 弾(640 点)", 6, 9, "商品企画", False), ("小口配送の統合", 2, 5, "SCM 部", False), ("拠点集約: 設計・投資判断", 2, 5, "SCM 部", True),
             ("拠点集約: 契約・移転準備", 5, 6, "SCM 部", False), ("拠点集約: 並行稼働・移転", 6, 11, "SCM 部 + 物流会社", True), ("共同配送の開始", 9, 11, "SCM 部", False),
             ("医療機器向け: 認証取得", 3, 14, "新規事業室", False), ("効果測定・是正(月次)", 6, 14, "PMO", False)]
    note_h = 16; ms_h = 24
    rows = rows_fill(BY + 36, BB - note_h - ms_h - 4, len(tasks), 0)
    for (y, rh), (lab, a, b, owner, crit_) in zip(rows, tasks):
        text(s, BX, y, lw - 4, rh, lab, size=T_SMALL + 0.5, anchor=MSO_ANCHOR.MIDDLE, margins=(4, 0, 0, 0))
        rect(s, gx + pw_ * a + 2, y + 5, pw_ * (b - a + 1) - 4, rh - 10, fill=ACCENT if crit_ else MID, shape=MSO_SHAPE.PENTAGON)
        text(s, gx + gw, y, ow, rh, owner, size=T_NOTE, color=DARK, anchor=MSO_ANCHOR.MIDDLE, margins=(4, 0, 0, 0))
        line(s, BX, y + rh, BX + BW, y + rh, color=PALE)
    y = BB - note_h - ms_h - 2
    text(s, BX, y, lw - 4, ms_h, "節目", size=T_SMALL + 0.5, bold=True, anchor=MSO_ANCHOR.MIDDLE, margins=(4, 0, 0, 0))
    for lab, at in [("承認制 全社適用", 2), ("投資判断", 5), ("集約完了", 11), ("ISO 13485", 14)]:
        cx = gx + pw_ * at + pw_ / 2
        rect(s, cx - 5, y + ms_h / 2 - 5, 10, 10, fill=DARK, shape=MSO_SHAPE.DIAMOND)
        text(s, cx + 6, y, 80, ms_h, lab, size=T_NOTE, anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0))
    for i in range(1, len(months)):
        line(s, gx + pw_ * i, BY + 34, gx + pw_ * i, y + ms_h, color=PALE)
    tx = gx + pw_ * 5.25
    line(s, tx, BY + 34, tx, y + ms_h, color=DARK, w=1, dash=True)
    shape_text(rect(s, tx - 22, BY + 20, 44, 12, fill=DARK), "今日 9/7", 7.5, True, WHITE)
    text(s, BX, BB - note_h + 2, BW, note_h, "■ 橙 = クリティカルパス(遅れると効果の発現時期がずれる)  ■ 灰 = その他。並行稼働期(10〜3 月)は在庫を 0.8 億円積み増し、納期遵守率 98% を維持する", size=T_NOTE, color=DARK, margins=(0, 0, 0, 0))
    source(s, "出典: 分科会別の実行計画(2026 年 8 月版)。効果の発現時期は各施策の完了月の翌月から計上")

# =====================================================================
@slide
def s14_org():
    s = new_slide("社長直轄の推進委員会の下に 3 分科会を置き、PMO が横串で効果と課題を月次で判断する",
                  "分科会は週次で実行し、意思決定は月次の委員会に集約する。分科会への権限委譲は行わず、投資判断と施策の中止は委員会が決める")
    def box(x, y, w, h, t, sub, dark=False):
        shape_text(rect(s, x, y, w, h, fill=DARK if dark else PALE), [(t, {"size": T_H2, "bold": True, "color": WHITE if dark else INK}), (sub, {"size": T_SMALL, "color": WHITE if dark else DARK})], T_H2, True, WHITE if dark else INK)
    ox, ow_ = BX, 560; bw_, bh_ = 180, 44; cx = ox + ow_ / 2
    box(cx - bw_ / 2, BY, bw_, bh_, "収益改善推進委員会", "委員長: 社長 / 月次", dark=True)
    line(s, cx, BY + bh_, cx, BY + bh_ + 14, color=DARK, w=1)
    y1 = BY + bh_ + 14
    box(cx - bw_ / 2, y1, bw_, bh_, "PMO", "経営企画 3 名(専任) / 週次")
    box(cx + bw_ / 2 + 30, y1, 150, bh_, "外部支援", "コンサル 4 名、物流設計 2 名")
    line(s, cx + bw_ / 2, y1 + bh_ / 2, cx + bw_ / 2 + 30, y1 + bh_ / 2, color=MID, w=1, dash=True)
    y2 = y1 + bh_ + 22
    line(s, cx, y1 + bh_, cx, y2 - 8, color=DARK, w=1)
    ws = [("値引き統制 分科会", "営業本部長 / 8 名"), ("SKU 整理 分科会", "商品企画部長 / 5 名"), ("拠点集約 分科会", "SCM 部長 / 6 名 + 物流 2 社")]
    wcols = cols_fill(ox, ow_, 3, 12)
    line(s, wcols[0][0] + wcols[0][1] / 2, y2 - 8, wcols[2][0] + wcols[2][1] / 2, y2 - 8, color=DARK, w=1)
    for (x, ww), (t, sub) in zip(wcols, ws):
        line(s, x + ww / 2, y2 - 8, x + ww / 2, y2, color=DARK, w=1)
        box(x, y2, ww, bh_, t, sub)
    ty = y2 + bh_ + 10
    for (x, ww), (kpi, tasks_) in zip(wcols, [("値引き率 9.7% → 5.0%", ["承認フロー・上限の設計(4 月)", "全社適用と定着(6 月)", "評価指標への反映(FY27)"]),
                                              ("SKU 3,860 → 2,320", ["廃番基準と代替品の設計(5 月)", "第 1 弾 900 点(8 月)", "第 2 弾 640 点(翌 1 月)"]),
                                              ("物流費率 9.4% → 8.3%", ["小口統合(9 月)", "拠点設計・投資判断(6 月)", "並行稼働・移転(翌 3 月)"])]):
        text(s, x, ty, ww, 16, f"KPI: {kpi}", size=T_SMALL + 0.5, bold=True, margins=(2, 0, 2, 0))
        line(s, x, ty + 17, x + ww, ty + 17, color=INK, w=0.75)
        text(s, x, ty + 20, ww, 56, bullets(tasks_), size=T_SMALL, margins=(2, 0, 2, 0))
    by = ty + 84
    rect(s, ox, by, ow_, BB - by, fill=PALE)
    text(s, ox + 8, by, ow_ - 16, BB - by, [("**体制の規模**: 専任 3 名 + 兼任 19 名 + 外部 6 名。兼任者は稼働の 20〜30% を充て、分科会の週次会議は 60 分に固定する", {}), ("**KPI の持ち方**: 委員会は営業利益率、分科会は値引き率・SKU 数・物流費率を月次で追う。PMO は 3 指標の計画差を毎月 5 営業日以内に集計する", {"sb": 4}), ("**外部支援の関わり**: 設計は主導、実行は分科会が主体。常駐は第 1 波の定着期(4〜6 月)に限る", {"sb": 4})], size=T_BODY, anchor=MSO_ANCHOR.MIDDLE)
    rx = ox + ow_ + 24; rw = BX + BW - rx
    band(s, rx, BY, rw, 22, "会議体と役割", align=PP_ALIGN.LEFT)
    rows = [("推進委員会", "月次 90 分", "投資判断、施策の中止・追加、効果の確認"), ("PMO", "週次", "進捗・効果・課題の一元管理、委員会資料"), ("分科会", "週次", "施策の実行、KPI の計測、課題の一次判断"), ("外部支援", "随時", "設計・分析・他社事例の提供")]
    y = BY + 28
    for a, b, c in rows:
        text(s, rx, y, 68, 40, a, size=T_SMALL + 0.5, bold=True, margins=(0, 0, 0, 0))
        text(s, rx + 70, y, 56, 40, b, size=T_SMALL, color=DARK, margins=(0, 0, 0, 0))
        text(s, rx + 128, y, rw - 128, 40, c, size=T_SMALL, margins=(0, 0, 0, 0))
        line(s, rx, y + 40, rx + rw, y + 40, color=PALE); y += 44
    rect(s, rx, y + 6, rw, BB - y - 6, fill=PALE)
    text(s, rx + 6, y + 6, rw - 12, BB - y - 6, [("**意思決定の原則**", {"size": T_H2}), ("分科会に権限を委譲しない。投資判断・施策の中止・KPI の変更は委員会のみが決め、PMO は判断材料を揃える", {"sb": 4}), ("委員会で決めなかった事項は「先送り」として記録し、翌月の議題に自動で載せる", {"sb": 4})], size=T_SMALL + 0.5, anchor=MSO_ANCHOR.MIDDLE)
    source(s, "出典: 推進体制案(2026 年 8 月、経営会議承認)。人数は各部門の指名者ベース")

# =====================================================================
@slide
def s15_asis_tobe():
    s = new_slide("値引きは「営業担当の裁量」から「本部が粗利で判断する仕組み」へ変え、5 つの観点すべてで基準を明文化する",
                  "権限・基準・プロセス・可視化・評価の 5 観点で As-Is と To-Be を対比した。To-Be は同業 B 社の運用を参考に、X 社の顧客構成に合わせて設定")
    hw = 110; aw = 356; ar = 40; tw = BW - hw - aw - ar
    band(s, BX, BY, hw - 4, 24, "観点", fill=DARK)
    band(s, BX + hw, BY, aw - 4, 24, "As-Is: 営業担当の裁量", fill=LIGHT, color=INK, align=PP_ALIGN.LEFT)
    band(s, BX + hw + aw + ar, BY, tw, 24, "To-Be: 本部承認制", fill=ACCENT, align=PP_ALIGN.LEFT)
    rows = [("権限", "値引き率に上限なし。担当者が案件ごとに判断し、事後報告のみ", "5% 以下は自動承認、超過は本部が 1 営業日で判断。承認番号のない値引きは無効"),
            ("基準", "顧客ランクと値引きが連動せず、下位顧客にも 10% 超が 18%", "顧客ランク A/B/C ごとに上限 8% / 5% / 3%。例外は営業部長の判断枠に限定"),
            ("プロセス", "承認は事後報告。月次の営業会議で集計を眺めるだけ", "申請 → 自動判定 → 本部審査 → 回答を 2 営業日で完了。見積から申請を自動生成"),
            ("可視化", "粗利は四半期集計。担当者が案件の影響を把握していない", "申請時に案件粗利を自動表示。月次で担当者別・顧客別の値引き率を共有"),
            ("評価", "売上高のみで評価。値引きは評価に反映されない", "粗利額と値引き率を評価指標に加える(FY27 から)。初年度は可視化に集中")]
    kp_h = 76
    for (y, rh), (a, b, c) in zip(rows_fill(BY + 30, BB - kp_h - 8, 5, 6), rows):
        rowhead(s, BX, y, hw - 4, rh, a, size=T_H2, align=PP_ALIGN.CENTER)
        text(s, BX + hw, y, aw - 4, rh, b, size=T_BODY, anchor=MSO_ANCHOR.MIDDLE, margins=(6, 2, 6, 2))
        rect(s, BX + hw + aw + 8, y + rh / 2 - 7, ar - 16, 14, fill=DARK, shape=MSO_SHAPE.RIGHT_ARROW)
        rect(s, BX + hw + aw + ar, y, tw, rh, fill=PALE)
        text(s, BX + hw + aw + ar, y, tw, rh, c, size=T_BODY, anchor=MSO_ANCHOR.MIDDLE, margins=(6, 2, 6, 2))
        line(s, BX + hw, y + rh + 3, BX + hw + aw - 4, y + rh + 3, color=LIGHT)
    y = BB - kp_h
    band(s, BX, y, BW, 20, "移行の要点", fill=DARK, align=PP_ALIGN.LEFT, size=T_BODY)
    text(s, BX, y + 24, BW, kp_h - 24, bullets([
        "**4 月に承認制を開始し、6 月に全社適用。** 初月は超過申請が 34% 出る見込みだが、3 か月で 15% 以下に収束させる(B 社実績)",
        "**評価への反映は FY27 から。** FY26 は粗利の可視化に集中し、担当者が自分の値引きの影響を把握できる状態を先に作る",
    ]), size=T_BODY)
    source(s, "出典: 営業日報 12,400 件(FY2025)、営業部門ヒアリング 24 名、同業 B 社の公開資料")

# =====================================================================
@slide
def s16_action_table():
    s = new_slide("12 施策を 3 波に分け、第 1 波の 5 施策で効果の 6 割を FY26 内に確定させる",
                  "第 1 波は投資 0.6 億円で年 8.2 億円。第 2 波は拠点集約を軸に投資 3.2 億円で年 6.7 億円。第 3 波は成長投資で FY27 以降に効果が出る")
    hdr = ["#", "施策", "担当", "着手", "完了", "投資\n(億円)", "効果\n(億円/年)", "KPI(目標)"]
    widths = [26, 232, 84, 50, 56, 52, 62, BW - 26 - 232 - 84 - 50 - 56 - 52 - 62]
    data = [["第 1 波(4〜9 月): 統制と整理"],
            ["1", "値引き承認制の導入", "営業本部", "4 月", "6 月", "0.1", "5.8", "平均値引き率 9.7% → 5.0%"],
            ["2", "顧客ランク別の値引き上限", "営業本部", "5 月", "6 月", "0.0", "1.2", "C ランクの 10% 超 18% → 0%"],
            ["3", "下位 SKU 廃番(第 1 弾 900 点)", "商品企画", "5 月", "8 月", "0.2", "1.9", "SKU 3,860 → 2,960"],
            ["4", "小口配送の統合", "SCM 部", "6 月", "9 月", "0.1", "1.0", "小口比率 41% → 25%"],
            ["5", "見積テンプレート統一", "営業本部", "4 月", "5 月", "0.2", "0.3", "見積作成 90 分 → 40 分"],
            ["第 2 波(6 月〜翌 3 月): 構造の是正"],
            ["6", "3 拠点集約", "SCM 部", "6 月", "翌 3 月", "3.1", "5.3", "物流費率 9.4% → 8.3%、納期遵守 98%"],
            ["7", "下位 SKU 廃番(第 2 弾 640 点)", "商品企画", "10 月", "翌 1 月", "0.1", "1.4", "SKU 2,960 → 2,320"],
            ["8", "段取り時間の短縮", "生産技術", "7 月", "12 月", "0.6", "1.5", "段取り 42 分 → 28 分"],
            ["第 3 波(7 月〜翌 6 月): 成長"],
            ["9", "医療機器向け認証取得(ISO 13485)", "新規事業室", "7 月", "翌 6 月", "0.4", "—", "認証取得、初年度受注 3 億円"],
            ["10", "医療機器向け営業体制", "新規事業室", "翌 1 月", "翌 6 月", "0.8", "—", "専任 4 名、見込み客 30 社"],
            ["合計", "", "", "", "", "5.6", "17.1", ""]]
    rows_n, cols_n = len(data) + 1, len(hdr)
    gt = s.shapes.add_table(rows_n, cols_n, pt(BX), pt(BY), pt(BW), pt(BH))
    tbl = gt.table
    tblPr = gt._element.graphic.graphicData.tbl.tblPr
    for attr in ("firstRow", "bandRow", "lastRow", "firstCol", "lastCol", "bandCol"):
        tblPr.set(attr, "0")
    for j, w in enumerate(widths):
        tbl.columns[j].width = pt(w)
    def cell(r, c, s_, size=T_BODY, bold=False, color=INK, fill=None, align=PP_ALIGN.LEFT):
        ce = tbl.cell(r, c); ce.margin_left = ce.margin_right = pt(4); ce.margin_top = ce.margin_bottom = pt(1)
        ce.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = ce.text_frame; tf.text = ""; p = tf.paragraphs[0]; p.alignment = align; _rich(p, s_, size, bold, color)
        if fill is None:
            ce.fill.background()
        else:
            ce.fill.solid(); ce.fill.fore_color.rgb = fill
    for j, h in enumerate(hdr):
        cell(0, j, h, size=T_H2, bold=True, color=WHITE, fill=DARK, align=PP_ALIGN.CENTER if j in (0, 3, 4, 5, 6) else PP_ALIGN.LEFT)
    n_sec = sum(1 for r in data if len(r) == 1); n_item = len(data) - n_sec
    sec_h = 20; hdr_h = 28; item_h = (BH - hdr_h - sec_h * n_sec) / n_item
    tbl.rows[0].height = pt(hdr_h)
    for i, row in enumerate(data, start=1):
        if len(row) == 1:
            tbl.rows[i].height = pt(sec_h)
            cell(i, 0, row[0], size=T_BODY, bold=True, fill=PALE); tbl.cell(i, 0).merge(tbl.cell(i, cols_n - 1)); continue
        tbl.rows[i].height = pt(item_h)
        total = row[0] == "合計"
        if total:
            tbl.cell(i, 0).merge(tbl.cell(i, 1)); row = ["合計", ""] + row[2:]
        for j, v in enumerate(row):
            num = j in (5, 6)
            cell(i, j, v, size=T_BODY, bold=total, align=PP_ALIGN.RIGHT if num else (PP_ALIGN.CENTER if j in (0, 3, 4) else PP_ALIGN.LEFT), fill=PALE if total else None)
    def set_borders(ce, bottom=True):
        tcPr = ce._tc.get_or_add_tcPr()
        for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            for el in tcPr.findall(qn(tag)):
                tcPr.remove(el)
            ln = etree.SubElement(tcPr, qn(tag))
            if tag == "a:lnB" and bottom:
                ln.set("w", "6350"); sf = etree.SubElement(ln, qn("a:solidFill")); c_ = etree.SubElement(sf, qn("a:srgbClr")); c_.set("val", "D9D9D9")
            else:
                ln.set("w", "0"); etree.SubElement(ln, qn("a:noFill"))
    for i in range(rows_n):
        for j in range(cols_n):
            set_borders(tbl.cell(i, j), bottom=i > 0)
    source(s, "効果は FY2027 通期の営業利益への寄与(億円)。投資は設備・システム・外部委託の合計。第 3 波の効果は FY28 以降に計上するため「—」")

# =====================================================================
@slide
def s17_kpi_dashboard():
    s = new_slide("8 月時点で値引き率は 6.8% まで下がり計画を上回るが、SKU 廃番と小口統合は遅れており 9 月に挽回が必要",
                  "第 1 波 5 施策のうち 3 つが計画どおり。SKU 廃番は営業からの個別要望で 180 点が保留、小口統合は物流会社との契約が 1 か月遅れている")
    tiles = [("平均値引き率", "6.8%", "計画 7.5% / 目標 5.0%", "▲ 計画比 −0.7pt", GREEN), ("SKU 数", "3,140", "計画 2,960 / 目標 2,320", "▼ 計画比 +180", RED),
             ("小口配送比率", "36%", "計画 31% / 目標 25%", "▼ 計画比 +5pt", AMBER), ("営業利益率(4〜8 月累計)", "3.4%", "計画 3.2% / FY27 目標 6.5%", "▲ 計画比 +0.2pt", GREEN)]
    th = 78
    for (x, tw), (lab, val, plan, delta, col) in zip(cols_fill(BX, BW, 4, 12), tiles):
        rect(s, x, BY, tw, th, fill=PALE); rect(s, x, BY, 4, th, fill=col)
        text(s, x + 10, BY + 4, tw - 14, 14, lab, size=T_SMALL, color=DARK, margins=(0, 0, 0, 0))
        text(s, x + 10, BY + 18, tw - 14, 34, val, size=T_KPI, bold=True, margins=(0, 0, 0, 0))
        text(s, x + 10, BY + 50, tw - 14, 12, plan, size=8, color=MUTED, margins=(0, 0, 0, 0))
        text(s, x + tw - 110, BY + 50, 104, 24, delta, size=T_SMALL, bold=True, color=col, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 4, 0))
    y1 = BY + th + 12; cw = 470; rx = BX + cw + 24; rw = BX + BW - rx
    band(s, BX, y1, cw, 20, "平均値引き率の月次推移(%)", align=PP_ALIGN.LEFT, size=T_BODY)
    band(s, rx, y1, rw, 20, "第 1 波 5 施策の状態と 9 月の対応", align=PP_ALIGN.LEFT, size=T_BODY)
    chh = BB - y1 - 24 - 74
    cd = CategoryChartData(); cd.categories = ["4月", "5月", "6月", "7月", "8月"]
    cd.add_series("実績", (9.4, 8.9, 8.1, 7.3, 6.8)); cd.add_series("計画", (9.2, 8.7, 8.2, 7.8, 7.5))
    ch = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, pt(BX), pt(y1 + 22), pt(cw), pt(chh), cd).chart
    style_chart(ch, size=8.5, legend=XL_LEGEND_POSITION.BOTTOM, gridlines=True)
    ch.value_axis.minimum_scale = 5; ch.value_axis.maximum_scale = 10
    line_series(ch.series[0], ACCENT, 2, pos=XL_LABEL_POSITION.BELOW, size=8.5); line_series(ch.series[1], MID, 1.25, labels=False, dash=True)
    ty = y1 + 22 + chh + 2
    text(s, BX, ty, cw, 14, "超過申請(5% 超)の件数と比率", size=T_SMALL, bold=True, margins=(0, 0, 0, 0))
    line(s, BX, ty + 15, BX + cw, ty + 15, color=INK, w=0.75)
    ex_rows = [("", "4月", "5月", "6月", "7月", "8月"), ("申請件数", "1,020", "1,080", "1,140", "1,110", "1,090"), ("超過申請", "347", "292", "228", "156", "131"), ("超過比率", "34%", "27%", "20%", "14%", "12%")]
    cw_ = [cw - 68 * 5] + [68] * 5; yy = ty + 17
    for i, r in enumerate(ex_rows):
        xx = BX
        for j, v in enumerate(r):
            text(s, xx, yy, cw_[j], 13, v, size=8.5, bold=(i in (0, 3)), color=DARK if i == 0 else INK, align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT, margins=(0, 0, 0, 0), anchor=MSO_ANCHOR.MIDDLE)
            xx += cw_[j]
        line(s, BX, yy + 14, BX + cw, yy + 14, color=PALE); yy += 14
    st = [("値引き承認制", GREEN, "計画どおり", "超過申請 12%(目標 15% 以下を 7 月に達成)"), ("ランク別上限", GREEN, "計画どおり", "C ランクの 10% 超は 0 件に"),
          ("SKU 廃番 第 1 弾", RED, "180 点が保留", "9 月委員会で顧客ごとに存廃を決定。代替品提示を条件に廃番"),
          ("小口配送の統合", AMBER, "契約 1 か月遅れ", "9/15 に契約。10 月から統合便、12 月に計画へ復帰"), ("見積テンプレート", GREEN, "完了", "5 月に全拠点で切替済み")]
    for (y, rh), (nm, col, stat, act) in zip(rows_fill(y1 + 26, BB, 5, 3), st):
        dot(s, rx + 7, y + rh / 2, 10, col)
        text(s, rx + 16, y, 100, rh, nm, size=T_SMALL + 0.5, bold=True, anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0))
        text(s, rx + 118, y, 82, rh, stat, size=T_SMALL, color=col, bold=True, anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0))
        text(s, rx + 202, y, rw - 202, rh, act, size=T_SMALL, anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0))
        line(s, rx, y + rh + 1, rx + rw, y + rh + 1, color=PALE)
    source(s, "出典: PMO 月次報告(2026 年 8 月度)。計画は 2026 年 4 月の承認計画。● 計画どおり ● 要注意 ● 遅延")

# =====================================================================
@slide
def s18_issue_tree():
    s = new_slide("「なぜ利益が減ったか」は 2 つの側面・6 つの論点に分解でき、5 つは検証済み、価格転嫁の可否だけが未検証",
                  "売る側(価格・商品)と作る側(生産・物流)に分け、各論点に仮説と検証結果を置いた。検証済みの 5 論点で低下 6.1pt のうち 5.6pt を説明できる")
    def tbox(x, y, w, h, t, sub=None, dark=False, fill=None):
        sp = rect(s, x, y, w, h, fill=DARK if dark else (fill or PALE))
        if sub:
            shape_text(sp, [(t, {"size": T_H2, "bold": True, "color": WHITE if dark else INK}), (sub, {"size": T_SMALL, "color": WHITE if dark else DARK})], T_H2, True, WHITE if dark else INK)
        else:
            shape_text(sp, t, T_BODY, True, WHITE if dark else INK)
    x0, w0 = BX, 120; x1, w1 = BX + 150, 120; x2, w2 = BX + 300, 250; x3 = BX + 570; w3 = BX + BW - x3
    hdr_y = BY
    band(s, x2, hdr_y, w2, 18, "論点", fill=LIGHT, color=INK, size=9)
    band(s, x3, hdr_y, w3, 18, "仮説の検証結果(粗利率への影響)", fill=LIGHT, color=INK, size=9, align=PP_ALIGN.LEFT)
    foot_h = 40
    leaves = [("売る側", [("価格: 値引きが拡大していないか", "**2.7pt** 平均値引き 4.8% → 9.7%。上位 20 顧客に 73% 集中", "検証済"),
                        ("商品: SKU が増えて非効率になっていないか", "**0.6pt** 3,860 SKU、段取り月 1,900 回。下位 40% は売上 3.2%", "検証済"),
                        ("価格転嫁: 原材料高を転嫁できていないか", "原材料 +12% に対し価格改定 +3%。転嫁の余地は顧客交渉で未確認", "未検証")]),
              ("作る側", [("生産: 稼働率・歩留まりが落ちていないか", "**0.3pt** 稼働率 71%(−4pt)。歩留まりは横ばい", "検証済"),
                        ("物流: 拠点・配送の固定費が重くないか", "**1.5pt** 稼働率 58%、小口配送 41%", "検証済"),
                        ("原材料: 市場要因はどれだけか", "**1.8pt** 特殊鋼 +12%。同業 8 社も同条件", "検証済")])]
    rows = rows_fill(hdr_y + 24, BB - foot_h - 8, 6, 6)
    for g, (grp, items) in enumerate(leaves):
        ys = [rows[g * 3 + k] for k in range(3)]
        gy = ys[0][0]; gh = ys[2][0] + ys[2][1] - gy
        tbox(x1, gy + gh / 2 - 20, w1, 40, grp)
        line(s, x1 + w1, gy + gh / 2, x2 - 12, gy + gh / 2, color=DARK, w=1)
        line(s, x2 - 12, ys[0][0] + ys[0][1] / 2, x2 - 12, ys[2][0] + ys[2][1] / 2, color=DARK, w=1)
        for (yy, rh), (q, ev, st) in zip(ys, items):
            line(s, x2 - 12, yy + rh / 2, x2, yy + rh / 2, color=DARK, w=1)
            tbox(x2, yy, w2, rh, q, fill=PALE if st == "検証済" else WHITE)
            if st != "検証済":
                rect(s, x2, yy, w2, rh, line=DARK, lw=1)
            text(s, x3, yy, w3 - 60, rh, ev, size=T_BODY, anchor=MSO_ANCHOR.MIDDLE, margins=(4, 0, 0, 0))
            shape_text(rect(s, x3 + w3 - 56, yy + rh / 2 - 10, 52, 20, fill=DARK if st == "検証済" else WHITE, line=DARK), st, T_NOTE, True, WHITE if st == "検証済" else INK)
    mid = (rows[0][0] + rows[5][0] + rows[5][1]) / 2
    tbox(x0, mid - 26, w0, 52, "なぜ利益が\n減ったか", "営業利益率 5.9% → 2.1%", dark=True)
    line(s, x0 + w0, mid, x1 - 12, mid, color=DARK, w=1)
    g1 = (rows[0][0] + rows[2][0] + rows[2][1]) / 2; g2 = (rows[3][0] + rows[5][0] + rows[5][1]) / 2
    line(s, x1 - 12, g1, x1 - 12, g2, color=DARK, w=1)
    line(s, x1 - 12, g1, x1, g1, color=DARK, w=1); line(s, x1 - 12, g2, x1, g2, color=DARK, w=1)
    rect(s, BX, BB - foot_h, BW, foot_h, fill=PALE)
    text(s, BX + 8, BB - foot_h, BW - 16, foot_h, "未検証の「価格転嫁」は、9 月の上位 20 顧客との交渉で余地を確認する。転嫁できれば原材料高 1.8pt のうち 0.9pt を回収でき、営業利益率の目標を 6.5% → 7.4% に引き上げられる", size=T_BODY, anchor=MSO_ANCHOR.MIDDLE)
    source(s, "出典: 原価差異分析(FY2023 → FY2025)、営業日報 12,400 件、生産実績(6 拠点)。影響は粗利率への寄与(pt)")

# =====================================================================
@slide
def s19_status_report():
    s = new_slide("第 6 週の全体進捗は「概ね順調」。値引き統制は先行して効果が出ており、SKU 廃番の保留 180 点の判断を今週の委員会に諮る",
                  "9 月 1 日〜5 日の実績と 9 月 8 日〜12 日の予定、対応が必要な課題・リスクをまとめた。委員会での決定事項は 2 件")
    y = BY
    for (x, w), (lab, col, txt) in zip(cols_fill(BX, BW, 4, 10), [("全体", GREEN, "概ね順調"), ("値引き統制", GREEN, "先行(効果発現)"), ("SKU 整理", RED, "遅延(180 点保留)"), ("拠点集約", AMBER, "要注意(契約遅れ)")]):
        rect(s, x, y, w, 30, fill=PALE); dot(s, x + 14, y + 15, 12, col)
        text(s, x + 26, y, w - 30, 30, [(lab, {"size": T_SMALL, "color": DARK}), (txt, {"size": T_H2, "bold": True})], anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0))
    y = BY + 40
    cols = cols_fill(BX, BW, 3, 20)
    heads = ["今週の実績(9/1〜9/5)", "来週の予定(9/8〜9/12)", "委員会への決定依頼"]
    contents = [
        bullets(["#値引き統制", "8 月の平均値引き率 6.8%(計画 7.5%)。超過申請 12%", "A ランク 20 社の上限交渉が 14 社で完了", "#SKU 整理", "第 1 弾 900 点のうち 720 点を廃番済み", "180 点は営業からの保留要望で判断待ち", "#拠点集約", "物流会社 2 社と統合便の料率で合意(9/15 契約)", "3 拠点案の配送シミュレーション完了(納期遵守 98.2%)"]),
        bullets(["#値引き統制", "残る A ランク 6 社の上限交渉", "月次の担当者別値引き率を営業会議で初共有(9/10)", "#SKU 整理", "保留 180 点を顧客別に整理し、代替品案を添付", "第 2 弾 640 点の候補リスト作成に着手", "#拠点集約", "統合便の契約締結(9/15)、10 月開始の配車計画", "投資判断資料(3.1 億円)の最終化"]),
        bullets(["#決定 1: 保留 SKU 180 点の扱い", "提案: 代替品を提示できる 150 点は廃番、専用品 30 点は顧客と個別協議", "影響: 廃番すれば効果 +0.3 億円/年、遅らせると第 2 弾もずれる", "#決定 2: 拠点集約の投資判断", "提案: 3 拠点案で 3.1 億円を承認し、10 月から並行稼働", "影響: 1 か月遅れると FY27 の効果が 0.4 億円減る"]),
    ]
    risk_h = 112
    for (x, w), h, c in zip(cols, heads, contents):
        y0 = panel_head(s, x, y, w, h, size=T_H2)
        text(s, x, y0 + 6, w, BB - risk_h - y0 - 14, c, size=T_BODY)
    y = BB - risk_h
    band(s, BX, y, BW, 20, "課題・リスク", fill=DARK, align=PP_ALIGN.LEFT, size=T_BODY)
    hdrs = [("#", 24), ("内容", 330), ("影響", 220), ("対応", 220), ("担当 / 期限", BW - 24 - 330 - 220 - 220)]
    xx = BX
    for lab, w in hdrs:
        text(s, xx, y + 22, w - 4, 14, lab, size=T_NOTE, bold=True, color=DARK, margins=(2, 0, 2, 0), align=PP_ALIGN.CENTER if lab == "#" else PP_ALIGN.LEFT)
        xx += w
    line(s, BX, y + 37, BX + BW, y + 37, color=LIGHT)
    rows = [("1", "保留 SKU 180 点の判断が営業と商品企画で平行線", "第 2 弾の候補選定が 2 週間遅れる", "委員会で決定(決定依頼 1)", "商品企画部長 / 9/9"),
            ("2", "統合便の開始に向けた配車計画で、上位 50 社の納期要件が 3 社分未確認", "10 月の統合便で納期遅延が起きる可能性", "9/12 までに 3 社へ確認、未確認なら個別配送を継続", "SCM 部 / 9/12"),
            ("3", "承認制の超過申請で本部審査の回答が 2 営業日を超えた件が 8 月に 11 件", "営業の不満と申請回避(口頭値引き)の再発", "審査担当を 2 名増員、条件付き承認の基準を明文化", "営業本部 / 9/19")]
    yy = y + 40
    for (yr, rh), r in zip(rows_fill(yy, BB, 3, 2), rows):
        xx = BX
        for (lab, w), v in zip(hdrs, r):
            text(s, xx, yr, w - 4, rh, v, size=T_SMALL, bold=(lab == "#"), anchor=MSO_ANCHOR.MIDDLE, margins=(2, 0, 2, 0), align=PP_ALIGN.CENTER if lab == "#" else PP_ALIGN.LEFT)
            xx += w
        line(s, BX, yr + rh + 1, BX + BW, yr + rh + 1, color=PALE)
    source(s, "出典: PMO 週次報告(第 6 週、2026 年 9 月 5 日時点)。● 順調 ● 要注意 ● 遅延")

# =====================================================================
@slide
def s20_survey():
    s = new_slide("営業担当 96 名の 7 割が「値引きの基準がない」と答え、承認制の導入には 8 割が賛成。懸念は回答速度に集中する",
                  "2026 年 7 月に全営業担当 108 名へ実施し 96 名が回答(回答率 89%)。5 段階の設問 5 問と自由記述で、承認制の設計に反映した")
    cw = 540
    band(s, BX, BY, cw, 22, "設問別の回答分布(n=96、%)", align=PP_ALIGN.LEFT)
    cd = CategoryChartData()
    cd.categories = ["値引きの基準が明確である", "案件の粗利を把握して値引きを決めている", "上位顧客の値引きは妥当だと思う", "本部承認制の導入に賛成", "承認に 2 営業日かかるなら困る"]
    for n_, v in (("そう思う", (6, 12, 21, 48, 39)), ("ややそう思う", (18, 25, 30, 33, 28)), ("あまり思わない", (41, 38, 31, 12, 22)), ("思わない", (35, 25, 18, 7, 11))):
        cd.add_series(n_, v)
    ch = s.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED_100, pt(BX), pt(BY + 26), pt(cw), pt(BB - BY - 26), cd).chart
    style_chart(ch, size=8.5, legend=XL_LEGEND_POSITION.BOTTOM)
    ch.value_axis.visible = False; ch.plots[0].gap_width = 60; ch.plots[0].overlap = 100
    ch.category_axis.reverse_order = True
    for ser, col, tc in zip(ch.series, [DARK, MID, LIGHT, PALE], [WHITE, WHITE, INK, INK]):
        ser.format.fill.solid(); ser.format.fill.fore_color.rgb = col; ser.format.line.color.rgb = WHITE; ser.format.line.width = Pt(0.75)
        dl = ser.data_labels; dl.show_value = True; dl.font.size = Pt(8.5); dl.font.color.rgb = tc; dl.number_format = '0'; dl.number_format_is_linked = False
    rx = BX + cw + 24; rw = BX + BW - rx
    band(s, rx, BY, rw, 22, "主な発見", align=PP_ALIGN.LEFT)
    text(s, rx, BY + 28, rw, 190, bullets([
        "#基準の不在は現場も認めている",
        "「基準が明確」は 24%。自由記述でも「顧客ごとに違う」「上司に聞くしかない」が 41 件",
        "#粗利を見ずに値引きしている",
        "「粗利を把握」は 37%。案件粗利の表示は承認制と同時に必須",
        "#承認制は歓迎、速度が条件",
        "賛成 81%。一方「2 営業日で困る」が 67%。自動承認 85% と本部審査 1 営業日を設計の前提にした",
    ]), size=T_BODY)
    ty = BY + 236
    text(s, rx, ty, rw, 16, "回答者の属性", size=T_H2, bold=True, margins=(0, 0, 0, 0))
    line(s, rx, ty + 18, rx + rw, ty + 18, color=INK, w=0.75)
    rows = [("", "人数", "比率"), ("A ランク担当", "22", "23%"), ("B ランク担当", "41", "43%"), ("C ランク担当", "33", "34%"), ("勤続 5 年未満", "37", "39%")]
    cw_ = [rw - 110, 55, 55]; yy = ty + 20
    for i, r in enumerate(rows):
        xx = rx
        for j, v in enumerate(r):
            text(s, xx, yy, cw_[j], 17, v, size=T_SMALL, bold=(i == 0), color=DARK if i == 0 else INK, align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT, margins=(0, 0, 0, 0), anchor=MSO_ANCHOR.MIDDLE)
            xx += cw_[j]
        line(s, rx, yy + 18, rx + rw, yy + 18, color=PALE); yy += 19
    source(s, "出典: 営業担当アンケート(2026 年 7 月 6〜17 日、n=96 / 108)。自由記述 214 件はテーマ別に分類")

# =====================================================================
@slide
def s21_issue_list():
    s = new_slide("残る論点は 6 つで、上位 3 つ(価格転嫁・拠点の顧客影響・SKU の代替率)は 10 月の委員会までに結論を出す",
                  "論点ごとに現状の理解・仮説・必要な分析・担当・期限を一覧にした。優先度は利益への影響と判断の期限で決めた")
    hdr = [("優先", 34), ("論点", 170), ("現状の理解", 210), ("仮説", 200), ("必要な分析・データ", 176), ("担当 / 期限", BW - 34 - 170 - 210 - 200 - 176)]
    x = BX
    for lab, w in hdr:
        band(s, x, BY, w - 4, 22, lab, align=PP_ALIGN.CENTER if lab == "優先" else PP_ALIGN.LEFT, size=T_BODY); x += w
    rows = [("高", "原材料高を価格転嫁できるか", "原材料 +12% に対し価格改定 +3%。上位 20 社とは年次交渉のみ", "上位 20 社の半数は 3〜5% の改定を受け入れる(粗利率 +0.9pt)", "顧客別の原価上昇額、競合の価格改定状況、交渉履歴", "営業本部長 / 10/10"),
            ("高", "3 拠点で翌日配送を維持できる範囲", "150km 圏 + 上位 50 社で 98.2%(シミュレーション)", "残り 1.8% は翌々日で許容される、または個別配送で吸収できる", "納期要件の未確認 3 社への確認、個別配送の原価", "SCM 部長 / 9/12"),
            ("高", "廃番 SKU の代替率", "第 1 弾 720 点は代替品で 84% を吸収。保留 180 点は未確認", "保留分も代替率 8 割で、失う売上は 0.6 億円以下", "保留 180 点の顧客別売上、代替品の受け入れ意向", "商品企画部長 / 9/9"),
            ("中", "承認制の例外枠の妥当性", "特注品 120 件/年と新規顧客の初回取引は営業部長判断", "例外は対象の 10% 未満に収まる", "例外申請の件数と粗利(8 月実績)", "PMO / 9/30"),
            ("中", "段取り短縮の目標値", "1 回 42 分。標準手順なし。外段取り化の余地は未計測", "外段取り化で 28 分は達成できる(B 工場の実績 26 分)", "工程別の段取り時間の実測(2 週間)", "生産技術 / 10/31"),
            ("低", "医療機器向けの初年度受注", "見込み客 30 社をリスト化。認証は翌 6 月", "初年度 3 億円は達成できる(競合 2 社の参入初年度 2.5〜4 億円)", "見込み客の需要規模、認証スケジュール", "新規事業室 / 12/20")]
    for (y, rh), r in zip(rows_fill(BY + 28, BB, 6, 4), rows):
        x = BX
        for (lab, w), v in zip(hdr, r):
            if lab == "優先":
                col = {"高": ACCENT, "中": DARK, "低": MID}[v]
                shape_text(rect(s, x + 3, y + rh / 2 - 11, 24, 22, fill=col), v, T_SMALL, True, WHITE)
            else:
                text(s, x, y, w - 4, rh, v, size=T_SMALL + 0.5, bold=(lab == "論点"), anchor=MSO_ANCHOR.MIDDLE, margins=(4, 1, 4, 1))
            x += w
        line(s, BX + 34, y + rh + 2, BX + BW, y + rh + 2, color=LIGHT)
    source(s, "出典: 論点管理表(2026 年 9 月 5 日時点)。優先度は利益への影響(pt)× 判断期限で 3 段階")

# =====================================================================
@slide
def s22_case_study():
    s = new_slide("事例: 同業 B 社(売上 620 億円)は値引き承認制と拠点集約で 2 年で営業利益率を 3.1% から 7.8% に改善した",
                  "B 社は X 社と顧客構成が近く(上位 20 社で売上の 58%)、承認制の設計と拠点集約の順序をそのまま参考にできる。数値は公開資料と業界紙から整理")
    pw = 190
    rect(s, BX, BY, pw, BH, fill=PALE)
    text(s, BX + 10, BY + 8, pw - 20, 20, "B 社の概要", size=T_H1, bold=True, margins=(0, 0, 0, 0))
    line(s, BX + 10, BY + 30, BX + pw - 10, BY + 30, color=INK, w=1)
    prof = [("業種", "産業用部品の製造"), ("売上", "620 億円(FY25)"), ("従業員", "1,840 名"), ("拠点", "工場 4、物流 7 → 3"), ("顧客", "上位 20 社で 58%"), ("実施期間", "2023 年 4 月〜2025 年 3 月"), ("体制", "社長直轄 PMO 4 名 + 外部")]
    for (yy, rh), (k, v) in zip(rows_fill(BY + 38, BB - 70, len(prof), 2), prof):
        text(s, BX + 10, yy, 56, rh, k, size=T_SMALL + 0.5, color=DARK, margins=(0, 0, 0, 0), anchor=MSO_ANCHOR.MIDDLE)
        text(s, BX + 66, yy, pw - 76, rh, v, size=T_BODY, margins=(0, 0, 0, 0), anchor=MSO_ANCHOR.MIDDLE)
        line(s, BX + 10, yy + rh + 1, BX + pw - 10, yy + rh + 1, color=LIGHT)
    rect(s, BX + 10, BB - 62, pw - 20, 54, fill=WHITE)
    text(s, BX + 10, BB - 62, pw - 20, 54, [("営業利益率", {"size": T_SMALL, "color": DARK, "align": PP_ALIGN.CENTER}), ("3.1% → 7.8%", {"size": 16, "bold": True, "color": ACCENT, "align": PP_ALIGN.CENTER})], anchor=MSO_ANCHOR.MIDDLE)
    x0 = BX + pw + 16; cols = cols_fill(x0, BX + BW - x0, 3, 26)
    heads = ["課題", "アプローチ", "効果"]
    contents = [
        bullets(["#値引きの野放し", "平均値引き率 8.9%、上位顧客で 12% 超", "承認プロセスがなく、事後の集計のみ", "#物流拠点 7 か所の過剰", "稼働率 52%、小口配送が 45%", "拠点間の横持ちが月 410 便", "#SKU 4,200 点", "下位 40% は売上の 4%、廃番基準なし", "#評価は売上高のみ", "粗利・値引きは評価に反映されず"], size=11),
        bullets(["#承認制を 3 か月で導入", "5% 以下は自動承認、超過は本部が 1 営業日", "顧客ランク別上限(10/6/3%)", "申請時に案件粗利を表示", "#拠点を 7 → 3 に 14 か月で集約", "上位 40 社 + 120km 圏で翌日配送を維持", "並行稼働 5 か月、在庫 +1.1 億円", "#SKU を 3 段階で 2,500 点に", "代替品対応表と顧客への事前照会", "#2 年目に評価へ粗利を追加"], size=11),
        bullets(["#値引き率 8.9% → 4.8%(8 か月)", "粗利率 +3.4pt。超過申請は初月 38% → 3 か月で 14%", "#物流費率 10.1% → 8.6%", "年 9.3 億円削減、投資 4.2 億円、回収 6 か月", "離反は 3 社(売上 1.2 億円)", "#SKU 集約で段取り −35%", "稼働率 +5pt、在庫回転 +0.8 回", "#営業利益率 3.1% → 7.8%", "2 年で +4.7pt。従業員満足度も +6pt"], size=11),
    ]
    for k, ((x, w), h, c) in enumerate(zip(cols, heads, contents)):
        y0 = panel_head(s, x, BY, w, h)
        text(s, x, y0 + 6, w, BB - y0 - 6 - 40, c, size=11)
        if k < 2:
            rect(s, x + w + 6, BY + 4, 14, 16, fill=DARK, shape=MSO_SHAPE.RIGHT_ARROW)
    rect(s, x0, BB - 34, BX + BW - x0, 34, fill=PALE)
    text(s, x0 + 8, BB - 34, BX + BW - x0 - 16, 34, "**X 社への示唆:** 承認制は「自動承認の比率」と「本部の回答速度」で現場の受容が決まる。拠点集約は投資判断を先に置き、並行稼働の在庫増を織り込む", size=T_BODY, anchor=MSO_ANCHOR.MIDDLE)
    source(s, "出典: B 社 有価証券報告書(FY2023〜FY2025)、決算説明資料、業界紙記事(2025 年 6 月)。数値は公開情報から整理した推定を含む")

# =====================================================================
@slide
def s23_project_approach():
    s = new_slide("本プロジェクトは「診断 → 設計 → 実行支援」の 3 フェーズ・9 か月で進め、フェーズ 1 の終わりに投資判断の材料を揃える",
                  "各フェーズの目的・主な活動・成果物・期間を揃えて示す。フェーズ 2 の途中から第 1 波の施策を先行して開始し、効果の発現を前倒しする")
    rw_ = 90; n = 3; gap = 8; x0 = BX + rw_ + 6; cw = (BX + BW - x0 - gap * (n - 1)) / n
    phases = [("フェーズ 1: 診断", "9〜10 月(6 週)"), ("フェーズ 2: 設計", "11〜1 月(12 週)"), ("フェーズ 3: 実行支援", "2〜5 月(16 週)")]
    for i, (h, t) in enumerate(phases):
        x = x0 + i * (cw + gap)
        sp = rect(s, x, BY, cw, 40, fill=DARK, shape=MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON)
        shape_text(sp, [(h, {"size": T_H1, "bold": True, "color": WHITE}), (t, {"size": T_SMALL, "color": WHITE})], T_H1, True, WHITE, margins=(14, 0, 8, 0))
    rowsdef = [("目的", ["利益低下の要因を定量化し、打ち手の候補と投資判断の材料を揃える", "3 施策の詳細を設計し、承認制と廃番基準を全社で合意する", "第 1 波を定着させ、拠点集約の並行稼働を安全に完了する"]),
               ("主な活動", [["原価差異分析(FY23〜25)", "営業日報 12,400 件の値引き分析", "拠点別原価と配送シミュレーション", "部門ヒアリング 24 名、顧客 31 社"],
                             ["承認フロー・ランク別上限の設計", "廃番基準と代替品対応表の作成", "3 拠点案の詳細設計と投資判断資料", "KPI と月次報告の仕組みを設計"],
                             ["承認制の全社展開と定着支援(常駐)", "SKU 廃番 第 1・2 弾の実行", "拠点移転の並行稼働の管理", "効果測定と委員会への月次報告"]]),
               ("成果物", [["現状分析報告書", "施策候補と効果試算", "投資判断資料(案)"], ["承認制 運用マニュアル", "廃番基準書・代替品対応表", "拠点集約 実行計画書"], ["月次効果レポート(8 回)", "定着度評価と改善提案", "最終報告書"]]),
               ("体制", ["コンサル 4 名 + 経営企画 2 名", "コンサル 4 名 + 分科会 3 つ(19 名)", "コンサル 3 名(うち 2 名常駐)+ PMO"])]
    for (y, rh), (lab, cells) in zip(rows_fill(BY + 48, BB, 4, 6), rowsdef):
        vchevron(s, BX, y, rw_, rh, lab, size=T_BODY)
        for i, c in enumerate(cells):
            x = x0 + i * (cw + gap)
            if isinstance(c, list):
                text(s, x, y, cw, rh, bullets(c), size=T_BODY, anchor=MSO_ANCHOR.MIDDLE)
            else:
                text(s, x, y, cw, rh, c, size=T_BODY, anchor=MSO_ANCHOR.MIDDLE, margins=(6, 2, 6, 2))
        line(s, x0, y + rh + 3, BX + BW, y + rh + 3, color=LIGHT, w=0.75)
    source(s, "出典: プロジェクト提案書(2026 年 8 月版)。期間は経営会議の承認を 9 月第 1 週に得た場合")

# =====================================================================
@slide
def s24_deliverables():
    s = new_slide("成果物は 4 点で、フェーズ 1 の「現状分析報告書」と「投資判断資料」を 10 月の経営会議に提出する",
                  "各成果物の内容・想定ページ数・提出時期を示す。分析の中間結果は週次で共有し、最終版で初めて見る事態を避ける")
    items = [("現状分析報告書", "10 月 10 日", "約 40 ページ", ["利益低下の要因分解(pt)", "部門別の課題と原因", "同業比較と取り戻せる差"], "chart"),
             ("施策候補と投資判断資料", "10 月 24 日", "約 25 ページ", ["12 施策の効果・投資・難易度", "拠点集約 3 案の比較", "実行の波と期待効果"], "matrix"),
             ("承認制 運用マニュアル / 廃番基準書", "翌 1 月 16 日", "約 60 ページ", ["業務フロー、上限、例外規定", "システム要件と画面", "廃番基準・代替品対応表"], "flow"),
             ("実行計画書と月次効果レポート", "翌 1 月 30 日 / 月次", "約 30 ページ + 月次 8 ページ", ["ガント・体制・KPI", "月次の効果と計画差", "課題・リスクと決定依頼"], "gantt")]
    for (x, w), (name, due, pages, bl, kind) in zip(cols_fill(BX, BW, 4, 16), items):
        fh = 128
        fx, fy = x + (w - 160) / 2, BY
        rect(s, fx, fy, 160, fh, fill=WHITE, line=LIGHT, lw=1)
        rect(s, fx + 10, fy + 10, 100, 4, fill=INK); rect(s, fx + 10, fy + 18, 130, 3, fill=LIGHT)
        if kind == "chart":
            for k, hh in enumerate([40, 52, 30, 66, 58]):
                rect(s, fx + 14 + k * 26, fy + 108 - hh, 16, hh, fill=DARK if k == 3 else MID)
        elif kind == "matrix":
            for r_ in range(4):
                for c_ in range(3):
                    rect(s, fx + 12 + c_ * 46, fy + 30 + r_ * 20, 42, 16, fill=DARK if r_ == 0 else PALE)
        elif kind == "flow":
            for k in range(4):
                rect(s, fx + 10 + k * 36, fy + 40, 30, 22, fill=PALE, line=DARK, lw=0.5)
                if k < 3:
                    line(s, fx + 40 + k * 36, fy + 51, fx + 46 + k * 36, fy + 51, color=DARK, w=0.75, arrow=True)
            for k in range(3):
                rect(s, fx + 10, fy + 72 + k * 12, 130, 4, fill=LIGHT)
        else:
            for k, (a, b) in enumerate([(0, 40), (20, 60), (50, 50), (70, 60), (90, 40)]):
                rect(s, fx + 12 + a, fy + 32 + k * 16, b, 8, fill=DARK if k in (0, 3) else MID, shape=MSO_SHAPE.PENTAGON)
        text(s, x, BY + fh + 10, w, 22, name, size=T_H2, bold=True, align=PP_ALIGN.CENTER, margins=(0, 0, 0, 0))
        line(s, x, BY + fh + 34, x + w, BY + fh + 34, color=INK, w=1.5)
        text(s, x, BY + fh + 40, w, 30, [(f"提出: {due}", {"size": T_SMALL, "color": DARK}), (f"分量: {pages}", {"size": T_SMALL, "color": DARK})], margins=(0, 0, 0, 0))
        text(s, x, BY + fh + 74, w, BB - BY - fh - 74 - 90, bullets(bl, size=11), size=11)
    # 提出スケジュール(時間軸)
    ty = BB - 80
    band(s, BX, ty, BW, 20, "提出スケジュールと承認の場", align=PP_ALIGN.LEFT, size=T_BODY)
    ly_ = ty + 50
    tx0, tx1 = BX + 60, BX + BW - 60
    line(s, tx0, ly_, tx1, ly_, color=DARK, w=1.25)
    marks = [("9 月", 0.0, "キックオフ", None), ("10/10", 0.14, "現状分析報告書", "経営会議"), ("10/24", 0.28, "投資判断資料", "推進委員会"), ("翌 1/16", 0.52, "運用マニュアル・廃番基準書", "分科会 → 委員会"), ("翌 1/30", 0.66, "実行計画書", "推進委員会"), ("2〜5 月", 0.84, "月次効果レポート", "毎月の委員会"), ("翌 5 月", 1.0, "最終報告", "経営会議")]
    for lab, f, what, where in marks:
        cx = tx0 + f * (tx1 - tx0)
        rect(s, cx - 5, ly_ - 5, 10, 10, fill=DARK if where else MID, shape=MSO_SHAPE.DIAMOND)
        text(s, cx - 50, ly_ - 24, 100, 14, lab, size=T_SMALL, bold=True, align=PP_ALIGN.CENTER, margins=(0, 0, 0, 0))
        text(s, cx - 56, ly_ + 6, 112, 24, [(what, {"size": 8.5}), (where or "", {"size": 8, "color": MUTED})], align=PP_ALIGN.CENTER, margins=(0, 0, 0, 0))
    source(s, "出典: プロジェクト提案書(2026 年 8 月版)。ページ数は同規模案件の実績からの目安")

# =====================================================================
@slide
def s25_goals():
    s = new_slide("プロジェクトの目的は「3 年で営業利益率を 2.1% から 6.5% へ戻す」ことであり、ゴールを 3 つの数値で定義する",
                  "背景にある利益低下の 7 割は自社要因で、仕組みを変えれば戻せる。ゴールは FY27 通期の営業利益率と、それを支える 2 つの中間指標で置く")
    lw_ = 300; mw = 200; gap = 20
    y0 = panel_head(s, BX, BY, lw_, "背景")
    text(s, BX, y0 + 8, lw_, BB - y0 - 8, bullets([
        "#利益が 3 年で半減した",
        "営業利益率 5.9%(FY22)→ 2.1%(FY25)。粗利率は 6.1pt 低下",
        "#7 割は自社の運営要因",
        "値引き 2.7pt、物流 1.5pt、SKU 0.6pt は意思決定で戻せる。原材料高 1.8pt は同業共通",
        "#過去 2 回の改善は途中で止まった",
        "効果を測らず、例外が増え、決定が先送りされた。今回は委員会の権限と月次の効果測定を最初に置く",
    ]), size=T_BODY)
    mx = BX + lw_ + gap
    rect(s, mx, BY, mw, BH, fill=DARK)
    text(s, mx + 12, BY + 16, mw - 24, 40, "目的", size=T_H1, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    text(s, mx + 12, BY + 60, mw - 24, 120, "3 年で営業利益率を **2.1% → 6.5%** に戻し、原材料高に耐える収益構造にする", size=14, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    line(s, mx + 30, BY + 190, mx + mw - 30, BY + 190, color=WHITE, w=0.75)
    text(s, mx + 12, BY + 200, mw - 24, BH - 210, [(t, {"level": 0, "color": WHITE, "size": T_SMALL + 0.5}) for t in ["投資 4.8 億円、回収 14 か月", "第 1 波の効果を FY26 下期に発現", "委員会が月次で投資と中止を判断"]], color=WHITE)
    rect(s, mx + mw + 4, BY + BH / 2 - 10, 14, 20, fill=DARK, shape=MSO_SHAPE.RIGHT_ARROW)
    gx = mx + mw + gap + 6; gw = BX + BW - gx
    y0 = panel_head(s, gx, BY, gw, "ゴール(FY27 通期)")
    goals = [("営業利益率", "6.5%", "FY25 2.1% → +4.4pt(+20.6 億円)", "最終ゴール"), ("平均値引き率", "5.0%", "FY25 9.7%。承認制で 8 か月以内に到達", "中間指標 1"), ("物流費率", "8.3%", "FY25 9.4%。3 拠点集約の完了後に到達", "中間指標 2")]
    for (y, rh), (k, v, d, tag) in zip(rows_fill(y0 + 10, BB, 3, 8), goals):
        rect(s, gx, y, gw, rh, fill=PALE)
        text(s, gx + 10, y + 4, 90, 14, tag, size=8, color=MUTED, margins=(0, 0, 0, 0))
        text(s, gx + 10, y + 16, 110, rh - 20, k, size=T_H2, bold=True, anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0))
        text(s, gx + 120, y, 90, rh, v, size=T_KPI, bold=True, color=ACCENT, anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0))
        text(s, gx + 214, y, gw - 224, rh, d, size=T_SMALL + 0.5, color=DARK, anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0))
    source(s, "出典: プロジェクト提案書(2026 年 8 月版)、X 社 財務データ(FY2022〜FY2025)")

# ---------------- 実行 ----------------
only = [int(x) for x in A.only.split(",") if x.strip()]
out_path = A.out
if only and A.out == ap.get_default("out"):
    # --only は指定したスライドだけの部分デッキになる。--out を明示しない事故防止に別名で保存する
    root, ext = os.path.splitext(A.out)
    out_path = f"{root}-partial{ext}"
    print(f"--only 指定のため出力先を {out_path} に変更(正本の {A.out} を誤って上書きしないため)")
for i, fn in enumerate(SLIDES, start=1):
    if only and i not in only:
        continue
    fn()
os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
prs.save(out_path)
A.out = out_path
print(f"saved {A.out} ({page[0]} slides)")
