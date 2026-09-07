# -*- coding: utf-8 -*-
"""
ゴールデンスライド v1 を、ユーザーのマスタ(プロジェクトアプローチ.pptx)の上に python-pptx で直接組む。
エンジン(layout.js)の制約は一切使わない。「一流コンサルの最終資料」の密度・構成・配置を、
人が作るのと同じ自由度で作り、これを品質の正解(比較の基準)にする。

  python scripts/make-golden.py [--master <path>] [--out docs/golden/golden-v1.pptx]

題材は架空の中堅部品製造業 X 社(売上 480 億円)の収益改善。数値はすべて評価用の架空データ。
"""
import argparse, copy, os, sys
sys.stdout.reconfigure(encoding="utf-8")
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Pt, Emu
from lxml import etree

ap = argparse.ArgumentParser()
ap.add_argument("--master", default=r"C:\Users\nakaj\Documents\プロジェクトアプローチ.pptx")
ap.add_argument("--out", default="docs/golden/golden-v1.pptx")
A = ap.parse_args()

# ---------------- デザイントークン(マスタに合わせる) ----------------
ACCENT = RGBColor(0xD0, 0x4A, 0x02)   # マスタの見出し帯の橙
INK = RGBColor(0x1A, 0x1A, 0x1A)
DARK = RGBColor(0x40, 0x40, 0x40)
MUTED = RGBColor(0x6E, 0x6E, 0x6E)
MID = RGBColor(0xA6, 0xA6, 0xA6)
LIGHT = RGBColor(0xD9, 0xD9, 0xD9)
PALE = RGBColor(0xF2, 0xF2, 0xF2)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Arial"
LAYOUT_NAME = "Title and Full Content - Subtitle"
# 版面(本文域)
BX, BY, BW, BH = 32, 128, 896, 364   # y 128〜492。出典は 492〜504
SRC_Y = 494

def pt(v):
    return Emu(int(round(v * 12700)))

# ---------------- 低レベル描画ヘルパ ----------------
def _set_font(run, size, bold=False, color=INK, italic=False, name=FONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = name
    run.font.color.rgb = color

def _bullet(paragraph, level=0, char="•"):
    """本物の箇条書き(ぶら下げインデント)を XML で付ける"""
    pPr = paragraph._p.get_or_add_pPr()
    marL = 170000 + level * 200000
    pPr.set("marL", str(marL))
    pPr.set("indent", str(-170000))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    bu = etree.SubElement(pPr, qn("a:buChar"))
    bu.set("char", char)

def _nobullet(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", "0")
    pPr.set("indent", "0")
    etree.SubElement(pPr, qn("a:buNone"))

def text(slide, x, y, w, h, content, size=12, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         margins=(4, 2, 4, 2), line_spacing=1.15, name=FONT):
    """content: str | list[str | (str, {opts})]。文字列中の **…** は太字。"""
    tb = slide.shapes.add_textbox(pt(x), pt(y), pt(w), pt(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [pt(m) for m in margins]
    items = content if isinstance(content, list) else [content]
    first = True
    for it in items:
        opts = {}
        if isinstance(it, tuple):
            it, opts = it
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = opts.get("align", align)
        p.line_spacing = line_spacing
        if opts.get("space_before"):
            p.space_before = Pt(opts["space_before"])
        lvl = opts.get("level")
        if lvl is not None:
            _bullet(p, lvl, opts.get("char", "•" if lvl == 0 else "–"))
        else:
            _nobullet(p)
        _rich(p, it, size=opts.get("size", size), bold=opts.get("bold", bold), color=opts.get("color", color), name=name)
    return tb

def _rich(p, s, size, bold, color, name=FONT):
    parts = s.split("**")
    for i, seg in enumerate(parts):
        if not seg:
            continue
        r = p.add_run()
        r.text = seg
        _set_font(r, size, bold=bold or (i % 2 == 1), color=color, name=name)

def rect(slide, x, y, w, h, fill=None, line=None, line_w=0.75, shape=MSO_SHAPE.RECTANGLE, name=None):
    s = slide.shapes.add_shape(shape, pt(x), pt(y), pt(w), pt(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(line_w)
    s.shadow.inherit = False
    if s.has_text_frame:
        s.text_frame.text = ""
    if name:
        s.name = name
    return s

def line(slide, x1, y1, x2, y2, color=LIGHT, w=0.75, dash=False):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, pt(x1), pt(y1), pt(x2), pt(y2))
    c.line.color.rgb = color
    c.line.width = Pt(w)
    if dash:
        ln = c.line._get_or_add_ln()
        prst = etree.SubElement(ln, qn("a:prstDash"))
        prst.set("val", "dash")
    return c

def band(slide, x, y, w, h, label, fill=ACCENT, color=WHITE, size=11, align=PP_ALIGN.CENTER, bold=True):
    """列見出しの帯(マスタの橙帯)"""
    s = rect(slide, x, y, w, h, fill=fill)
    tf = s.text_frame
    tf.margin_left = tf.margin_right = pt(4)
    tf.margin_top = tf.margin_bottom = pt(1)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = align
    _rich(p, label, size, bold, color)
    return s

def rowhead(slide, x, y, w, h, label, chevron=False, size=11, fill=LIGHT):
    """行見出し: 灰色の箱。順序があるときはマスタと同じ下向きの五角形"""
    if chevron:
        s = rect(slide, x, y, w, h, fill=fill, shape=MSO_SHAPE.PENTAGON)
        s.rotation = 90
        # 回転後に同じ枠に収まるよう幅と高さを入れ替える
        s.left, s.top, s.width, s.height = pt(x + (w - h) / 2), pt(y + (h - w) / 2), pt(h), pt(w)
        tb = text(slide, x, y, w, h, label, size=size, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        return s, tb
    s = rect(slide, x, y, w, h, fill=fill)
    tf = s.text_frame
    tf.margin_left = tf.margin_right = pt(6)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    _rich(p, label, size, True, INK)
    return s, None

def harvey(slide, cx, cy, d, level):
    """ハーベイボール(0〜4)。外周の円 + 塗りの扇形"""
    x, y = cx - d / 2, cy - d / 2
    if level >= 4:
        rect(slide, x, y, d, d, fill=DARK, line=DARK, shape=MSO_SHAPE.OVAL)
        return
    rect(slide, x, y, d, d, fill=WHITE, line=DARK, line_w=1, shape=MSO_SHAPE.OVAL)
    if level <= 0:
        return
    s = slide.shapes.add_shape(MSO_SHAPE.PIE, pt(x), pt(y), pt(d), pt(d))
    s.fill.solid(); s.fill.fore_color.rgb = DARK
    s.line.fill.background()
    # 角度は 60000 分の 1 度。python-pptx は /100000 で正規化する。12 時(270°)から時計回り
    start, end = 270.0, (270.0 + 90.0 * level) % 360
    s.adjustments[0] = start * 60000 / 100000
    s.adjustments[1] = end * 60000 / 100000
    s.shadow.inherit = False

def diamond(slide, cx, cy, d, fill=DARK):
    s = rect(slide, cx - d / 2, cy - d / 2, d, d, fill=fill, shape=MSO_SHAPE.DIAMOND)
    return s

def source(slide, s):
    text(slide, BX, SRC_Y, BW, 12, s, size=8, color=MUTED, margins=(0, 0, 0, 0))

# ---------------- スライドの骨組み ----------------
prs = Presentation(A.master)
layout = next(l for l in prs.slide_layouts if l.name == LAYOUT_NAME)
# 既存スライドを全部消す(マスタとレイアウトだけ使う)
for sldId in list(prs.slides._sldIdLst):
    prs.part.drop_rel(sldId.rId)
    prs.slides._sldIdLst.remove(sldId)

page = [0]

def new_slide(title, lead, title_size=22, lead_size=13):
    page[0] += 1
    s = prs.slides.add_slide(layout)
    # 本文プレースホルダは使わない(空のまま残すと「テキストを入力」が出る)
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
    sub.text_frame.text = lead
    sub.text_frame.word_wrap = True
    for p in sub.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(lead_size)
            r.font.color.rgb = DARK
            r.font.name = FONT
    # ページ番号(マスタの SlideId 相当)
    text(s, 919, 513, 9, 9, str(page[0]), size=7.5, color=INK, margins=(0, 0, 0, 0), align=PP_ALIGN.RIGHT)
    return s

# ---------------- グラフ共通 ----------------
def style_chart(chart, size=9, legend=None, gridlines=False):
    chart.font.size = Pt(size)
    chart.font.name = FONT
    chart.has_title = False
    if legend:
        chart.has_legend = True
        chart.legend.position = legend
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(size)
    else:
        chart.has_legend = False
    try:
        va = chart.value_axis
        va.has_major_gridlines = gridlines
        if gridlines:
            va.major_gridlines.format.line.color.rgb = PALE
        va.format.line.fill.background()
        va.tick_labels.font.size = Pt(size)
        va.tick_labels.font.color.rgb = MUTED
        ca = chart.category_axis
        ca.format.line.color.rgb = LIGHT
        ca.tick_labels.font.size = Pt(size)
        ca.tick_labels.font.color.rgb = INK
        ca.has_major_gridlines = False
    except Exception:
        pass

# =====================================================================
# 1. エグゼクティブサマリー: 論点 × 主張 × 根拠 × インパクト
# =====================================================================
s = new_slide(
    "収益悪化の主因は値引きと物流費であり、価格統制と拠点集約で営業利益率を 2.1% から 6.5% に戻せる",
    "粗利率は 3 年で 6.1pt 低下したが、市場要因(原材料高)は 1.8pt にとどまる。残り 4.3pt は自社の運営に起因し、投資 4.8 億円・回収 14 か月で是正できる",
    title_size=20)
cols = [("#", 40), ("主張", 250), ("根拠となる事実", 396), ("利益インパクト(FY27)", 210)]
x = BX
for lab, w in cols:
    band(s, x, BY, w - 4, 22, lab, align=PP_ALIGN.LEFT if lab != "#" else PP_ALIGN.CENTER)
    x += w
rows = [
    ("1", "**値引きを本部承認制にし、平均値引き率を 9.7% → 5% 以下へ**",
     ["平均値引き率は 4.8%(FY23)→ 9.7%(FY25)に拡大。粗利率低下 6.1pt のうち **2.7pt** を占める",
      "値引き総額 15.6 億円の 73% が上位 20 顧客に集中。承認制の対象を絞れば統制できる",
      "同業 B 社は承認制導入後 8 か月で値引き率を 4.1pt 圧縮(公開資料)"],
     "+2.4pt\n(+11.5 億円)"),
    ("2", "**6 拠点を 3 拠点に集約し、共同配送で物流費率 9.4% → 8.3% へ**",
     ["拠点稼働率は平均 58%、小口配送(1 便 20 万円未満)が件数の 41%",
      "翌日配送は上位 50 社 + 半径 150km で維持でき、顧客離反リスクは限定的(31 社ヒアリング)",
      "投資 3.1 億円、年間削減 5.3 億円、回収 7 か月"],
     "+1.1pt\n(+5.3 億円)"),
    ("3", "**下位 40% の 1,540 SKU を 2 段階で廃番し、段取り時間を 33% 削減**",
     ["3,860 SKU のうち下位 40% は売上の 3.2%、粗利率 15% 未満",
      "段取り替え月 1,900 回・1 回 42 分。SKU 集約で月 1,270 回・28 分へ",
      "廃番基準(年間売上 300 万円未満かつ粗利率 15% 未満)を新設"],
     "+0.8pt\n(+3.8 億円)"),
]
y = BY + 26
rh = 96
for num, claim, facts, impact in rows:
    x = BX
    text(s, x, y, 36, rh, num, size=20, bold=True, color=ACCENT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP); x += 40
    text(s, x, y, 246, rh, claim, size=12, anchor=MSO_ANCHOR.TOP); x += 250
    text(s, x, y, 392, rh, [(f, {"level": 0}) for f in facts], size=10.5, anchor=MSO_ANCHOR.TOP); x += 396
    text(s, x, y, 206, rh, impact, size=14, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    line(s, BX, y + rh + 2, BX + BW, y + rh + 2, color=ACCENT, w=0.75)
    y += rh + 6
# 合計の帯
rect(s, BX, y + 2, BW, 30, fill=PALE)
text(s, BX + 8, y + 2, 600, 30, "3 施策の合計: 営業利益率 **2.1% → 6.5%**(FY27 通期)。投資 4.8 億円、回収 14 か月。市場要因(原材料高 1.8pt)は医療機器向け参入で 3 年かけて相殺する",
     size=11, anchor=MSO_ANCHOR.MIDDLE)
text(s, BX + BW - 200, y + 2, 192, 30, "+4.4pt(+20.6 億円)", size=14, bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, color=ACCENT)
source(s, "出典: X 社 財務データ(FY2023〜FY2025)、営業日報 12,400 件、拠点別原価集計、顧客ヒアリング 31 社。インパクトは FY2027 通期の営業利益への寄与")

# =====================================================================
# 2. 推移(ネイティブ折れ線) × 意味合い
# =====================================================================
s = new_slide(
    "粗利率は 3 年連続で低下し、FY25 に業界平均を 4.6pt 下回った。原材料高だけでは説明できない",
    "低下幅 6.1pt のうち原材料高は 1.8pt。残る 4.3pt は値引きと物流費という自社の運営要因で、同業他社は同じ原材料高の下で 30% 前後を維持している")
cw = 540
band(s, BX, BY, cw, 22, "粗利率の推移(%)", align=PP_ALIGN.LEFT)
cd = CategoryChartData()
cd.categories = ["FY21", "FY22", "FY23", "FY24", "FY25"]
cd.add_series("X 社", (32.4, 31.8, 31.2, 28.0, 25.1))
cd.add_series("業界平均(8 社)", (30.1, 30.4, 30.2, 29.9, 29.7))
gf = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, pt(BX), pt(BY + 26), pt(cw), pt(250), cd)
ch = gf.chart
style_chart(ch, size=9, legend=XL_LEGEND_POSITION.BOTTOM, gridlines=True)
ch.value_axis.minimum_scale = 20
ch.value_axis.maximum_scale = 35
ch.value_axis.major_unit = 5
for i, ser in enumerate(ch.series):
    col = ACCENT if i == 0 else MID
    ser.format.line.color.rgb = col
    ser.format.line.width = Pt(2.25 if i == 0 else 1.5)
    ser.marker.style = 8  # circle
    ser.marker.size = 6
    ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = col
    ser.marker.format.line.color.rgb = col
    ser.smooth = False
    dl = ser.data_labels
    dl.show_value = True
    dl.number_format = '0.0'
    dl.number_format_is_linked = False
    dl.font.size = Pt(9)
    dl.font.color.rgb = col
    dl.position = XL_LABEL_POSITION.ABOVE if i == 0 else XL_LABEL_POSITION.BELOW
    if i == 0:
        for k in (3, 4):
            ch.series[0].points[k].data_label.position = XL_LABEL_POSITION.BELOW
            ch.series[0].points[k].data_label.font.size = Pt(9)
            ch.series[0].points[k].data_label.font.color.rgb = ACCENT
    if i == 1:
        for k in (3, 4):
            ser.points[k].data_label.position = XL_LABEL_POSITION.ABOVE
            ser.points[k].data_label.font.size = Pt(9)
            ser.points[k].data_label.font.color.rgb = MID
# 注釈(FY24 の急落)
text(s, BX + 250, BY + 36, 230, 40, "FY24: 主要顧客 A 社との年次交渉で **7% の値引き**に応じ、単年で −3.2pt", size=9, color=DARK)
line(s, BX + 360, BY + 76, BX + 392, BY + 118, color=DARK, w=0.75)
# 右: 意味合い
rx = BX + cw + 24
rw = BW - cw - 24
band(s, rx, BY, rw, 22, "意味合い", align=PP_ALIGN.LEFT)
text(s, rx, BY + 30, rw, 200, [
    ("**低下の 7 割は自社要因**", {}),
    ("値引き率の拡大 2.7pt、物流費の増加 1.5pt、SKU 増による段取り 0.6pt", {"level": 0}),
    ("原材料高(特殊鋼 +12%)は 1.8pt で、同業 8 社も同じ条件", {"level": 0}),
    ("**業界平均との差 4.6pt は「取り戻せる差」**", {"space_before": 6}),
    ("同業 B 社は FY24 に承認制を導入し、値引き率を 4.1pt 圧縮した", {"level": 0}),
    ("差 4.6pt のうち 4.3pt は 3 施策で FY27 までに回収可能", {"level": 0}),
], size=11)
# 同業比較の小表
ty = BY + 166
text(s, rx, ty, rw, 16, "FY25 の同業比較", size=10, bold=True, margins=(0, 0, 0, 0))
line(s, rx, ty + 18, rx + rw, ty + 18, color=DARK, w=0.75)
cmp_rows = [("", "X 社", "同業平均", "差"), ("粗利率", "25.1%", "29.7%", "−4.6pt"), ("平均値引き率", "9.7%", "5.2%", "+4.5pt"), ("物流費率", "9.4%", "8.1%", "+1.3pt"), ("原材料費率", "41.8%", "41.2%", "+0.6pt")]
cw_ = [rw - 70 * 3, 70, 70, 70]
yy = ty + 20
for r_i, r in enumerate(cmp_rows):
    xx = rx
    for c_i, v in enumerate(r):
        text(s, xx, yy, cw_[c_i], 16, v, size=9.5, bold=(r_i == 0), color=DARK if r_i == 0 else INK, align=PP_ALIGN.LEFT if c_i == 0 else PP_ALIGN.RIGHT, margins=(0, 0, 0, 0), anchor=MSO_ANCHOR.MIDDLE)
        xx += cw_[c_i]
    line(s, rx, yy + 17, rx + rw, yy + 17, color=PALE)
    yy += 18
rect(s, rx, BY + 282, rw, 44, fill=PALE)
text(s, rx + 6, BY + 282, rw - 12, 44, "So what: 原材料高への対応(価格転嫁)よりも先に、値引きと物流の統制を打つべき", size=11, bold=True, anchor=MSO_ANCHOR.MIDDLE)
source(s, "出典: X 社 決算資料(FY2021〜FY2025)、業界平均は同業 8 社の公表値の単純平均。原材料高の影響は原価差異分析による")

# =====================================================================
# 3. ウォーターフォール(要因分解)
# =====================================================================
s = new_slide(
    "粗利率 6.1pt の低下は、値引き 2.7pt・原材料 1.8pt・物流 1.5pt・SKU 0.6pt に分解でき、7 割が運営要因",
    "FY23 → FY25 の原価差異を要因別に積み上げた。為替と数量は +0.5pt の押し上げで、自社要因が無ければ低下は 1.3pt にとどまっていた")
cw = 560
band(s, BX, BY, cw, 22, "粗利率の変化要因(FY23 → FY25、pt)", align=PP_ALIGN.LEFT)
cats = ["FY23", "値引き拡大", "原材料高", "物流費増", "SKU 段取り", "為替・数量", "FY25"]
vals = [31.2, -2.7, -1.8, -1.5, -0.6, 0.5, 25.1]
base, up, down, total = [], [], [], []
run = 0.0
for i, (c, v) in enumerate(zip(cats, vals)):
    if i == 0 or i == len(cats) - 1:
        base.append(0); total.append(v); up.append(0); down.append(0); run = v
    elif v < 0:
        base.append(run + v); down.append(-v); up.append(0); total.append(0); run += v
    else:
        base.append(run); up.append(v); down.append(0); total.append(0); run += v
cd = CategoryChartData()
cd.categories = cats
cd.add_series("base", base)
cd.add_series("total", total)
cd.add_series("減少", down)
cd.add_series("増加", up)
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, pt(BX), pt(BY + 26), pt(cw), pt(262), cd)
ch = gf.chart
style_chart(ch, size=9, gridlines=False)
ch.value_axis.minimum_scale = 20
ch.value_axis.maximum_scale = 34
ch.value_axis.visible = False
ch.plots[0].gap_width = 45
ch.plots[0].overlap = 100
fills = [None, DARK, ACCENT, MID]
for ser, col in zip(ch.series, fills):
    if col is None:
        ser.format.fill.background()
        ser.format.line.fill.background()
    else:
        ser.format.fill.solid(); ser.format.fill.fore_color.rgb = col
        ser.format.line.fill.background()
# ラベル(手描き: 値を棒の上に)
plot_x0, plot_w = BX + 30, cw - 50
slot = plot_w / len(cats)
for i, v in enumerate(vals):
    lab = f"{v:.1f}" if (i == 0 or i == len(cats) - 1) else f"{v:+.1f}"
    top = max(base[i] + up[i] + down[i], total[i])
    yy = BY + 26 + 262 - 28 - (top - 20) / 14 * (262 - 40)
    text(s, plot_x0 + slot * i, yy - 14, slot, 14, lab, size=9, bold=True, align=PP_ALIGN.CENTER, color=INK, margins=(0, 0, 0, 0))
# 右: 読み取り
rx = BX + cw + 24
rw = BW - cw - 24
band(s, rx, BY, rw, 22, "読み取り", align=PP_ALIGN.LEFT)
text(s, rx, BY + 30, rw, 260, [
    ("**運営要因 4.8pt / 市場要因 1.8pt**", {}),
    ("値引き・物流・SKU は自社の意思決定で戻せる", {"level": 0}),
    ("原材料高は同業共通。価格転嫁は交渉力の回復後", {"level": 0}),
    ("**値引き 2.7pt の内訳**", {"space_before": 6}),
    ("上位 20 顧客への個別値引き 1.9pt", {"level": 0}),
    ("下位顧客への一律値引き 0.8pt", {"level": 0}),
    ("**物流費 1.5pt の内訳**", {"space_before": 6}),
    ("小口配送の増加 0.9pt、拠点固定費 0.6pt", {"level": 0}),
], size=11)
ty = BY + 200
text(s, rx, ty, rw, 16, "施策による回収見込み(FY27)", size=10, bold=True, margins=(0, 0, 0, 0))
line(s, rx, ty + 18, rx + rw, ty + 18, color=DARK, w=0.75)
rec_rows = [("要因", "低下", "回収", "施策"), ("値引き", "2.7", "2.4", "承認制"), ("物流", "1.5", "1.1", "拠点集約"), ("SKU 段取り", "0.6", "0.8", "廃番 + 段取り短縮"), ("原材料", "1.8", "0.9*", "価格転嫁(未検証)")]
cw_ = [70, 40, 40, rw - 150]
yy = ty + 20
for r_i, r in enumerate(rec_rows):
    xx = rx
    for c_i, v in enumerate(r):
        text(s, xx, yy, cw_[c_i], 16, v, size=9.5, bold=(r_i == 0), color=DARK if r_i == 0 else INK, align=PP_ALIGN.RIGHT if c_i in (1, 2) else PP_ALIGN.LEFT, margins=(0, 0, 4, 0) if c_i in (1, 2) else (4 if c_i == 3 else 0, 0, 0, 0), anchor=MSO_ANCHOR.MIDDLE)
        xx += cw_[c_i]
    line(s, rx, yy + 17, rx + rw, yy + 17, color=PALE)
    yy += 18
text(s, rx, yy + 2, rw, 12, "* 転嫁できた場合。9 月の顧客交渉で確認", size=8, color=MUTED, margins=(0, 0, 0, 0))
source(s, "出典: X 社 原価差異分析(FY2023 → FY2025)。要因の帰属は標準原価との差異を、価格差異・数量差異・費目別に配分して算出")

# =====================================================================
# 4. 選択肢 × 評価軸(ハーベイボール + 数値)
# =====================================================================
s = new_slide(
    "物流拠点は 3 拠点集約(案 B)が投資回収と顧客影響の両面で最も優れ、年 5.3 億円を 7 か月で回収する",
    "案 A は効果が 1.2 億円と小さく、案 C は削減 6.8 億円と最大だが上位 20 社しか翌日配送を維持できず、推定 9 億円/年の受注を失う")
hx, hw = BX, 190
cw3 = (BW - hw) / 3
band(s, hx + hw, BY, cw3 - 4, 22, "案 A: 現状維持 + 効率化", fill=LIGHT, color=INK)
band(s, hx + hw + cw3, BY, cw3 - 4, 22, "案 B: 3 拠点に集約(推奨)", fill=ACCENT)
band(s, hx + hw + cw3 * 2, BY, cw3 - 4, 22, "案 C: 2 拠点に集約", fill=LIGHT, color=INK)
crit = [
    ("年間削減額", ["1.2 億円", "**5.3 億円**", "6.8 億円"], None),
    ("初期投資", ["0.4 億円", "3.1 億円", "5.6 億円"], None),
    ("回収期間", ["4 か月", "**7 か月**", "10 か月"], None),
    ("翌日配送の維持", ["全顧客", "上位 50 社 + 150km 圏", "上位 20 社のみ"], [4, 3, 1]),
    ("実施の容易さ", ["3 か月・投資小", "9 か月・並行稼働", "14 か月・大規模移転"], [4, 3, 1]),
    ("在庫・品質リスク", ["変化なし", "移行期に +0.8 億円", "移行期に +1.9 億円"], [4, 3, 2]),
    ("顧客離反リスク", ["なし", "限定的(31 社中 2 社が懸念)", "推定 −9 億円/年"], [4, 3, 1]),
]
y = BY + 26
rh = 34
for lab, cells, balls in crit:
    rowhead(s, hx, y, hw - 4, rh - 4, lab, size=10.5)
    for j, c in enumerate(cells):
        cx0 = hx + hw + cw3 * j
        if j == 1:
            rect(s, cx0, y, cw3 - 4, rh - 4, fill=PALE)
        if balls:
            harvey(s, cx0 + 16, y + (rh - 4) / 2, 14, balls[j])
            text(s, cx0 + 28, y, cw3 - 34, rh - 4, c, size=10.5, anchor=MSO_ANCHOR.MIDDLE)
        else:
            text(s, cx0 + 4, y, cw3 - 12, rh - 4, c, size=11, anchor=MSO_ANCHOR.MIDDLE)
    line(s, hx + hw, y + rh - 2, BX + BW, y + rh - 2, color=LIGHT)
    y += rh
# 総合
rowhead(s, hx, y, hw - 4, 30, "総合評価", size=10.5, fill=DARK)
for j, (c, col) in enumerate([("効果不足", INK), ("推奨: 効果・回収・顧客影響のバランス", WHITE), ("顧客影響が大きい", INK)]):
    cx0 = hx + hw + cw3 * j
    rect(s, cx0, y, cw3 - 4, 30, fill=ACCENT if j == 1 else PALE)
    text(s, cx0 + 4, y, cw3 - 12, 30, c, size=10.5, bold=True, color=col, anchor=MSO_ANCHOR.MIDDLE)
text(s, BX + 600, y + 34, 296, 14, "● 良い  ◕ やや良い  ◑ 中  ◔ やや悪い  ○ 悪い", size=8, color=MUTED, align=PP_ALIGN.RIGHT, margins=(0, 0, 0, 0))
source(s, "出典: 拠点別原価集計(FY2025)、顧客ヒアリング 31 社(2026 年 5〜6 月)、物流会社 3 社見積。顧客離反は上位 20 社以外の翌日配送依存案件を集計")

# =====================================================================
# 5. 2×2 優先順位(散布)
# =====================================================================
s = new_slide(
    "効果が大きく実行しやすい値引き統制と SKU 整理を先行し、拠点集約は投資判断を Q3 に置いて並行で準備する",
    "12 施策を利益効果と実行難易度で整理した。左上 3 施策で効果の 76%(15.6 億円)を占め、6 か月以内に着手できる")
px, py, pw, ph = BX + 40, BY + 10, 560, 300
rect(s, px, py, pw, ph, line=LIGHT)
line(s, px + pw / 2, py, px + pw / 2, py + ph, color=LIGHT, dash=True)
line(s, px, py + ph / 2, px + pw, py + ph / 2, color=LIGHT, dash=True)
# 軸
text(s, px, py + ph + 4, pw, 14, "実行難易度  →  高(投資額・関係部門数・期間で評価)", size=9, color=MUTED, align=PP_ALIGN.CENTER, margins=(0, 0, 0, 0))
ax = text(s, px - 34 - ph / 2 + 8, py + ph / 2 - 8, ph, 16, "利益効果  →  大(FY27 営業利益への寄与、pt)", size=9, color=MUTED, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0))
ax.rotation = 270
for lab, xx, yy in [("先行して実行", px + 6, py + 4), ("投資判断のうえ実行", px + pw / 2 + 6, py + 4), ("手が空けば実行", px + 6, py + ph - 18), ("見送り", px + pw / 2 + 6, py + ph - 18)]:
    text(s, xx, yy, 200, 14, lab, size=9, bold=True, color=MUTED, margins=(0, 0, 0, 0))
pts = [
    ("値引き承認制", 0.12, 0.90, 2.4), ("SKU 廃番", 0.25, 0.72, 0.8), ("小口配送の統合", 0.30, 0.55, 0.4),
    ("3 拠点集約", 0.70, 0.80, 1.1), ("医療機器向け参入", 0.88, 0.85, 3.0),
    ("営業日報の標準化", 0.10, 0.22, 0.1), ("見積テンプレート統一", 0.18, 0.30, 0.15), ("在庫基準の見直し", 0.35, 0.18, 0.2),
    ("基幹システム刷新", 0.82, 0.30, 0.3), ("工場自動化(第 2 期)", 0.75, 0.20, 0.4), ("海外拠点の再編", 0.90, 0.12, 0.2), ("人事制度改定", 0.62, 0.10, 0.1),
]
for lab, fx, fy, imp in pts:
    d = 10 + imp * 9
    cx, cy = px + fx * pw, py + (1 - fy) * ph
    rect(s, cx - d / 2, cy - d / 2, d, d, fill=ACCENT if fy > 0.5 and fx < 0.5 else (DARK if fy > 0.5 else MID), shape=MSO_SHAPE.OVAL)
    lab_s = f"{lab} ({imp:+.1f}pt)" if imp >= 0.3 else lab
    if fx > 0.8:
        text(s, cx - d / 2 - 132, cy - 8, 130, 16, lab_s, size=8.5, margins=(0, 0, 0, 0), anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)
    else:
        text(s, cx + d / 2 + 2, cy - 8, 130, 16, lab_s, size=8.5, margins=(0, 0, 0, 0), anchor=MSO_ANCHOR.MIDDLE)
# 右: 判断
rx = px + pw + 24
rw = BX + BW - rx
band(s, rx, BY, rw, 22, "判断", align=PP_ALIGN.LEFT)
text(s, rx, BY + 30, rw, 300, [
    ("**先行 3 施策(6 か月以内)**", {}),
    ("値引き承認制 +2.4pt", {"level": 0}), ("SKU 廃番 +0.8pt", {"level": 0}), ("小口配送の統合 +0.4pt", {"level": 0}),
    ("**投資判断(Q3)**", {"space_before": 6}),
    ("3 拠点集約 +1.1pt、投資 3.1 億円", {"level": 0}), ("医療機器向け参入 +3pt/3 年、投資 1.2 億円", {"level": 0}),
    ("**見送り・後回し**", {"space_before": 6}),
    ("基幹刷新・自動化は拠点集約の完了後に再評価", {"level": 0}),
    ("円の大きさ = 利益効果(pt)", {"space_before": 8, "size": 8.5, "color": MUTED}),
], size=10.5)
source(s, "出典: 施策別の効果試算(FY2027 営業利益ベース)。実行難易度は投資額・関係部門数・実施期間を 3 段階で評価し合成")

# =====================================================================
# 6. ロードマップ(年度 × 月、担当列、節目、今日線)
# =====================================================================
s = new_slide(
    "値引き統制を 4 月に開始し、拠点集約を翌年 3 月までに完了させる。効果は FY26 下期から段階的に発現する",
    "第 1 波(値引き・SKU・小口統合)は投資 0.6 億円で年 8.2 億円。拠点集約は 6 月の投資判断後、並行稼働 6 か月を経て 3 月に完了する")
lw, ow = 170, 96
gx = BX + lw
gw = BW - lw - ow
months = ["4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", "1月", "2月", "3月", "4月", "5月", "6月"]
pw_ = gw / len(months)
# 年度グループ
band(s, gx, BY, pw_ * 12 - 2, 16, "2026 年度", fill=LIGHT, color=INK, size=9)
band(s, gx + pw_ * 12, BY, pw_ * 3 - 2, 16, "2027 年度", fill=LIGHT, color=INK, size=9)
for i, m in enumerate(months):
    text(s, gx + pw_ * i, BY + 18, pw_, 14, m, size=8.5, bold=True, align=PP_ALIGN.CENTER, margins=(0, 0, 0, 0))
text(s, gx + gw, BY + 18, ow, 14, "担当", size=8.5, bold=True, align=PP_ALIGN.LEFT, margins=(4, 0, 0, 0))
line(s, BX, BY + 34, BX + BW, BY + 34, color=DARK, w=1)
tasks = [
    ("値引き承認制の導入", 0, 2, "営業本部", True),
    ("顧客ランク別の上限設定", 1, 2, "営業本部", False),
    ("SKU 廃番 第 1 弾(900 点)", 1, 4, "商品企画", False),
    ("SKU 廃番 第 2 弾(640 点)", 6, 9, "商品企画", False),
    ("小口配送の統合", 2, 5, "SCM 部", False),
    ("拠点集約: 設計・投資判断", 2, 5, "SCM 部", True),
    ("拠点集約: 契約・移転準備", 5, 6, "SCM 部", False),
    ("拠点集約: 並行稼働・移転", 6, 11, "SCM 部 + 物流会社", True),
    ("共同配送の開始", 9, 11, "SCM 部", False),
    ("医療機器向け: 認証取得", 3, 14, "新規事業室", False),
    ("効果測定・是正(月次)", 6, 14, "PMO", False),
]
rh = 24
y = BY + 38
for lab, a, b, owner, crit_ in tasks:
    text(s, BX, y, lw - 4, rh, lab, size=9.5, anchor=MSO_ANCHOR.MIDDLE, margins=(4, 0, 0, 0))
    bx0 = gx + pw_ * a + 2
    bw0 = pw_ * (b - a + 1) - 4
    rect(s, bx0, y + 5, bw0, rh - 10, fill=ACCENT if crit_ else MID, shape=MSO_SHAPE.PENTAGON)
    text(s, gx + gw, y, ow, rh, owner, size=8.5, color=DARK, anchor=MSO_ANCHOR.MIDDLE, margins=(4, 0, 0, 0))
    line(s, BX, y + rh, BX + BW, y + rh, color=PALE)
    y += rh
# 節目
text(s, BX, y, lw - 4, rh, "節目", size=9.5, bold=True, anchor=MSO_ANCHOR.MIDDLE, margins=(4, 0, 0, 0))
for lab, at in [("承認制 全社適用", 2), ("投資判断", 5), ("集約完了", 11), ("ISO 13485", 14)]:
    cx = gx + pw_ * at + pw_ / 2
    diamond(s, cx, y + rh / 2, 9)
    text(s, cx + 6, y, 80, rh, lab, size=8.5, anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0))
# 縦罫と今日線
for i in range(1, len(months)):
    line(s, gx + pw_ * i, BY + 34, gx + pw_ * i, y + rh, color=PALE)
tx = gx + pw_ * 5 + pw_ * 0.25
line(s, tx, BY + 34, tx, y + rh, color=DARK, w=1, dash=True)
tl = rect(s, tx - 22, BY + 20, 44, 12, fill=DARK)
text(s, tx - 22, BY + 20, 44, 12, "今日 9/7", size=7.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0))
text(s, BX, y + rh + 6, BW, 14, "■ 橙 = クリティカルパス(遅れると効果の発現時期がずれる)  ■ 灰 = その他。並行稼働期(10〜3 月)は在庫を 0.8 億円積み増し、納期遵守率 98% を維持する", size=8.5, color=DARK, margins=(0, 0, 0, 0))
source(s, "出典: 分科会別の実行計画(2026 年 8 月版)。効果の発現時期は各施策の完了月の翌月から計上")

# =====================================================================
# 7. 工程(横の五角形 + 各段の詳細 + 統制ルール)
# =====================================================================
s = new_slide(
    "値引きは申請から回答まで 2 営業日で回し、営業の機動性を落とさずに 5% 超を本部が判断する",
    "現在は担当者の裁量で平均 9.7% の値引き。5% 以下を自動承認、超過を本部判断とし、超過申請を 3 か月で 15% 以下に収束させる")
steps = [
    ("申請", ["営業が案件・値引き率・理由を入力", "案件粗利を自動表示", "所要 10 分"], "入力率 100%"),
    ("自動判定", ["5% 以下は即時承認", "顧客ランク別の上限(8/5/3%)で判定", "超過は本部へ"], "即時承認 85%"),
    ("本部審査", ["粗利・顧客ランク・競合状況で判断", "1 営業日以内", "条件付き承認あり"], "回答 2 営業日"),
    ("回答・実行", ["承認・条件付き・却下を通知", "営業が顧客へ回答", "見積に承認番号を付与"], "承認番号 100%"),
    ("可視化・是正", ["月次で担当者別・顧客別の値引き率", "営業会議で扱う", "四半期で上限を見直し"], "値引き率 5% 以下"),
]
n = len(steps)
gap = 8
sw = (BW - gap * (n - 1)) / n
for i, (head, bullets, kpi) in enumerate(steps):
    x = BX + i * (sw + gap)
    rect(s, x, BY, sw, 34, fill=ACCENT if i == 2 else LIGHT, shape=MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON)
    text(s, x + 10, BY, sw - 20, 34, f"{i + 1:02d}  {head}", size=12, bold=True, color=WHITE if i == 2 else INK, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    text(s, x, BY + 42, sw, 84, [(b, {"level": 0}) for b in bullets], size=10)
    rect(s, x, BY + 130, sw, 26, fill=PALE)
    text(s, x + 4, BY + 130, sw - 8, 26, f"KPI: {kpi}", size=9.5, bold=True, anchor=MSO_ANCHOR.MIDDLE)
# 統制ルール
band(s, BX, BY + 176, BW, 22, "統制ルール(顧客ランク別の値引き上限)", align=PP_ALIGN.LEFT)
cols = [("ランク", 90), ("定義", 250), ("自動承認の上限", 120), ("本部承認の上限", 120), ("対象顧客数", 100), ("FY25 実績の平均値引き", 216)]
x = BX
for lab, w in cols:
    text(s, x, BY + 202, w - 4, 18, lab, size=9.5, bold=True, color=DARK, anchor=MSO_ANCHOR.MIDDLE, margins=(4, 0, 0, 0))
    x += w
line(s, BX, BY + 221, BX + BW, BY + 221, color=DARK, w=0.75)
rows = [("A", "年間売上 5 億円以上、または戦略顧客", "5%", "8%", "20 社", "11.2%(承認制で 8% 以下へ)"),
        ("B", "年間売上 1〜5 億円", "3%", "5%", "84 社", "7.4%"),
        ("C", "年間売上 1 億円未満", "0%", "3%", "228 社", "4.1%")]
y = BY + 224
for r in rows:
    x = BX
    for (lab, w), v in zip(cols, r):
        text(s, x, y, w - 4, 22, v, size=10, anchor=MSO_ANCHOR.MIDDLE, margins=(4, 0, 0, 0), bold=(lab == "ランク"))
        x += w
    line(s, BX, y + 23, BX + BW, y + 23, color=PALE)
    y += 24
text(s, BX, y + 6, BW, 40, [
    ("**適用外**: 特注品(年間 120 件)と新規顧客の初回取引は、案件粗利 20% 以上を条件に営業部長が判断する", {"level": 0}),
    ("**見直し**: 四半期ごとに上限を再設定。承認制の導入後 3 か月で超過申請が 15% を超えていれば上限を 1pt 引き上げる", {"level": 0}),
], size=9.5)
source(s, "出典: 営業日報 12,400 件(FY2025)の値引き実績。即時承認率・回答日数は同業 B 社の導入実績(公開資料)を参考に設定")

# =====================================================================
# 8. 推進体制(体制図 + 役割表)
# =====================================================================
s = new_slide(
    "社長直轄の推進委員会の下に 3 分科会を置き、PMO が横串で効果と課題を月次で判断する",
    "分科会は週次で実行し、意思決定は月次の委員会に集約する。分科会への権限委譲は行わず、投資判断と施策の中止は委員会が決める")
def box(x, y, w, h, title, sub, dark=False):
    rect(s, x, y, w, h, fill=DARK if dark else PALE)
    text(s, x, y + 2, w, h * 0.55, title, size=10.5, bold=True, color=WHITE if dark else INK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM, margins=(2, 0, 2, 0))
    text(s, x, y + h * 0.52, w, h * 0.46, sub, size=8.5, color=WHITE if dark else DARK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP, margins=(2, 0, 2, 0))
ox, ow_ = BX, 560
bw_, bh_ = 170, 44
cx = ox + ow_ / 2
box(cx - bw_ / 2, BY, bw_, bh_, "収益改善推進委員会", "委員長: 社長 / 月次", dark=True)
line(s, cx, BY + bh_, cx, BY + bh_ + 16, color=DARK, w=1)
box(cx - bw_ / 2, BY + bh_ + 16, bw_, bh_, "PMO", "経営企画 3 名(専任) / 週次")
box(cx + bw_ / 2 + 30, BY + bh_ + 16, 150, bh_, "外部支援", "コンサル 4 名、物流設計 2 名")
line(s, cx + bw_ / 2, BY + bh_ + 16 + bh_ / 2, cx + bw_ / 2 + 30, BY + bh_ + 16 + bh_ / 2, color=LIGHT, w=1, dash=True)
y2 = BY + bh_ * 2 + 16 + 20
line(s, cx, BY + bh_ * 2 + 16, cx, y2 - 8, color=DARK, w=1)
ws = [("値引き統制 分科会", "営業本部長 / 8 名"), ("SKU 整理 分科会", "商品企画部長 / 5 名"), ("拠点集約 分科会", "SCM 部長 / 6 名 + 物流 2 社")]
wgap = 12
ww = (ow_ - wgap * 2) / 3
line(s, ox + ww / 2, y2 - 8, ox + ow_ - ww / 2, y2 - 8, color=DARK, w=1)
for i, (t, sub) in enumerate(ws):
    x = ox + i * (ww + wgap)
    line(s, x + ww / 2, y2 - 8, x + ww / 2, y2, color=DARK, w=1)
    box(x, y2, ww, bh_, t, sub)
# 右: 役割表
rx = ox + ow_ + 24
rw = BX + BW - rx
band(s, rx, BY, rw, 22, "会議体と役割", align=PP_ALIGN.LEFT)
rows = [("推進委員会", "月次 90 分", "投資判断、施策の中止・追加、効果の確認"),
        ("PMO", "週次", "進捗・効果・課題の一元管理、委員会資料"),
        ("分科会", "週次", "施策の実行、KPI の計測、課題の一次判断"),
        ("外部支援", "随時", "設計・分析・他社事例の提供")]
y = BY + 28
for a, b, c in rows:
    text(s, rx, y, 70, 36, a, size=9.5, bold=True, anchor=MSO_ANCHOR.TOP, margins=(0, 0, 0, 0))
    text(s, rx + 72, y, 58, 36, b, size=9.5, color=DARK, anchor=MSO_ANCHOR.TOP, margins=(0, 0, 0, 0))
    text(s, rx + 132, y, rw - 132, 36, c, size=9.5, anchor=MSO_ANCHOR.TOP, margins=(0, 0, 0, 0))
    line(s, rx, y + 36, rx + rw, y + 36, color=PALE)
    y += 40
rect(s, rx, y + 6, rw, 60, fill=PALE)
text(s, rx + 6, y + 6, rw - 12, 60, "**意思決定の原則**\n分科会に権限を委譲しない。投資判断・施策の中止・KPI の変更は委員会のみが決め、PMO は判断材料を揃える", size=9.5, anchor=MSO_ANCHOR.MIDDLE)
# 分科会ごとの担務(箱の真下に同じ幅で)
ty = y2 + bh_ + 10
for i, (kpi, tasks_, due) in enumerate([
    ("値引き率 9.7% → 5.0%", ["承認フロー・上限の設計(4 月)", "全社適用と定着(6 月)", "評価指標への反映(FY27)"], "6 月"),
    ("SKU 3,860 → 2,320", ["廃番基準と代替品の設計(5 月)", "第 1 弾 900 点(8 月)", "第 2 弾 640 点(翌 1 月)"], "翌 1 月"),
    ("物流費率 9.4% → 8.3%", ["小口統合(9 月)", "拠点設計・投資判断(6 月)", "並行稼働・移転(翌 3 月)"], "翌 3 月"),
]):
    x = ox + i * (ww + wgap)
    text(s, x, ty, ww, 16, f"KPI: {kpi}", size=9, bold=True, margins=(2, 0, 2, 0))
    line(s, x, ty + 17, x + ww, ty + 17, color=DARK, w=0.75)
    text(s, x, ty + 20, ww, 60, [(t_, {"level": 0}) for t_ in tasks_], size=9, margins=(2, 0, 2, 0))
# 下: 人員サマリ
text(s, ox, ty + 84, ow_, 60, [
    ("**体制の規模**: 専任 3 名 + 兼任 19 名 + 外部 6 名。兼任者は稼働の 20〜30% を充てる", {}),
    ("**KPI の持ち方**: 委員会は営業利益率、分科会は値引き率・SKU 数・物流費率を月次で追う", {"space_before": 4}),
], size=10)
source(s, "出典: 推進体制案(2026 年 8 月、経営会議承認)。人数は各部門の指名者ベース")

# =====================================================================
# 9. As-Is / To-Be(観点別の対比 + 移行の要点)
# =====================================================================
s = new_slide(
    "値引きは「営業担当の裁量」から「本部が粗利で判断する仕組み」へ変え、5 つの観点すべてで基準を明文化する",
    "権限・基準・プロセス・可視化・評価の 5 観点で As-Is と To-Be を対比した。To-Be は同業 B 社の運用を参考に、X 社の顧客構成に合わせて設定")
cols = [("観点", 110), ("As-Is: 営業担当の裁量", 360), ("", 40), ("To-Be: 本部承認制", 386)]
x = BX
for lab, w in cols:
    if lab:
        band(s, x, BY, w - 4, 22, lab, fill=ACCENT if "To-Be" in lab else (LIGHT if "As-Is" in lab else DARK), color=INK if "As-Is" in lab else WHITE, align=PP_ALIGN.LEFT)
    x += w
rows = [("権限", "値引き率に上限なし。担当者が案件ごとに判断", "5% 以下は自動承認、超過は本部が 1 営業日で判断"),
        ("基準", "顧客ランクと値引きが連動せず、下位顧客にも 10% 超が 18%", "顧客ランク A/B/C ごとに上限 8% / 5% / 3%"),
        ("プロセス", "承認は事後報告。月次の営業会議で把握", "申請 → 自動判定 → 本部審査 → 回答を 2 営業日で完了"),
        ("可視化", "粗利は四半期集計。担当者が影響を把握していない", "申請時に案件粗利を自動表示。月次で担当者別・顧客別に共有"),
        ("評価", "売上高のみで評価。値引きは評価に反映されない", "粗利額と値引き率を評価指標に加える(FY27 から)")]
y = BY + 28
rh = 48
for a, b, c in rows:
    rowhead(s, BX, y, 106, rh - 6, a, size=10.5)
    text(s, BX + 110, y, 356, rh - 6, b, size=10.5, anchor=MSO_ANCHOR.MIDDLE, margins=(6, 2, 6, 2))
    ar = rect(s, BX + 476, y + (rh - 6) / 2 - 7, 24, 14, fill=ACCENT, shape=MSO_SHAPE.RIGHT_ARROW)
    rect(s, BX + 510, y, 382, rh - 6, fill=PALE)
    text(s, BX + 510, y, 382, rh - 6, c, size=10.5, anchor=MSO_ANCHOR.MIDDLE, margins=(6, 2, 6, 2))
    line(s, BX + 110, y + rh - 3, BX + BW, y + rh - 3, color=PALE)
    y += rh
band(s, BX, y + 6, BW, 20, "移行の要点", fill=DARK, align=PP_ALIGN.LEFT, size=10)
text(s, BX, y + 30, BW, 60, [
    ("**4 月に承認制を開始し、6 月に全社適用**。初月は超過申請が 34% 出る見込みだが、3 か月で 15% 以下に収束させる(B 社実績)", {"level": 0}),
    ("**評価への反映は FY27 から**。FY26 は粗利の可視化に集中し、担当者が自分の値引きの影響を把握できる状態を先に作る", {"level": 0}),
], size=10)
source(s, "出典: 営業日報 12,400 件(FY2025)、営業部門ヒアリング 24 名、同業 B 社の公開資料")

# =====================================================================
# 10. 実行計画(ネイティブ表: 施策 × 担当 × 期限 × 投資 × 効果 × KPI)
# =====================================================================
s = new_slide(
    "12 施策を 3 波に分け、第 1 波の 5 施策で効果の 6 割を FY26 内に確定させる",
    "第 1 波は投資 0.6 億円で年 8.2 億円。第 2 波は拠点集約を軸に投資 3.2 億円で年 6.7 億円。第 3 波は成長投資で FY27 以降に効果が出る")
hdr = ["#", "施策", "担当", "着手", "完了", "投資\n(億円)", "効果\n(億円/年)", "KPI(目標)"]
widths = [24, 230, 84, 50, 56, 52, 60, 340]
data = [
    ["第 1 波(4〜9 月): 統制と整理"],
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
    ["合計", "", "", "", "", "5.6", "17.1", ""],
]
tx0, ty0 = BX, BY
rows_n, cols_n = len(data) + 1, len(hdr)
gt = s.shapes.add_table(rows_n, cols_n, pt(tx0), pt(ty0), pt(BW), pt(20 * rows_n))
tbl = gt.table
# 既定のスタイル(縞・強調)を外す
tblPr = gt._element.graphic.graphicData.tbl.tblPr
for attr in ("firstRow", "bandRow", "lastRow", "firstCol", "lastCol", "bandCol"):
    tblPr.set(attr, "0")
for j, w in enumerate(widths):
    tbl.columns[j].width = pt(w)
def cell(r, c, s_, size=9.5, bold=False, color=INK, fill=None, align=PP_ALIGN.LEFT):
    ce = tbl.cell(r, c)
    ce.margin_left = ce.margin_right = pt(4)
    ce.margin_top = ce.margin_bottom = pt(2)
    ce.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = ce.text_frame
    tf.text = ""
    p = tf.paragraphs[0]
    p.alignment = align
    _rich(p, s_, size, bold, color)
    if fill is None:
        ce.fill.background()
    else:
        ce.fill.solid(); ce.fill.fore_color.rgb = fill
for j, h in enumerate(hdr):
    cell(0, j, h, size=9, bold=True, color=WHITE, fill=ACCENT, align=PP_ALIGN.CENTER if j in (0, 3, 4, 5, 6) else PP_ALIGN.LEFT)
tbl.rows[0].height = pt(28)
for i, row in enumerate(data, start=1):
    tbl.rows[i].height = pt(22)
    if len(row) == 1:
        cell(i, 0, row[0], size=9.5, bold=True, fill=PALE)
        tbl.cell(i, 0).merge(tbl.cell(i, cols_n - 1))
        continue
    total = row[0] == "合計"
    if total:
        tbl.cell(i, 0).merge(tbl.cell(i, 1))
        row = ["合計", ""] + row[2:]
    for j, v in enumerate(row):
        num = j in (5, 6)
        cell(i, j, v, size=9.5, bold=total, align=PP_ALIGN.RIGHT if num else (PP_ALIGN.CENTER if j in (0, 3, 4) else PP_ALIGN.LEFT), fill=PALE if total else None)
# 罫線: 行の下に薄い線だけ(縦罫なし)
def set_borders(ce, bottom=True):
    tcPr = ce._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
        ln = etree.SubElement(tcPr, qn(tag))
        if tag == "a:lnB" and bottom:
            ln.set("w", "6350")
            sf = etree.SubElement(ln, qn("a:solidFill")); c_ = etree.SubElement(sf, qn("a:srgbClr")); c_.set("val", "D9D9D9")
        else:
            ln.set("w", "0"); etree.SubElement(ln, qn("a:noFill"))
for i in range(rows_n):
    for j in range(cols_n):
        set_borders(tbl.cell(i, j), bottom=i > 0)
source(s, "効果は FY2027 通期の営業利益への寄与(億円)。投資は設備・システム・外部委託の合計。第 3 波の効果は FY28 以降に計上するため「—」")

# =====================================================================
# 11. KPI ダッシュボード(4 タイル + 月次推移 + 論点)
# =====================================================================
s = new_slide(
    "8 月時点で値引き率は 6.8% まで下がり計画を上回るが、SKU 廃番と小口統合は遅れており 9 月に挽回が必要",
    "第 1 波 5 施策のうち 3 つが計画どおり。SKU 廃番は営業からの個別要望で 180 点が保留、小口統合は物流会社との契約が 1 か月遅れている")
tiles = [("平均値引き率", "6.8%", "計画 7.5% / 目標 5.0%", "▲ 計画比 −0.7pt", ACCENT),
         ("SKU 数", "3,140", "計画 2,960 / 目標 2,320", "▼ 計画比 +180", DARK),
         ("小口配送比率", "36%", "計画 31% / 目標 25%", "▼ 計画比 +5pt", DARK),
         ("営業利益率(累計)", "3.4%", "計画 3.2% / FY27 目標 6.5%", "▲ 計画比 +0.2pt", ACCENT)]
tw = (BW - 12 * 3) / 4
for i, (lab, val, plan, delta, col) in enumerate(tiles):
    x = BX + i * (tw + 12)
    rect(s, x, BY, tw, 92, fill=PALE)
    rect(s, x, BY, 4, 92, fill=col)
    text(s, x + 10, BY + 4, tw - 14, 16, lab, size=9.5, color=DARK, margins=(0, 0, 0, 0))
    text(s, x + 10, BY + 20, tw - 14, 36, val, size=26, bold=True, margins=(0, 0, 0, 0))
    text(s, x + 10, BY + 56, tw - 14, 14, plan, size=8.5, color=MUTED, margins=(0, 0, 0, 0))
    text(s, x + 10, BY + 72, tw - 14, 16, delta, size=9.5, bold=True, color=col, margins=(0, 0, 0, 0))
# 月次推移(値引き率)
cw = 520
band(s, BX, BY + 104, cw, 20, "平均値引き率の月次推移(%)と超過申請の件数", align=PP_ALIGN.LEFT, size=10)
cd = CategoryChartData()
cd.categories = ["4月", "5月", "6月", "7月", "8月"]
cd.add_series("実績", (9.4, 8.9, 8.1, 7.3, 6.8))
cd.add_series("計画", (9.2, 8.7, 8.2, 7.8, 7.5))
gf = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, pt(BX), pt(BY + 126), pt(cw), pt(150), cd)
ch = gf.chart
style_chart(ch, size=8.5, legend=XL_LEGEND_POSITION.BOTTOM, gridlines=True)
ch.value_axis.minimum_scale = 5
ch.value_axis.maximum_scale = 10
for i, ser in enumerate(ch.series):
    col = ACCENT if i == 0 else MID
    ser.format.line.color.rgb = col
    ser.format.line.width = Pt(2 if i == 0 else 1.25)
    if i == 1:
        ser.format.line.dash_style = 4  # dash
    ser.marker.style = 8; ser.marker.size = 5
    ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = col
    ser.marker.format.line.color.rgb = col
    ser.smooth = False
    if i == 0:
        dl = ser.data_labels; dl.show_value = True; dl.number_format = '0.0'; dl.number_format_is_linked = False
        dl.font.size = Pt(8.5); dl.font.color.rgb = col; dl.position = XL_LABEL_POSITION.BELOW
ty = BY + 282
text(s, BX, ty, cw, 14, "超過申請(5% 超)の件数と比率", size=9.5, bold=True, margins=(0, 0, 0, 0))
line(s, BX, ty + 16, BX + cw, ty + 16, color=DARK, w=0.75)
ex_rows = [("", "4月", "5月", "6月", "7月", "8月"), ("申請件数", "1,020", "1,080", "1,140", "1,110", "1,090"), ("超過申請", "347", "292", "228", "156", "131"), ("超過比率", "34%", "27%", "20%", "14%", "12%")]
cw_ = [cw - 80 * 5] + [80] * 5
yy = ty + 18
for r_i, r in enumerate(ex_rows):
    xx = BX
    for c_i, v in enumerate(r):
        text(s, xx, yy, cw_[c_i], 15, v, size=9, bold=(r_i == 0 or r_i == 3), color=DARK if r_i == 0 else INK, align=PP_ALIGN.LEFT if c_i == 0 else PP_ALIGN.RIGHT, margins=(0, 0, 0, 0), anchor=MSO_ANCHOR.MIDDLE)
        xx += cw_[c_i]
    line(s, BX, yy + 16, BX + cw, yy + 16, color=PALE)
    yy += 16
# 右: 論点と対応
rx = BX + cw + 24
rw = BX + BW - rx
band(s, rx, BY + 104, rw, 20, "遅れの原因と 9 月の対応", align=PP_ALIGN.LEFT, size=10)
text(s, rx, BY + 130, rw, 170, [
    ("**SKU 廃番: 180 点が保留**", {}),
    ("営業が「特定顧客の専用品」として保留を要望", {"level": 0}),
    ("→ 9 月の委員会で顧客ごとに存廃を決定。代替品の提示を条件に廃番", {"level": 0}),
    ("**小口統合: 契約が 1 か月遅れ**", {"space_before": 6}),
    ("物流会社 2 社との料率交渉が長引いた", {"level": 0}),
    ("→ 9 月 15 日に契約。10 月から統合便を開始し、12 月に計画へ復帰", {"level": 0}),
], size=9.5)
source(s, "出典: PMO 月次報告(2026 年 8 月度)。営業利益率は 4〜8 月累計、計画は 2026 年 4 月の承認計画")

# =====================================================================
# 12. 論点ツリー(仮説と検証状況)
# =====================================================================
s = new_slide(
    "「なぜ利益が減ったか」は 2 つの側面・6 つの論点に分解でき、5 つは検証済み、価格転嫁の可否だけが未検証",
    "売る側(価格・商品)と作る側(生産・物流)に分け、各論点に仮説と検証結果を置いた。検証済みの 5 論点で低下 6.1pt のうち 5.6pt を説明できる")
def tbox(x, y, w, h, t, sub=None, dark=False, fill=None):
    rect(s, x, y, w, h, fill=DARK if dark else (fill or PALE))
    if sub:
        text(s, x, y, w, h * 0.5, t, size=10, bold=True, color=WHITE if dark else INK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM, margins=(3, 0, 3, 0))
        text(s, x, y + h * 0.5, w, h * 0.5, sub, size=8.5, color=WHITE if dark else DARK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP, margins=(3, 0, 3, 0))
    else:
        text(s, x, y, w, h, t, size=10, bold=True, color=WHITE if dark else INK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margins=(3, 0, 3, 0))
x0, w0 = BX, 120
x1, w1 = BX + 150, 130
x2, w2 = BX + 310, 250
x3 = BX + 580
w3 = BX + BW - x3
mid = BY + 168
tbox(x0, mid - 26, w0, 52, "なぜ利益が\n減ったか", "営業利益率 5.9% → 2.1%", dark=True)
line(s, x0 + w0, mid, x1 - 10, mid, color=DARK, w=1)
line(s, x1 - 10, BY + 44 + 46, x1 - 10, BY + 176 + 46, color=DARK, w=1)
leaves = [
    ("売る側", BY + 44, [
        ("価格: 値引きが拡大していないか", "**2.7pt** 平均値引き 4.8% → 9.7%。上位 20 顧客に 73% 集中", "検証済"),
        ("商品: SKU が増えて非効率になっていないか", "**0.6pt** 3,860 SKU、段取り月 1,900 回。下位 40% は売上 3.2%", "検証済"),
        ("価格転嫁: 原材料高を転嫁できていないか", "原材料 +12% に対し価格改定 +3%。転嫁の余地は顧客交渉で未確認", "未検証"),
    ]),
    ("作る側", BY + 176, [
        ("生産: 稼働率・歩留まりが落ちていないか", "**0.3pt** 稼働率 71%(−4pt)。歩留まりは横ばい", "検証済"),
        ("物流: 拠点・配送の固定費が重くないか", "**1.5pt** 稼働率 58%、小口配送 41%", "検証済"),
        ("原材料: 市場要因はどれだけか", "**1.8pt** 特殊鋼 +12%。同業 8 社も同条件", "検証済"),
    ]),
]
for grp, gy, items in leaves:
    tbox(x1, gy + 26, w1, 40, grp)
    line(s, x1 - 10, gy + 46, x1, gy + 46, color=DARK, w=1)
    line(s, x1 + w1, gy + 46, x2 - 10, gy + 46, color=DARK, w=1)
    line(s, x2 - 10, gy + 6 + 17, x2 - 10, gy + 86 + 17, color=DARK, w=1)
    for k, (q, ev, st) in enumerate(items):
        yy = gy + k * 40 + 6
        line(s, x2 - 10, yy + 17, x2, yy + 17, color=DARK, w=1)
        tbox(x2, yy, w2, 34, q, fill=PALE if st == "検証済" else WHITE)
        if st != "検証済":
            rect(s, x2, yy, w2, 34, line=DARK, line_w=1)
        text(s, x3, yy, w3 - 60, 34, ev, size=9.5, anchor=MSO_ANCHOR.MIDDLE, margins=(4, 0, 0, 0))
        rect(s, x3 + w3 - 56, yy + 7, 52, 20, fill=DARK if st == "検証済" else WHITE, line=DARK, line_w=0.75)
        text(s, x3 + w3 - 56, yy + 7, 52, 20, st, size=8.5, bold=True, color=WHITE if st == "検証済" else INK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margins=(0, 0, 0, 0))
band(s, x2, BY + 20, w2, 18, "論点", fill=LIGHT, color=INK, size=9)
band(s, x3, BY + 20, w3, 18, "仮説の検証結果(粗利率への影響)", fill=LIGHT, color=INK, size=9, align=PP_ALIGN.LEFT)
rect(s, BX, BY + 312, BW, 40, fill=PALE)
text(s, BX + 8, BY + 312, BW - 16, 40, "未検証の「価格転嫁」は、9 月の上位 20 顧客との交渉で余地を確認する。転嫁できれば原材料高 1.8pt のうち 0.9pt を回収でき、営業利益率の目標を 6.5% → 7.4% に引き上げられる", size=10, anchor=MSO_ANCHOR.MIDDLE)
source(s, "出典: 原価差異分析(FY2023 → FY2025)、営業日報 12,400 件、生産実績(6 拠点)。影響は粗利率への寄与(pt)")

# ---------------- 保存 ----------------
os.makedirs(os.path.dirname(A.out) or ".", exist_ok=True)
prs.save(A.out)
print(f"saved {A.out} ({page[0]} slides)")
