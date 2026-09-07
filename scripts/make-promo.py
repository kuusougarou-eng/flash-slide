# -*- coding: utf-8 -*-
"""
マーケ用 30 秒動画の組み立て(録画フレーム → 合成フレーム → ffmpeg)。
  python scripts/make-promo.py <framesDir> <out.mp4> --slide x0,y0,x1,y1 --pane x0,y0,x1,y1 [--fps 6]
- 各フレーム(PowerPoint ウィンドウの PrintWindow キャプチャ)から「スライド領域」と「タスクペイン」を切り出し、
  1920×1080 に左右で合成(左: スライド、右: ペインを拡大)。
- 先頭にタイトルカード、末尾にエンドカードを付ける(訴求: 速度と品質)。
"""
import sys, os, glob, argparse, subprocess
from PIL import Image, ImageDraw, ImageFont

W, H = 1920, 1080
PANE_TOP = 0.62  # ペイン上部(入力・結果・カード)の割合
PANE_BOT = 0.12  # ペイン下部(実行ボタン/タイマー/出力先)の割合
FONT_B = "C:/Windows/Fonts/meiryob.ttc"
FONT_R = "C:/Windows/Fonts/meiryo.ttc"

def font(path, size):
    return ImageFont.truetype(path, size)

def card(lines, out, sub=None, bg="#141414", fg="#ffffff", accent="#fd5108"):
    im = Image.new("RGB", (W, H), bg)
    d = ImageDraw.Draw(im)
    y = H // 2 - 40 * len(lines)
    for i, (txt, size) in enumerate(lines):
        f = font(FONT_B, size)
        tw = d.textlength(txt, font=f)
        d.text(((W - tw) / 2, y), txt, font=f, fill=fg)
        y += size + 24
    if sub:
        f = font(FONT_R, 34)
        tw = d.textlength(sub, font=f)
        d.text(((W - tw) / 2, y + 20), sub, font=f, fill="#bdbdbd")
    d.rectangle([W // 2 - 60, H - 140, W // 2 + 60, H - 134], fill=accent)
    im.save(out)

def compose(frame, slide_box, pane_box, caption=None):
    src = Image.open(frame).convert("RGB")
    canvas = Image.new("RGB", (W, H), "#1f1f1f")
    d = ImageDraw.Draw(canvas)
    # 左: スライド(幅 1240 に収める)
    s = src.crop(slide_box)
    k = min(1240 / s.width, (H - 120) / s.height)
    s = s.resize((int(s.width * k), int(s.height * k)), Image.LANCZOS)
    sx, sy = 40, (H - s.height) // 2
    canvas.paste(s, (sx, sy))
    # 右: ペイン。中央の空白を抜いて「上部(入力・結果・カード)」+「下部(ボタン/タイマー)」を積み、読める大きさに拡大
    full = src.crop(pane_box)
    top_h = int(full.height * PANE_TOP)
    bot_h = int(full.height * PANE_BOT)
    p = Image.new("RGB", (full.width, top_h + bot_h + 6), "#e6e6e6")
    p.paste(full.crop((0, 0, full.width, top_h)), (0, 0))
    p.paste(full.crop((0, full.height - bot_h, full.width, full.height)), (0, top_h + 6))
    k2 = min(620 / p.width, (H - 80) / p.height)
    p = p.resize((int(p.width * k2), int(p.height * k2)), Image.LANCZOS)
    px, py = W - p.width - 40, (H - p.height) // 2
    canvas.paste(p, (px, py))
    d.rectangle([px - 1, py - 1, px + p.width, py + p.height], outline="#3a3a3a")
    if caption:
        f = font(FONT_B, 40)
        tw = d.textlength(caption, font=f)
        d.rectangle([0, H - 96, W, H], fill=(0, 0, 0))
        d.text(((W - tw) / 2, H - 78), caption, font=f, fill="#ffffff")
    return canvas

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("frames"); ap.add_argument("out")
    ap.add_argument("--slide", required=True); ap.add_argument("--pane", required=True)
    ap.add_argument("--fps", type=float, default=6)
    ap.add_argument("--captions", default="", help="'start:end:text;...' フレーム番号範囲ごとの字幕")
    a = ap.parse_args()
    slide = tuple(int(v) for v in a.slide.split(","))
    pane = tuple(int(v) for v in a.pane.split(","))
    caps = []
    for seg in filter(None, a.captions.split(";")):
        s, e, t = seg.split(":", 2); caps.append((int(s), int(e), t))
    frames = sorted(glob.glob(os.path.join(a.frames, "f*.png")))
    work = os.path.join(a.frames, "_out"); os.makedirs(work, exist_ok=True)
    n = 0
    def emit(im):
        nonlocal n; n += 1; im.save(os.path.join(work, "c%06d.png" % n))
    # タイトルカード 3 秒
    card([("Flash Slide", 120), ("考えを、そのまま1〜2枚に。", 64)], os.path.join(work, "_title.png"), sub="PowerPoint アドイン")
    t = Image.open(os.path.join(work, "_title.png"))
    for _ in range(int(a.fps * 3)): emit(t)
    for i, fr in enumerate(frames, 1):
        cap = next((c[2] for c in caps if c[0] <= i <= c[1]), None)
        emit(compose(fr, slide, pane, cap))
    # エンドカード 3 秒
    card([("貼る → 数秒 → 編集できるスライド", 64), ("4 案から選ぶだけ。文章で指示しない。", 48)], os.path.join(work, "_end.png"), sub="Flash Slide")
    e = Image.open(os.path.join(work, "_end.png"))
    for _ in range(int(a.fps * 3)): emit(e)
    subprocess.check_call(["ffmpeg", "-y", "-framerate", ("%.3f" % a.fps), "-i", os.path.join(work, "c%06d.png"), "-c:v", "libx264", "-pix_fmt", "yuv420p", "-crf", "20", "-movflags", "+faststart", a.out])
    print("wrote", a.out, "frames", n, "sec", n / a.fps)

if __name__ == "__main__":
    main()
