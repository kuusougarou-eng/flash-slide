# -*- coding: utf-8 -*-
"""マニフェスト用アイコン(16/32/64/80px)を生成する。オレンジ角丸 + 白い稜線(ボルト)。"""
import os, sys
from PIL import Image, ImageDraw

sys.stdout.reconfigure(encoding="utf-8")
OUT = os.path.join(os.path.dirname(__file__), "..", "public", "assets")
os.makedirs(OUT, exist_ok=True)
ACCENT = (0xFD, 0x51, 0x08, 255)

def make(size):
    s = size * 8  # supersample
    im = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    r = s * 0.18
    d.rounded_rectangle([0, 0, s - 1, s - 1], radius=r, fill=ACCENT)
    # bolt polygon (normalized coords)
    pts = [(0.58, 0.10), (0.30, 0.55), (0.48, 0.55), (0.40, 0.90), (0.72, 0.42), (0.54, 0.42)]
    d.polygon([(x * s, y * s) for x, y in pts], fill=(255, 255, 255, 255))
    im = im.resize((size, size), Image.LANCZOS)
    im.save(os.path.join(OUT, f"icon-{size}.png"))

for sz in (16, 32, 64, 80, 128):
    make(sz)
print("icons written to", os.path.abspath(OUT))
