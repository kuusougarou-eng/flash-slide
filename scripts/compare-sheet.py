# before/after のスライド PNG を左右に並べた比較シートを作る
#   python scripts/compare-sheet.py [--before debug/before] [--after debug] [--out debug/compare] [--per 4]
# before/after の両方にある t-<name>-1.png を名前順に並べ、--per 枚ずつ 1 枚のシートにする
import argparse, os, glob, sys
from PIL import Image, ImageDraw, ImageFont

ap = argparse.ArgumentParser()
ap.add_argument("--before", default="debug/before")
ap.add_argument("--after", default="debug")
ap.add_argument("--out", default="debug/compare")
ap.add_argument("--per", type=int, default=4)
ap.add_argument("--width", type=int, default=760)
a = ap.parse_args()

names = sorted(os.path.basename(p)[2:-6] for p in glob.glob(os.path.join(a.before, "t-*-1.png")))
names = [n for n in names if os.path.exists(os.path.join(a.after, f"t-{n}-1.png"))]
if not names:
    print("no pairs"); sys.exit(1)
W = a.width
H = int(W * 9 / 16)
label_h = 26
gap = 12
try:
    font = ImageFont.truetype("C:/Windows/Fonts/meiryo.ttc", 15)
except Exception:
    font = ImageFont.load_default()
sheets = []
for i in range(0, len(names), a.per):
    chunk = names[i : i + a.per]
    sheet = Image.new("RGB", (W * 2 + gap * 3, (H + label_h + gap) * len(chunk) + gap + label_h), "white")
    d = ImageDraw.Draw(sheet)
    d.text((gap, gap), "BEFORE", fill="#666", font=font)
    d.text((W + gap * 2, gap), "AFTER", fill="#666", font=font)
    y = gap + label_h
    for n in chunk:
        for k, src in enumerate([a.before, a.after]):
            im = Image.open(os.path.join(src, f"t-{n}-1.png")).convert("RGB").resize((W, H), Image.LANCZOS)
            x = gap + k * (W + gap)
            sheet.paste(im, (x, y + label_h))
            d.rectangle([x, y + label_h, x + W - 1, y + label_h + H - 1], outline="#BBB")
        d.text((gap, y + 4), n, fill="#222", font=font)
        y += H + label_h + gap
    out = f"{a.out}-{len(sheets) + 1}.png"
    sheet.save(out)
    sheets.append(out)
print("\n".join(sheets))
