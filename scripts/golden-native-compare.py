"""Compare actual PowerPoint renders of the reference and compiled OPC packages."""
import json,sys
from pathlib import Path
rendered=sys.argv[1] if len(sys.argv)>1 else 'rendered'
report_name=sys.argv[2] if len(sys.argv)>2 else 'visual-report.json'
from PIL import Image, ImageChops, ImageStat, ImageDraw
root=Path('debug/golden-native')
results=[]
for p in sorted((root/'reference').glob('*.png')):
    q=root/rendered/p.name
    if not q.exists():
        results.append({'id':p.stem,'missing':True})
        continue
    a=Image.open(p).convert('RGB'); b=Image.open(q).convert('RGB')
    if a.size!=b.size:
        results.append({'id':p.stem,'sizeMismatch':True})
        continue
    d=ImageChops.difference(a,b)
    results.append({'id':p.stem,'pixelIdentical':d.getbbox() is None,'meanAbsoluteError':sum(ImageStat.Stat(d).mean)/3})
report={'total':len(results),'identical':sum(r.get('pixelIdentical',False) for r in results),'results':results}
(root/report_name).write_text(json.dumps(report,indent=2),encoding='utf-8')
print(json.dumps(report,indent=2))
for part in range(2):
    canvas=Image.new('RGB',(1600,1000),'#dddddd')
    for k in range(16):
        i=part*16+k+1; p=root/'reference'/f'{i:02}.png'
        if not p.exists():continue
        im=Image.open(p);im.thumbnail((390,220)); x=k%4*400;y=k//4*250
        canvas.paste(im,(x,y+22));ImageDraw.Draw(canvas).text((x+6,y+4),str(i),fill='black')
    canvas.save(root/f'reference-contact-{part+1}.png')
