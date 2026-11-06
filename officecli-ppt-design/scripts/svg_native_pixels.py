"""Fixture-specific native silhouette/color/hole transparency audit; Pillow + NumPy."""
import argparse,json
from pathlib import Path
import numpy as np
from PIL import Image,ImageChops
p=argparse.ArgumentParser();p.add_argument('--evidence',required=True);p.add_argument('--structure',required=True);p.add_argument('--report',required=True);a=p.parse_args();out=Path(a.report)
if out.exists():raise FileExistsError('Refusing overwrite')
d=Path(a.evidence);m=json.loads(Path(a.structure).read_text(encoding='utf-8'));assert m['viewBox']==[0,0,600,360]
images={k:Image.open(d/(k+'.png')).convert('RGB') for k in ('before','converted','edited','reopen')};assert all(im.size==(1600,900) for im in images.values())
box=m['imageBoxEmu'];cw,ch=m['canvasEmu']
def location(x,y):return round((box[0]+x/600*box[2])/cw*1600),round((box[1]+y/360*box[3])/ch*900)
def sample(im,x,y,color):
 px,py=location(x,y);want=tuple(bytes.fromhex(color));pixels=list(im.crop((px-2,py-2,px+3,py+3)).getdata());assert all(q==want for q in pixels),(x,y,color,pixels)
samples=[]
for phase,im in images.items():
 for x,y,color,meaning in [(265,150,'D9E4E6','upper hole reveals native background'),(265,210,'EFDCC7','lower hole reveals different background'),(20,40,'D9E4E6','outside contour is transparent'),(565,180,m['sourcePalette'][1],'second path preserves its source color'),(265,60,m['editedPalette'][0] if phase in ('edited','reopen') else m['sourcePalette'][0],'first real path color')]:
  sample(im,x,y,color);samples.append({'phase':phase,'svgPoint':[x,y],'color':color,'meaning':meaning})
x0,y0=location(0,0);x1,y1=location(600,360)
def mask(im,palette):
 ar=np.asarray(im.crop((x0,y0,x1,y1))).astype(int);return np.logical_or.reduce([np.max(abs(ar-np.array(tuple(bytes.fromhex(q)))),axis=2)<6 for q in palette])
def iou(a,b):return float((a&b).sum()/(a|b).sum())
before=mask(images['before'],m['sourcePalette']);converted=mask(images['converted'],m['sourcePalette']);edited=mask(images['edited'],m['editedPalette'])
initial=iou(before,converted);edit=iou(converted,edited);assert initial>0.98 and edit>0.999,(initial,edit)
assert ImageChops.difference(images['edited'],images['reopen']).getbbox() is None
for phase in ('converted','edited','reopen'):
 for region in [(0,0,1600,y0-3),(0,y1+3,1600,900)]:assert ImageChops.difference(images['before'].crop(region),images[phase].crop(region)).getbbox() is None
out.write_text(json.dumps({'passed':True,'beforeAfterSilhouetteIoU':initial,'editPreservesSilhouetteIoU':edit,'reopenPixelIdentical':True,'samples':samples,'scope':'Semantic transparency/color samples and silhouette check for the original 600x360 two-path ring/arrow fixture.'},indent=2),encoding='utf-8');print(out)
