"""Measure original diagnostic color areas, edge-on phase and perspective."""
import argparse,json
from pathlib import Path
import numpy as np
from PIL import Image
p=argparse.ArgumentParser();p.add_argument('folder');p.add_argument('--pattern',default='f-*.png');p.add_argument('--fps',type=float,default=6);a=p.parse_args();o=Path(a.folder);rows=[]
for i,f in enumerate(sorted(o.glob(a.pattern))):
 im=np.asarray(Image.open(f).convert('RGB'),dtype=float);r,g,b=im[:,:,0],im[:,:,1],im[:,:,2]
 teal=(g>100)&(b>70)&(r<g*0.65);orange=(r>110)&(g>55)&(b<r*0.65)
 mask=teal|orange;ys,xs=np.where(mask)
 if len(xs)<10:continue
 y0,y1=ys.min(),ys.max();widths=[]
 for frac in [0.25,0.75]:
  y=int(y0+(y1-y0)*frac);x=np.where(mask[max(0,y-2):y+3].any(axis=0))[0];widths.append(int(x.max()-x.min()+1) if len(x) else 0)
 heights=[]
 for frac in [0.15,0.85]:
  x=int(xs.min()+(xs.max()-xs.min())*frac);y=np.where(mask[:,max(0,x-1):x+2].any(axis=1))[0];heights.append(int(y.max()-y.min()+1) if len(y) else 0)
 rows.append({'frame':f.name,'time':round(i/a.fps,3),'width':int(xs.max()-xs.min()+1),'height':int(y1-y0+1),'frontPixels':int(teal.sum()),'backPixels':int(orange.sum()),'widthUpperLower':widths,'heightLeftRight':heights,'perspectiveAsymmetry':abs(heights[0]-heights[1])})
assert rows
wide=max(q['width'] for q in rows);edge=[q for q in rows if q['width']<wide*0.15]
front=[q for q in rows if q['frontPixels']>10000 and q['backPixels']<100];back=[q for q in rows if q['backPixels']>10000 and q['frontPixels']<100]
# Y-axis 3D rotation produces unequal left/right edge heights. A 2D x-squash does not.
perspective=[q for q in rows if q['width']>wide*0.15 and q['width']<wide*0.9 and q['perspectiveAsymmetry']>=4 and q['frontPixels']+q['backPixels']>5000]
assert front and back and edge and perspective,'Need both faces, edge-on phase and actual perspective distortion'
report={'passed':True,'maximumWidth':wide,'edgeFrames':edge,'perspectiveFrames':perspective,'frontFrames':len(front),'backFrames':len(back),'frames':rows}
(o/'motion-audit.json').write_text(json.dumps(report,indent=2),encoding='utf-8');print(json.dumps({k:v for k,v in report.items() if k!='frames'}))
