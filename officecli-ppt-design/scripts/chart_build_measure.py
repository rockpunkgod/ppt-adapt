"""Fixture-specific independent native pixel visibility audit; Pillow + NumPy."""
import argparse,json
from pathlib import Path
import numpy as np
from PIL import Image
p=argparse.ArgumentParser();p.add_argument('--native',required=True);p.add_argument('--desktop',required=True);p.add_argument('--report',required=True);p.add_argument('--elements',action='store_true');a=p.parse_args();out=Path(a.report)
if out.exists():raise FileExistsError('Refusing overwrite')
native=Path(a.native);desktop=Path(a.desktop)
colors=[np.array([0,127,134]),np.array([232,135,53])]
ref=np.asarray(Image.open(native/'after-1.png').convert('RGB')).astype(int)
bars=[]
for series,col in enumerate(colors):
 mask=(np.max(np.abs(ref-col),axis=2)<8)
 present=np.flatnonzero(mask[250:688].sum(axis=0)>35)
 runs=np.split(present,np.where(np.diff(present)>1)[0]+1)
 for run in runs:
  if len(run)>40:bars.append({'series':series,'x':int(round(float(run.mean()))),'left':int(run[0]),'right':int(run[-1])})
bars.sort(key=lambda q:q['x']);assert len(bars)==8
for k,b in enumerate(bars):assert b['series']==k%2
def normalize(path):
 im=Image.open(path).convert('RGB');w,h=im.size
 if w/h>16/9:
  nw=round(h*16/9);x=(w-nw)//2;im=im.crop((x,0,x+nw,h))
 else:
  nh=round(w*9/16);y=(h-nh)//2;im=im.crop((0,y,w,y+nh))
 return np.asarray(im.resize((1600,900),Image.Resampling.LANCZOS)).astype(int)
checks=[]
for slide,n in ([(1,8),(2,8)] if a.elements else [(1,2),(2,4)]):
 for click in range(n+1):
  im=normalize(desktop/f's{slide}-c{click}.png');ratios=[]
  for bar in bars:
   region=im[650:675,bar['x']-12:bar['x']+12]
   ratios.append(float((np.max(np.abs(region-colors[bar['series']]),axis=2)<16).mean()))
  actual=[i for i,q in enumerate(ratios) if q>0.9]
  assert all(q<0.05 or q>0.9 for q in ratios),('Ambiguous/partial state',slide,click,ratios)
  if click==0:expected=[]
  elif a.elements:expected=sorted(([0,2,4,6,1,3,5,7] if slide==1 else list(range(8)))[:click])
  elif slide==1:expected=[0,2,4,6] if click==1 else list(range(8))
  else:expected=list(range(2*click))
  checks.append({'slide':slide,'click':click,'expectedBarIndices':expected,'actualBarIndices':actual,'ratios':ratios,'passed':actual==expected})
assert all(q['passed'] for q in checks),checks
out.write_text(json.dumps({'passed':True,'checks':checks,'barCalibration':bars,'scope':'Fixed fixture/color/axis geometry; exact physical desktop snapshots after settled clicks. Movie must also be reviewed for intermediate motion.'},indent=2),encoding='utf-8');print(out)
