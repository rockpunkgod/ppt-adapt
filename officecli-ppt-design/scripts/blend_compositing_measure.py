"""Compare 48 interior RGB witnesses to alternative blend semantics, not visual guesses."""
import argparse,json,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
import numpy as np
from PIL import Image
p=argparse.ArgumentParser();p.add_argument('--probe',type=Path,required=True);p.add_argument('--fields',type=Path,required=True);p.add_argument('--evidence',type=Path,required=True);p.add_argument('--report',type=Path,required=True);p.add_argument('--deck-name',default='native-blend-probe.pptx');a=p.parse_args()
if a.report.exists():p.error('Report already exists')
panels=json.loads((a.probe/'panels.json').read_text(encoding='utf-8'));refs={f.stem:np.asarray(Image.open(f).convert('RGB')) for f in a.fields.glob('*.png')};rows=[]
for panel in panels:
 s=panel['slide'];im=np.asarray(Image.open(a.evidence/f'native-{s}.png').convert('RGB'));re=np.asarray(Image.open(a.evidence/f'reopened-{s}.png').convert('RGB'));h,w=im.shape[:2];x,y,rw,rh=panel['rectCm'];actual=[];expected={k:[] for k in refs if k not in ['local-mask','local-gradient-rgba']}
 for row in range(6):
  for col in range(8):
   fx=(col+0.5)/8;fy=(row+0.5)/6;px=round((x+rw*fx)/33.87*w);py=round((y+rh*fy)/19.05*h);actual.append(np.median(im[py-2:py+3,px-2:px+3],axis=(0,1)).astype(int).tolist())
   for k in expected:expected[k].append(refs[k][int(fy*384),int(fx*512)].tolist())
 actual=np.array(actual);scores=[]
 for name,ex in expected.items():
  d=np.abs(actual-np.array(ex));scores.append({'reference':name,'meanAbsRGB':float(d.mean()),'maxAbsRGB':int(d.max())})
 rows.append({'slide':s,'panel':panel['name'],'witnesses':actual.tolist(),'bestCandidates':sorted(scores,key=lambda r:r['meanAbsRGB'])[:6],'roundtripWholePageMax':int(np.abs(im.astype(int)-re.astype(int)).max())})
local=next(p for p in panels if p['name']=='local-mult');im=np.asarray(Image.open(a.evidence/'native-2.png').convert('RGB'));mask=np.asarray(Image.open(a.fields/'local-mask.png'));expected=refs['local-multiply-encoded'];base=refs['backdrop'];localRows=[]
for y0 in range(12,384,12):
 for x0 in range(12,512,12):
  if min(x0%64,64-x0%64,y0%64,64-y0%64)<5:continue
  patch=mask[y0-4:y0+5,x0-4:x0+5]
  if not (np.all(patch==0) or np.all(patch==255)):continue
  x,y,rw,rh=local['rectCm'];px=round((x+rw*x0/512)/33.87*im.shape[1]);py=round((y+rh*y0/384)/19.05*im.shape[0]);actual=im[py,px].astype(int);error=np.abs(actual-expected[y0,x0].astype(int));localRows.append({'xy':[x0,y0],'inside':bool(mask[y0,x0]),'actual':actual.tolist(),'expected':expected[y0,x0].tolist(),'maxError':int(error.max())})
nodes=[]
for f in [a.probe/a.deck_name,a.evidence/'roundtrip.pptx']:
 with zipfile.ZipFile(f) as z:
  for part in ['ppt/slides/slide1.xml','ppt/slides/slide2.xml']:
   root=ET.fromstring(z.read(part));nodes.append({'file':f.name,'part':part,'blendValues':[n.get('blend') for n in root.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}fillOverlay')]})
a.report.write_text(json.dumps({'panels':rows,'localWitnesses':localRows,'localMaxError':max(r['maxError'] for r in localRows),'localInsideCount':sum(r['inside'] for r in localRows),'localOutsideCount':sum(not r['inside'] for r in localRows),'roundtripNodes':nodes,'inputIntegrity':json.loads((a.evidence/'integrity.json').read_text(encoding='utf-8')) if (a.evidence/'integrity.json').exists() else None},indent=2),encoding='utf-8')
for r in rows:print(r['slide'],r['panel'],r['bestCandidates'][:2])
print('Local samples:',len(localRows),'max error:',max(r['maxError'] for r in localRows))
