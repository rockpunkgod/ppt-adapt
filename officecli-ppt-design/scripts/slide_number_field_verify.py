import argparse,json
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as E
import numpy as np
from PIL import Image
p=argparse.ArgumentParser();p.add_argument('--evidence',type=Path,required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
checks=[]
with ZipFile(a.evidence/'field-final.pptx') as z:
 for n in ['ppt/slideMasters/slideMaster1.xml','ppt/slideLayouts/slideLayout1.xml']:
  r=E.fromstring(z.read(n));fields=r.findall('.//a:fld[@type="slidenum"]',ns)
  checks.append({'case':n,'passed':len(fields)==1,'nativeFields':len(fields)})
 for n in [f'ppt/slides/slide{i}.xml' for i in range(1,5)]:
  r=E.fromstring(z.read(n));fields=r.findall('.//a:fld[@type="slidenum"]',ns)
  checks.append({'case':n,'passed':len(fields)==1,'nativeFields':len(fields)})
  shapes=r.findall('.//p:sp',ns)
  controls=[s for s in shapes if s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='independent-number']
  expected=n.rsplit('slide',1)[1].split('.')[0]
  checks.append({'case':n+' independent text','passed':len(controls)==1 and ''.join(controls[0].itertext())==expected})
templates={}
for i in range(1,5):
 im=np.asarray(Image.open(a.evidence/f'duplicated-{i}.png').convert('RGB'))
 templates[i]=np.min(im[500:600,1420:1560],axis=2)<150
for f in sorted(a.evidence.glob('*.png')):
 img=np.asarray(Image.open(f).convert('RGB')).astype(int)
 top=img[500:600,1420:1560];bottom=img[700:800,1420:1560]
 error=int(np.max(np.abs(top-bottom)));ink=int(np.count_nonzero(np.min(bottom,axis=2)<150))
 mask=np.min(bottom,axis=2)<150
 scores={i:float(np.count_nonzero(mask&t)/np.count_nonzero(mask|t)) for i,t in templates.items()}
 expected=int(f.stem.rsplit('-',1)[1]);score=scores[expected];wrong=max(v for k,v in scores.items() if k!=expected)
 checks.append({'case':f.name,'passed':score>=.95 and score-wrong>.1 and ink>100,'maxRgbEdgeDifference':error,'inkPixels':ink,'expectedDigit':expected,'templateIoU':scores})
states=json.loads((a.evidence/'states.json').read_text(encoding='utf-8-sig'))
first=[r['slideId'] for r in states if r['phase']=='initial'];dup=[r['slideId'] for r in states if r['phase']=='duplicated'];order=[r['slideId'] for r in states if r['phase']=='reordered']
checks.append({'case':'duplicate and reorder identity','passed':len(set(dup))==4 and first==[dup[0],dup[1],dup[3]] and order==[dup[3],dup[0],dup[1],dup[2]]})
for i in range(1,5):
 same=Image.open(a.evidence/f'reordered-{i}.png').tobytes()==Image.open(a.evidence/f'reopened-{i}.png').tobytes()
 checks.append({'case':f'reopen page {i}','passed':same})
report={'passed':all(c['passed'] for c in checks),'checks':checks,'scope':'Native slide-number fields; duplicate/reorder; master style via layout; SaveAs/reopen; no plain-text field substitution.'}
a.report.write_text(json.dumps(report,indent=2),encoding='utf-8');print(json.dumps(report));assert report['passed']
