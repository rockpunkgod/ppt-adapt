"""Read-only editable text + two-background transparency and counter QA."""
import argparse,json,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import numpy as np
from PIL import Image,ImageDraw,ImageFilter
p=argparse.ArgumentParser();p.add_argument('--deck',type=Path,required=True);p.add_argument('--native-dir',type=Path,required=True);p.add_argument('--mask-dir',type=Path,required=True);p.add_argument('--edited-dir',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args()
if a.output.exists():p.error('Fresh report required')
n={'a':'http://schemas.openxmlformats.org/drawingml/2006/main','p':'http://schemas.openxmlformats.org/presentationml/2006/main'};checks={};metrics={}
def need(k,v):checks[k]=bool(v)
for file in [a.deck,a.native_dir/'roundtrip.pptx']:
 with zipfile.ZipFile(file) as z:
  for i in [1,2]:
   r=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'));sp=r.findall('.//p:sp',n);need(file.stem+str(i)+'-one-textbox',len(sp)==1)
   need(file.stem+str(i)+'-editable-text',''.join(x.text or '' for x in r.findall('.//a:t',n))=='OPEN')
   need(file.stem+str(i)+'-no-shape-gradient',not r.findall('.//p:spPr/a:gradFill',n));need(file.stem+str(i)+'-no-pictures',not r.findall('.//p:pic',n))
   stops=r.findall('.//a:rPr/a:gradFill/a:gsLst/a:gs',n);need(file.stem+str(i)+'-five-font-stops',len(stops)==5)
   need(file.stem+str(i)+'-positions',[int(x.get('pos')) for x in stops]==[0,30000,55000,85000,100000])
   need(file.stem+str(i)+'-transparent-last-two',all(x.find('a:srgbClr/a:alpha',n).get('val')=='0' for x in stops[-2:]))
images=[np.asarray(Image.open(a.native_dir/f'native-{i}.png').convert('RGB')).astype(float) for i in [1,2]]
maskim=Image.open(a.mask_dir/'native-1.png').convert('RGB');mask=(np.asarray(maskim).min(axis=2)>245);ys,xs=np.nonzero(mask);x0,x1=xs.min(),xs.max();span=x1-x0
bg=[x[0,0] for x in images];delta=bg[1]-bg[0];alpha=1-np.sum((images[1]-images[0])*delta,axis=2)/float(delta@delta)
for name,lo,hi in [('opaque',0.03,0.2),('fade',0.62,0.75),('transparent',0.91,0.99)]:
 region=mask.copy();region[:,:int(x0+lo*span)]=False;region[:,int(x0+hi*span):]=False;values=alpha[region];metrics[name]={'count':len(values),'medianAlpha':float(np.median(values))};need(name+'-has-glyph-samples',len(values)>100)
need('opaque-glyph-independent-of-background',metrics['opaque']['medianAlpha']>0.98)
need('fade-is-partly-transparent',0.05<metrics['fade']['medianAlpha']<0.9)
need('transparent-end-reveals-both-backgrounds',abs(metrics['transparent']['medianAlpha'])<0.02)
# A solid-white reference gives real enclosed counters without guessing letter coordinates.
bw=Image.fromarray(np.where(mask,255,0).astype('uint8')).copy();ImageDraw.floodfill(bw,(0,0),128);holes=np.asarray(bw)==0
holes=Image.fromarray((holes*255).astype('uint8')).filter(ImageFilter.MinFilter(7));holes=np.asarray(holes)>0;metrics['counterPixels']=int(holes.sum());need('real-enclosed-letter-counters',holes.sum()>500)
for i,image in enumerate(images):
 error=np.max(abs(image[holes]-bg[i]));metrics['counterBackgroundError'+str(i+1)]=float(error);need('counters-reveal-background-'+str(i+1),error<1)
 again=np.asarray(Image.open(a.native_dir/f'reopened-{i+1}.png').convert('RGB')).astype(float);need('save-reopen-identical-pixels-'+str(i+1),np.array_equal(image,again))
rows=json.loads((a.native_dir/'native-readback.json').read_text(encoding='utf-8-sig'))
need('native-font-fill-gradient',all(x['fontFillType']==3 for x in rows));need('native-reopened-text-editable',all(x['hasTextFrame']==-1 and x['text']=='OPEN' for x in rows))
edited=json.loads((a.edited_dir/'native-readback.json').read_text(encoding='utf-8-sig'));need('actual-edit-open-to-opal',all(x['text']=='OPAL' and x['hasTextFrame']==-1 and x['fontFillType']==3 for x in edited))
with zipfile.ZipFile(a.edited_dir/'roundtrip.pptx') as z:
 for i in [1,2]:
  r=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'));need('edit-still-has-five-font-stops-'+str(i),len(r.findall('.//a:rPr/a:gradFill/a:gsLst/a:gs',n))==5)
  im=np.asarray(Image.open(a.edited_dir/f'native-{i}.png').convert('RGB'));again=np.asarray(Image.open(a.edited_dir/f'reopened-{i}.png').convert('RGB'));need('edit-reopen-identical-pixels-'+str(i),np.array_equal(im,again));need('edit-changed-visible-glyphs-'+str(i),np.count_nonzero(np.any(im!=images[i-1],axis=2))>1000)
result={'passed':all(checks.values()),'checks':checks,'metrics':metrics,'method':'alpha=1-dot(lightPixel-darkPixel, lightBG-darkBG)/squaredNorm(BGdelta), sampled only inside the independent solid-white glyph mask; no screenshot alpha channel assumed'}
a.output.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps({'passed':result['passed'],'checks':len(checks),'failed':[k for k,v in checks.items() if not v],'metrics':metrics},indent=2));raise SystemExit(0 if result['passed'] else 1)
