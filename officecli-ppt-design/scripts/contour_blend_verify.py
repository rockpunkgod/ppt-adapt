"""Read final native vector count/holes and independent native shadow pixels."""
import argparse,json,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import numpy as np
from PIL import Image,ImageDraw
p=argparse.ArgumentParser();p.add_argument('--deck',required=True);p.add_argument('--model',required=True);p.add_argument('--native',required=True);p.add_argument('--out',required=True);a=p.parse_args();out=Path(a.out)
if out.exists():raise FileExistsError('Refusing overwrite')
m=json.loads(Path(a.model).read_text());ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
with zipfile.ZipFile(a.deck) as z:
 slides=[E.fromstring(z.read('ppt/slides/slide%d.xml'%i)) for i in (1,2,3)]
 assert not any('/media/' in n for n in z.namelist()),'No bitmap/vector picture substitutes'
 maps=[{s.find('p:nvSpPr/p:cNvPr',ns).get('name'):s for s in r.findall('p:cSld/p:spTree/p:sp',ns)} for r in slides]
 for name in ('endpoint-near','solid-100','solid-front-cap'):
  s=maps[0][name];assert s.find('p:spPr/a:custGeom',ns) is not None;assert len(s.findall('.//a:close',ns))==2
 for prefix,n in [('solid-',202),('iso-dense-',202),('iso-sparse-',32)]:
  d=maps[0] if prefix=='solid-' else maps[2];arr=[v for k,v in d.items() if k.startswith(prefix) and k[len(prefix):].isdigit()];assert len(arr)==n,(prefix,len(arr));assert all(v.find('p:spPr/a:custGeom',ns) is not None for v in arr)
 shadow=maps[1]['long-shadow'];assert shadow.find('p:spPr/a:custGeom',ns) is not None;assert not any(k.startswith('shadow-edge-') for k in maps[1]);assert 'shadow-clip' not in maps[1]
im=np.asarray(Image.open(Path(a.native)/'slide-2.png').convert('RGB'));H,W=im.shape[:2];sx=W/33.87;sy=H/19.05

def raster(rings):
 acc=np.zeros((H,W),np.int8)
 for r in rings:
  signed=sum(r[i][0]*r[(i+1)%len(r)][1]-r[(i+1)%len(r)][0]*r[i][1] for i in range(len(r)))
  bit=Image.new('1',(W,H));ImageDraw.Draw(bit).polygon([(round(x*sx),round(y*sy)) for x,y in r],fill=1);acc+=np.asarray(bit,np.int8)*(1 if signed>0 else -1)
 return acc!=0
P=m['originalCm'];ox,oy=m['shadow']['origin'];dx,dy=m['shadow']['delta'];expected=np.zeros((H,W),bool)
for j in range(202):expected|=raster([[[x+ox+dx*j/201,y+oy+dy*j/201] for x,y in r] for r in P])
x,y,w,h=m['shadow']['clipCm'];clip=np.zeros((H,W),bool);clip[round(y*sy):round((y+h)*sy)+1,round(x*sx):round((x+w)*sx)+1]=True
cap=raster([[[x+ox,y+oy] for x,y in r] for r in P]);expected&=clip&~cap
# Actual translucent navy on the panel. Far-right labels are outside this ROI.
color=np.array([int('17324E'[i:i+2],16) for i in (0,2,4)])*.18+np.array([231,239,247])*.82
actual=np.max(abs(im.astype(float)-color),axis=2)<3
roi=np.zeros((H,W),bool);roi[round(4.7*sy):round(14.6*sy),round(4.1*sx):round(14.6*sx)]=True;actual&=roi
intersection=np.sum(actual&expected);union=np.sum(actual|expected);iou=float(intersection/union);assert iou>.97,iou
im1=np.asarray(Image.open(Path(a.native)/'slide-1.png').convert('RGB'));bg=np.array([246,248,252]);holeChecks=[]
for name,(cx,cy) in [('near',(3.2+.8*2.1,5.2+.8*1.5)),('far',(8.9+.8*(.94*2.1+.9),5.2+.8*(.94*1.5+.7)))]:
 px,py=round(cx*sx),round(cy*sy);patch=im1[py-1:py+2,px-1:px+2].astype(float);err=float(np.max(abs(patch-bg)));assert err<2,(name,'hole',err);holeChecks.append({'endpoint':name,'maximumBackgroundError':err})
result={'passed':True,'nativeEndpointHoleChecks':holeChecks,'nativeShadowPixelIoU':iou,'expectedCompositeRGB':color.tolist(),'objects':{'solid':202,'dense':202,'sparse':32,'shadow':1},'editableCustGeom':True,'pictures':0,'scope':'Native merged/clipped silhouette + true compound path structure; no Multiply or arbitrary-glyph correspondence claim'}
out.write_text(json.dumps(result,indent=2));print(json.dumps(result))
