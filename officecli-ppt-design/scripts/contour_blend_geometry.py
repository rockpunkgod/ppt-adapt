"""Independent coordinate correspondence check of saved custGeom planes."""
import json,zipfile,xml.etree.ElementTree as E,math,argparse
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--deck',required=True);p.add_argument('--model',required=True);p.add_argument('--out',required=True);a=p.parse_args();out=Path(a.out)
if out.exists():raise FileExistsError('Refusing overwrite')
m=json.loads(Path(a.model).read_text());ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'};maxerr=0;checked=0

def rings(s):
 xf=s.find('p:spPr/a:xfrm',ns);off=xf.find('a:off',ns);ext=xf.find('a:ext',ns);x,y=int(off.get('x'))/360000,int(off.get('y'))/360000;w,h=int(ext.get('cx'))/360000,int(ext.get('cy'))/360000;allrings=[]
 for path in s.findall('p:spPr/a:custGeom/a:pathLst/a:path',ns):
  pw,ph=float(path.get('w')),float(path.get('h'));r=[]
  for cmd in path:
   tag=cmd.tag.split('}')[-1]
   if tag=='moveTo':
    if r:allrings.append(r)
    r=[]
   if tag in ('moveTo','lnTo'):
    q=cmd.find('a:pt',ns);r.append([x+float(q.get('x'))/pw*w,y+float(q.get('y'))/ph*h])
   elif tag=='close':allrings.append(r);r=[]
   else:raise AssertionError('Unexpected non-linear path')
 return allrings
with zipfile.ZipFile(a.deck) as z:
 for slide in (1,3):
  r=E.fromstring(z.read('ppt/slides/slide%d.xml'%slide))
  for s in r.findall('p:cSld/p:spTree/p:sp',ns):
   name=s.find('p:nvSpPr/p:cNvPr',ns).get('name')
   if name.startswith('solid-') and name[6:].isdigit():
    t=int(name[6:])/201;expect=[[[19+1.5*((1-.06*t)*x+.9*t),5.2+1.5*((1-.06*t)*y+.7*t)] for x,y in ring] for ring in m['originalCm']]
   elif name.startswith(('iso-dense-','iso-sparse-')):
    dense=name.startswith('iso-dense-');t=int(name.split('-')[-1])/(201 if dense else 31);xb=3.2 if dense else 19
    shape=[[[math.sqrt(3)/2*(x-y),.5*(x+y)] for x,y in ring] for ring in m['strokeCm']];ix=min(x for ring in shape for x,y in ring);iy=min(y for ring in shape for x,y in ring);expect=[[[xb+x-ix,5.1+y-iy+4.2*t] for x,y in ring] for ring in shape]
   else:continue
   c1,c2=('26BFAF','17324E') if name.startswith('solid-') else ('39D7CA','102236')
   color=''.join(f'{round(int(c1[k:k+2],16)*(1-t)+int(c2[k:k+2],16)*t):02X}' for k in [0,2,4])
   assert s.find('p:spPr/a:solidFill/a:srgbClr',ns).get('val')==color,(name,'color')
   actual=rings(s);assert len(actual)==len(expect)
   for ar,er in zip(actual,expect):
    assert len(ar)==len(er)
    maxerr=max(maxerr,max(math.dist(x,y) for x,y in zip(ar,er)))
   checked+=1
assert checked==436 and maxerr<.0001,(checked,maxerr)
result={'passed':True,'checkedPlanes':checked,'checkedLinearColors':checked,'maximumCoordinateErrorCm':maxerr,'scope':'All202 solid and202+32 iso saved path points independently reconstructed; this alone does not validate band distance/native appearance'}
out.write_text(json.dumps(result,indent=2));print(json.dumps(result))
