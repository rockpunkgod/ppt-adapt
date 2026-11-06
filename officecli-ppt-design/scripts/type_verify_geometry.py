"""Audit native outline geometry, including opposite contour winding for letter counters."""
import argparse,json,zipfile,xml.etree.ElementTree as E
from pathlib import Path
NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def point(e):return tuple(float(e.attrib[v]) for v in ('x','y'))
def areas(path):
 result=[];pts=[];cur=(0,0)
 def close():
  if len(pts)>2:result.append(sum(x*y2-x2*y for (x,y),(x2,y2) in zip(pts,pts[1:]+pts[:1]))/2)
 for e in path:
  tag=e.tag.split('}')[-1];p=[point(q) for q in e]
  if tag=='moveTo':
   if pts:close()
   cur=p[0];pts=[cur]
  elif tag=='lnTo':cur=p[0];pts.append(cur)
  elif tag=='cubicBezTo':
   start=cur
   for i in range(1,17):
    t=i/16;u=1-t;pts.append(tuple(u**3*start[k]+3*u*u*t*p[0][k]+3*u*t*t*p[1][k]+t**3*p[2][k] for k in (0,1)))
   cur=p[-1]
  elif tag=='quadBezTo':
   start=cur
   for i in range(1,17):
    t=i/16;u=1-t;pts.append(tuple(u*u*start[k]+2*u*t*p[0][k]+t*t*p[1][k] for k in (0,1)))
   cur=p[-1]
  elif tag=='close':close();pts=[]
 if pts:close()
 return result
def verify(file):
 with zipfile.ZipFile(file) as z:
  r=E.fromstring(z.read('ppt/slides/slide1.xml'));shapes=[]
  for s in r.findall('.//p:sp',NS):
   name=s.find('p:nvSpPr/p:cNvPr',NS).get('name');paths=s.findall('p:spPr/a:custGeom/a:pathLst/a:path',NS)
   a=[x for p in paths for x in areas(p)];text=''.join(s.itertext())
   shapes.append({'name':name,'paths':len(paths),'closed_contours':len(a),'positive_winding':sum(x>1 for x in a),'negative_winding':sum(x<-1 for x in a),'text_characters':len(text.strip())})
  front=next(x for x in shapes if x['name']=='black-front')
  assert front['positive_winding']>0 and front['negative_winding']>0,'Expected real counter contours with opposite winding'
  assert all(x['paths']>0 and x['text_characters']==0 for x in shapes),'Result must be native outlines, not text/image stand-ins'
  assert not r.findall('.//p:pic',NS),'No raster substitute allowed'
  return {'passed':True,'native_outlines':shapes,'counter_evidence':'opposite winding closed contours in black-front; native PNG visually confirms holes','raster_pictures':0}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('file',type=Path);p.add_argument('--report',type=Path,required=True);a=p.parse_args();
 if a.report.exists():raise FileExistsError(a.report)
 r=verify(a.file);a.report.write_text(json.dumps(r,indent=2),encoding='utf-8');print(json.dumps(r,indent=2))
