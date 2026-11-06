import argparse,json,xml.etree.ElementTree as E,zipfile
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--deck',type=Path,required=True);p.add_argument('--model',type=Path,required=True);p.add_argument('--report',type=Path,required=True);p.add_argument('--ellipse-cap',action='store_true');a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
model=json.loads(a.model.read_text(encoding='utf-8'));ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'};checks={};rows=[]
def shape(root,name):return next(q for q in root.findall('.//p:sp',ns) if q.find('p:nvSpPr/p:cNvPr',ns).get('name')==name)
def outline(q):return q.find('p:spPr/a:custGeom/a:pathLst/a:path',ns)
def area(ps):return sum(x*y2-x2*y for (x,y),(x2,y2) in zip(ps,ps[1:]+ps[:1]))/2
def contours(path):
 out=[];c=[]
 for cmd in path:
  k=cmd.tag.rsplit('}',1)[-1]
  if k=='moveTo':c=[]
  if k in ['moveTo','lnTo']:q=cmd[0];c.append((int(q.get('x'))/1e6,int(q.get('y'))/1e6))
  if k=='close':out.append(c)
 return out
def bounds(samples,u):
 for l,r in zip(samples,samples[1:]):
  if l[0]-1e-6<=u<=r[0]+1e-6:
   f=(u-l[0])/(r[0]-l[0]);return l[1]+(r[1]-l[1])*f,l[2]+(r[2]-l[2])*f
 raise ValueError('Output x outside target')
with zipfile.ZipFile(a.deck) as z:
 roots={i:E.fromstring(z.read(f'ppt/slides/slide{i}.xml')) for i in [1,2,3,4]};checks['editableOriginalTextBackup']=shape(roots[1],'editable-backup').find('.//a:t',ns).text=='BOLD';checks['nativeSourceGlyph']=outline(shape(roots[2],'source-glyph')) is not None
 checks['sourceEightClosedContours']=len(outline(shape(roots[2],'source-glyph')).findall('a:close',ns))==8
 for target in model['targets']:
  s=target['slide'];q=shape(roots[s],'warped-glyph');path=outline(q);cs=contours(path);pts=[p for c in cs for p in c];samples=target['targetSamples']
  checks[f's{s}eightNativeClosedContours']=len(cs)==8;checks[f's{s}noRasterPictures']=not roots[s].findall('.//p:pic',ns);checks[f's{s}noPresetTextWarp']=not roots[s].findall('.//a:prstTxWarp',ns)
  checks[f's{s}positiveBandHeight']=all(l[0]<r[0] for l,r in zip(samples,samples[1:])) and min(b-t for _,t,b in samples)>0
  if a.ellipse_cap and s==3:
   del checks[f's{s}positiveBandHeight']
   checks['specificEllipseCap']=target['name']=='ellipse-cap' and target.get('collapsedBoundaryTips') is True
   checks['exactCollapsedBoundaryTips']=samples[0][0]==0 and samples[-1][0]==1 and all(t==b for _,t,b in [samples[0],samples[-1]])
   checks['positiveInteriorHeight']=all(t<b for _,t,b in samples[1:-1]) and all(l[0]<r[0] for l,r in zip(samples,samples[1:]))
   left,right=target['intersections'];scale=target['verticalNormalization']
   checks['upperBoundaryIsEllipse']=0<left<.5<right<1 and all(abs((2*(left+(right-left)*u)-1)**2+(2*t*scale-1)**2-1)<1e-10 for u,t,b in samples)
   checks['lowerBoundaryIsBowedKnife']=all(abs(b*scale-(.14+.035*(left+(right-left)*u)+.08*(left+(right-left)*u)*(1-left-(right-left)*u)))<1e-10 for u,t,b in samples)
   checks['actualGlyphTipsCollapse']=all(any(abs(u-edge)<1e-6 for u,v in pts) and all(abs(v-samples[index][1])<2e-6 for u,v in pts if abs(u-edge)<1e-6) for edge,index in [(0,0),(1,-1)])
  inside=True
  for u,v in pts:
   t,b=bounds(samples,u)
   if not (t-0.000002<=v<=b+0.000002):inside=False
  checks[f's{s}verticesWithinCustomBoundary']=inside
  checks[f's{s}matchingModelVertexCount']=len(pts)==target['vertices']
  border=outline(shape(roots[s],'target-boundary'));checks[f's{s}actualCustomTarget']=border is not None and len(border.findall('a:close',ns))==1
  areas=[area(c) for c in cs];rows.append({'slide':s,'contours':len(cs),'vertices':len(pts),'signedAreas':areas,'minimumBandHeight':min(b-t for _,t,b in samples)})
result={'passed':all(checks.values()),'checks':checks,'targets':rows,'note':'Checks actual native paths and bounded monotone model; no claim of Adobe Envelope algorithm equivalence or arbitrary target topology. Native rendered hole/letter checks remain separate.'}
a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2))
