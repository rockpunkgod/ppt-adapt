import argparse,json,math,shutil,subprocess,zipfile,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--input',type=Path,required=True);p.add_argument('--officecli',required=True);p.add_argument('--outdir',type=Path,required=True);p.add_argument('--targets',type=Path);p.add_argument('--ellipse-cap',action='store_true',help='Specific ellipse upper cap cut by a bowed knife line; only its two boundary tips may collapse');a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
if a.ellipse_cap and a.targets:raise ValueError('Ellipse cap cannot be combined with arbitrary targets')
a.outdir.mkdir(parents=True);ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def polyline_contours(path):
 w=float(path.get('w'));h=float(path.get('h'));contours=[];current=[];position=None
 def point(q):return (float(q.get('x'))/w,float(q.get('y'))/h)
 for cmd in path:
  kind=cmd.tag.rsplit('}',1)[-1];pts=[point(q) for q in cmd]
  if kind=='moveTo':
   if current:raise ValueError('Unclosed native contour')
   current=[pts[0]];position=pts[0]
  elif kind in ['lnTo','quadBezTo','cubicBezTo']:
   if position is None:raise ValueError('Path command before Move')
   controls=[position]+pts;length=sum(math.dist(x,y) for x,y in zip(controls,controls[1:]));n=max(1,math.ceil(length/0.0025))
   for k in range(1,n+1):
    t=k/n;work=controls[:]
    while len(work)>1:work=[((1-t)*x[0]+t*y[0],(1-t)*x[1]+t*y[1]) for x,y in zip(work,work[1:])]
    current.append(work[0])
   position=pts[-1]
  elif kind=='close':
   contours.append(current);current=[];position=None
  else:raise ValueError('Unsupported source path command: '+kind)
 if current:raise ValueError('Open native contour')
 return contours
with zipfile.ZipFile(a.input) as z:
 root=E.fromstring(z.read('ppt/slides/slide2.xml'));source=next(q for q in root.findall('.//p:sp',ns) if q.find('p:nvSpPr/p:cNvPr',ns).get('name')=='source-glyph');geometry=source.find('p:spPr/a:custGeom',ns)
 if geometry is None:raise ValueError('Expected real native glyph outline')
 contours=[c for path in geometry.findall('a:pathLst/a:path',ns) for c in polyline_contours(path)]
 if not all(-0.000001<=u<=1.000001 and -0.000001<=v<=1.000001 for c in contours for u,v in c):raise ValueError('Source outline outside normalized native bounds')
def bez(v,t):return (1-t)**3*v[0]+3*(1-t)**2*t*v[1]+3*(1-t)*t*t*v[2]+t**3*v[3]
if a.targets:targets=json.loads(a.targets.read_text(encoding='utf-8'))
else:
 targets=[{'name':'asymmetric-ribbon','samples':[[i/128,bez([0.22,0.00,0.42,0.06],i/128),bez([0.86,0.68,1.05,0.88],i/128)] for i in range(129)]},{'name':'wavy-band','samples':[[i/128,0.16+0.12*math.sin(math.pi*i/128)+0.06*i/128,0.86-0.10*math.sin(2*math.pi*i/128)+0.04*i/128] for i in range(129)]}]
if a.ellipse_cap:
 # The course cuts an ellipse with a shallow bowed, slanted knife line.
 # Solve both intersections instead of replacing the tips by epsilon height.
 def upper(x):return .5-.5*math.sqrt(max(0,1-(2*x-1)**2))
 def knife(x):return .14+.035*x+.02*4*x*(1-x)
 def intersection(left,right):
  sign=upper(left)-knife(left)
  for _ in range(80):
   mid=(left+right)/2
   if (upper(mid)-knife(mid))*sign>0:left=mid
   else:right=mid
  return (left+right)/2
 l=intersection(0,.5);r=intersection(.5,1)
 raw=[(i/256,upper(l+(r-l)*i/256),knife(l+(r-l)*i/256)) for i in range(257)]
 scale=max(b for _,_,b in raw);samples=[[u,t/scale,b/scale] for u,t,b in raw]
 for index in [0,-1]:samples[index][1]=samples[index][2]
 targets[0]={'name':'ellipse-cap','samples':samples,'collapsedBoundaryTips':True,'ellipse':{'cx':.5,'cy':.5,'rx':.5,'ry':.5},'knife':'0.14 + 0.035*x + 0.08*x*(1-x)','intersections':[l,r],'verticalNormalization':scale}
if len(targets)!=2:raise ValueError('Fixture expects two target profiles')
def validate(samples,collapsed=False):
 if samples[0][0]!=0 or samples[-1][0]!=1:raise ValueError('Target must span x0..1')
 if collapsed:
  if not (samples[0][1]==samples[0][2] and samples[-1][1]==samples[-1][2]):raise ValueError('Expected exact zero-height boundary tips')
  if not all(0<=t<=b<=1 for _,t,b in samples) or not all(t<b for _,t,b in samples[1:-1]):raise ValueError('Only the two boundary tips may collapse')
 elif not all(0<=u<=1 and 0<=t<b<=1 for u,t,b in samples):raise ValueError('Target is not a bounded positive-height band')
 if not all(x[0]<y[0] for x,y in zip(samples,samples[1:])):raise ValueError('Target x must be strictly monotone')
for target in targets:
 if target.get('collapsedBoundaryTips') and not a.ellipse_cap:raise ValueError('Custom targets retain the original strictly positive-height contract')
 validate(target['samples'],target.get('collapsedBoundaryTips',False))
def bounds(samples,u):
 u=max(0,min(1,u));lo=0;hi=len(samples)-1
 while hi-lo>1:
  mid=(lo+hi)//2
  if samples[mid][0]<=u:lo=mid
  else:hi=mid
 x,t,b=samples[lo];xx,tt,bb=samples[hi];f=(u-x)/(xx-x);return t+(tt-t)*f,b+(bb-b)*f
def geom(cs):
 out=['<a:custGeom xmlns:a="'+ns['a']+'"><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="w" b="h"/><a:pathLst><a:path w="1000000" h="1000000">']
 for c in cs:
  for j,(u,v) in enumerate(c):out.append(f'<a:{"moveTo" if j==0 else "lnTo"}><a:pt x="{round(u*1000000)}" y="{round(v*1000000)}"/></a:{"moveTo" if j==0 else "lnTo"}>')
  out.append('<a:close/>')
 out.append('</a:path></a:pathLst></a:custGeom>');return ''.join(out)
deck=a.outdir/'envelope.pptx';shutil.copy2(a.input,deck)
def cli(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8')
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
cli('open',deck);reports=[];operations=[]
for s,target in enumerate(targets,3):
 samples=target['samples'];validate(samples,target.get('collapsedBoundaryTips',False));warped=[]
 for c in contours:
  out=[]
  for u,v in c:
   top,bottom=bounds(samples,u);out.append((u,top+(bottom-top)*v))
  warped.append(out)
 border=[(x,t) for x,t,b in samples]+[(x,b) for x,t,b in reversed(samples)]
 for name,cs in [('target-boundary',[border]),('warped-glyph',warped)]:operations.append({'command':'raw-set','part':f'/slide[{s}]','xpath':f'//p:sp[p:nvSpPr/p:cNvPr[@name="{name}"]]/p:spPr/a:prstGeom','action':'replace','xml':geom(cs)})
 reports.append({'slide':s,'name':target['name'],'contours':len(warped),'vertices':sum(map(len,warped)),'minimumJacobianNormalized':min(b-t for _,t,b in samples),'mapping':'x=u; y=top(u)+(bottom(u)-top(u))*v','targetSamples':samples,**({'collapsedBoundaryTips':True,'ellipse':target['ellipse'],'knife':target['knife'],'intersections':target['intersections'],'verticalNormalization':target['verticalNormalization']} if target.get('collapsedBoundaryTips') else {})})
batch=a.outdir/'native-geometry-batch.json';batch.write_text(json.dumps(operations),encoding='utf-8');cli('batch',deck,'--input',batch,'--json');cli('close',deck);print(cli('validate',deck,'--json'));(a.outdir/'warp-model.json').write_text(json.dumps({'sourceContours':len(contours),'sourceVertices':sum(map(len,contours)),'targets':reports,'limits':('Specific ellipse cap allows only its two boundary tips to collapse; interior height stays positive. ' if a.ellipse_cap else '')+'Only x-monotone non-crossing top/bottom bands. Positive interior Jacobian prevents interior folding before polyline discretization; collapsed boundary tips require separate native topology and readability checks. Not an Adobe Envelope solver or arbitrary topology.'},indent=2),encoding='utf-8')
