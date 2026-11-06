"""OfficeCLI-native matched-contour interpolation; exact sweep mesh for COM union."""
import argparse,json,math,copy,zipfile,hashlib,subprocess,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--geometry',required=True);p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);a=p.parse_args();o=Path(a.out)
if o.exists():raise FileExistsError('Refusing overwrite')
o.mkdir(parents=True);src=Path(a.geometry).read_bytes();g=json.loads(src);(o/'input-geometry.json').write_bytes(src);bounds=g['bounds'];unit=g['unitsToCm'];P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';N={'p':P,'a':A};E.register_namespace('p',P);E.register_namespace('a',A)
original=[[[float(x-bounds['x'])*unit,float(y-bounds['y'])*unit] for x,y in ring] for ring in g['original']];stroke=[[[float(x-bounds['x'])*unit,float(y-bounds['y'])*unit] for x,y in ring] for ring in g['strokes'][0]['contours']]
assert len(original)==2 and len(stroke)==4 and g['text']=='P' and g['font']=='Arial Black' and abs(bounds['height']*unit-5)<1e-6
assert g.get('strokeMethod','').startswith('Shapely2.1.2') and g['strokes'][0]['offsetMm']==.3,'Validated buffer band required; do not use GDI Widen negative fixture'
file=o/'blend-base.pptx'
def run(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8');assert r.returncode==0,r.stdout+r.stderr;return r.stdout
def batch(ops,name):
 f=o/(name+'.json');f.write_text(json.dumps(ops),encoding='utf-8');r=run('batch',file,'--input',f,'--json');assert json.loads(r)['success'];(o/(name+'-log.json')).write_text(r,encoding='utf-8')
run('create',file);run('set',file,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm');ops=[]
def label(i,name,text,x,y,w,h,size=16,color='17324E'):
 ops.append(dict(command='add',parent=f'/slide[{i}]',type='shape',props={'name':name,'text':text,'x':f'{x}cm','y':f'{y}cm','width':f'{w}cm','height':f'{h}cm','font':'Arial','size':str(size),'color':color,'fill':'none','line':'none','margin':'0cm'}))
for i,(title,bg,sub) in enumerate([('One contour. Real depth.','F6F8FC','Two matching endpoints / 200 intermediate steps'),('A shadow follows its contour.','F6F8FC','45-degree sweep / native vector union and crop'),('One plane. Two densities.','102236','200 and 30 intermediate steps / editable isometric planes')],1):
 ops.append(dict(command='add',parent='/',type='slide',props={'layout':'blank','background':bg}));fg='FFFFFF' if i==3 else '17324E';label(i,'title',title,1.7,1.3,30,1.6,28,fg);label(i,'subtitle',sub,1.8,3.15,30,0.8,15,fg);label(i,'footer','NOVA / CONTOUR STUDY',1.8,17.7,29,.7,11,fg)
label(1,'endpoint-a-label','NEAR',3.2,10.3,4,1,14);label(1,'endpoint-b-label','FAR / 94% SCALE',8.9,10.3,6,1,14);label(1,'hero-label','INTERPOLATED BODY + SEPARATE CAP',18.2,15.2,13,1,13)
label(2,'clip-label','Clipped to this panel',18.2,5.3,12,1,20);label(2,'method','The hole contributes to the sweep.\nIt is never replaced with a rectangle.\n\nVector geometry remains editable.',18.2,7.1,12,4.5,16)
label(3,'dense-label','200 STEPS',3.2,15.4,11,1,15,'FFFFFF');label(3,'sparse-label','30 STEPS',19,15.4,11,1,15,'FFFFFF')
batch(ops,'foundation');run('close',file)
with zipfile.ZipFile(file) as z:trees=[E.fromstring(z.read(f'ppt/slides/slide{i}.xml')).find('p:cSld/p:spTree',N) for i in [1,2,3]]
ident=200000;records=[]
def rgba(c1,c2,t):return ''.join(f'{round(int(c1[k:k+2],16)*(1-t)+int(c2[k:k+2],16)*t):02X}' for k in [0,2,4])
def mapped(contours,fn):return [[[float(v) for v in fn(x,y)] for x,y in ring] for ring in contours]
def addshape(slide,name,contours,color,line=None):
 global ident
 pts=[p for ring in contours for p in ring];xs=[p[0] for p in pts];ys=[p[1] for p in pts];x0,y0,x1,y1=min(xs),min(ys),max(xs),max(ys);w,h=x1-x0,y1-y0
 if w<1e-7 or h<1e-7:return
 ident+=1;s=E.Element('{'+P+'}sp');nv=E.SubElement(s,'{'+P+'}nvSpPr');E.SubElement(nv,'{'+P+'}cNvPr',id=str(ident),name=name);E.SubElement(nv,'{'+P+'}cNvSpPr');E.SubElement(nv,'{'+P+'}nvPr');sp=E.SubElement(s,'{'+P+'}spPr');xf=E.SubElement(sp,'{'+A+'}xfrm');E.SubElement(xf,'{'+A+'}off',x=str(round(x0*360000)),y=str(round(y0*360000)));E.SubElement(xf,'{'+A+'}ext',cx=str(round(w*360000)),cy=str(round(h*360000)))
 geom=E.SubElement(sp,'{'+A+'}custGeom')
 for k in ['avLst','gdLst','ahLst','cxnLst']:E.SubElement(geom,'{'+A+'}'+k)
 E.SubElement(geom,'{'+A+'}rect',l='0',t='0',r='r',b='b');pl=E.SubElement(geom,'{'+A+'}pathLst');path=E.SubElement(pl,'{'+A+'}path',w='100000',h='100000')
 for ring in contours:
  for j,(x,y) in enumerate(ring):q=E.SubElement(path,'{'+A+'}'+('moveTo' if j==0 else 'lnTo'));E.SubElement(q,'{'+A+'}pt',x=str(round((x-x0)/w*100000)),y=str(round((y-y0)/h*100000)))
  E.SubElement(path,'{'+A+'}close')
 if color=='none':E.SubElement(sp,'{'+A+'}noFill')
 else:f=E.SubElement(sp,'{'+A+'}solidFill');E.SubElement(f,'{'+A+'}srgbClr',val=color)
 ln=E.SubElement(sp,'{'+A+'}ln')
 if line:f=E.SubElement(ln,'{'+A+'}solidFill');E.SubElement(f,'{'+A+'}srgbClr',val=line)
 else:E.SubElement(ln,'{'+A+'}noFill')
 trees[slide-1].append(s);records.append({'slide':slide,'name':name,'id':ident,'contourCount':len(contours),'pointCount':len(pts),'boundsCm':[x0,y0,w,h]})
# Interpolation uses corresponding nodes of the same original; no arbitrary-glyph morph claim.
def blended(t):return mapped(original,lambda x,y:((1-.06*t)*x+.9*t,(1-.06*t)*y+.7*t))
addshape(1,'endpoint-near',mapped(original,lambda x,y:(3.2+.8*x,5.2+.8*y)),'26BFAF');addshape(1,'endpoint-far',mapped(blended(1),lambda x,y:(8.9+.8*x,5.2+.8*y)),'17324E')
for j in range(201,-1,-1):
 t=j/201;addshape(1,f'solid-{j:03d}',mapped(blended(t),lambda x,y:(19+1.5*x,5.2+1.5*y)),rgba('26BFAF','17324E',t))
addshape(1,'solid-front-cap',mapped(original,lambda x,y:(19+1.5*x,5.2+1.5*y)),'F2CA70')
# Exact swept-region construction: endpoints plus each boundary segment swept over d.
clip=[4.2,4.8,10.3,9.7];addshape(2,'panel-background',[[[4.2,4.8],[14.5,4.8],[14.5,14.5],[4.2,14.5]]],'E7EFF7')
front=mapped(original,lambda x,y:(5+x,5+y));delta=(6.,6.);addshape(2,'shadow-start',front,'17324E');addshape(2,'shadow-end',mapped(front,lambda x,y:(x+6,y+6)),'17324E');k=0
for ring in front:
 for j,p0 in enumerate(ring):
  p1=ring[(j+1)%len(ring)];cross=(p1[0]-p0[0])*delta[1]-(p1[1]-p0[1])*delta[0]
  if abs(cross)<1e-8:continue
  quad=[p0,p1,[p1[0]+6,p1[1]+6],[p0[0]+6,p0[1]+6]];addshape(2,f'shadow-edge-{k:03d}',[quad],'17324E');k+=1
addshape(2,'shadow-clip',[[[4.2,4.8],[14.5,4.8],[14.5,14.5],[4.2,14.5]]],'FFFFFF');addshape(2,'shadow-front-cap',front,'F2CA70')
# Widened filled contour bands are projected as geometry, including all four rings.
iso=mapped(stroke,lambda x,y:(math.sqrt(3)/2*(x-y),.5*(x+y)));ix=min(x for r in iso for x,y in r);iy=min(y for r in iso for x,y in r);iso=mapped(iso,lambda x,y:(x-ix,y-iy))
for steps,xbase,prefix in [(200,3.2,'iso-dense'),(30,19.,'iso-sparse')]:
 for j in range(steps+1,-1,-1):
  t=j/(steps+1);addshape(3,f'{prefix}-{j:03d}',mapped(iso,lambda x,y:(xbase+x,5.1+y+4.2*t)),rgba('39D7CA','102236',t))
ops=[{'command':'raw-set','part':f'/slide[{i}]','xpath':'/p:sld/p:cSld/p:spTree','action':'replace','xml':E.tostring(tree,encoding='unicode')} for i,tree in enumerate(trees,1)];batch(ops,'geometry');run('close',file)
model={'glyph':g['text'],'font':g['font'],'sourceSha256':hashlib.sha256(src).hexdigest(),'strokeHalfWidthMm':g['strokes'][0]['offsetMm'],'originalCm':original,'strokeCm':stroke,'solid':{'steps':200,'objects':202,'backScale':.94,'backTranslation':[.9,.7],'origin':[19,5.2],'scale':1.5},'shadow':{'origin':[5,5],'delta':[6,6],'clipCm':clip,'meshShapes':k+2,'finalFill':'17324E','finalAlpha':.18,'construction':'Exact boundary sweep union plus start/end; compare independent202-position oracle'},'isometric':{'matrix':[[math.sqrt(3)/2,-math.sqrt(3)/2],[.5,.5]],'depth':4.2,'steps':[200,30],'objectCounts':[202,32]},'shapes':records}
(o/'model.json').write_text(json.dumps(model,indent=2),encoding='utf-8');(o/'validation.json').write_text(run('validate',file,'--json'),encoding='utf-8');print(json.dumps({'deck':str(file),'customShapeCount':len(records),'shadowMesh':k+2}))
