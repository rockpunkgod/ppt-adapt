"""Original two-dimensional editable local color fields, OfficeCLI authoring only."""
import argparse,json,subprocess,xml.etree.ElementTree as E,hashlib,random
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);p.add_argument('--seed',type=int,default=18);a=p.parse_args();o=Path(a.out)
if o.exists():raise FileExistsError('Refusing overwrite')
o.mkdir(parents=True);f=o/'soft-color-field.pptx';A='http://schemas.openxmlformats.org/drawingml/2006/main';P='http://schemas.openxmlformats.org/presentationml/2006/main';ops=[];records=[]
def run(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8');assert r.returncode==0,r.stdout+r.stderr;return r.stdout

def add(i,name,x,y,w,h,**props):ops.append(dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name=name,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',**props)))
def text(i,name,t,x,y,w,h,size=14,color='102236'):
 add(i,name,x,y,w,h,text=t,font='Arial',size=str(size),color=color,fill='none',line='none',margin='0cm')
run('create',f);run('set',f,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
for i in [1,2,3]:ops.append(dict(command='add',parent='/',type='slide',props={'layout':'blank','background':'FFFFFF'}))
rng=random.Random(a.seed);jitter=lambda:rng.uniform(-.18,.18)
# COLD + WARM areas, then intermediate local transitions. A fixed seed modifies only placement slightly.
field=[dict(name='cold-cyan',cx=4+jitter(),cy=4+jitter(),w=42,h=33,color='10BFD0',alpha=1),dict(name='warm-gold',cx=31+jitter(),cy=6+jitter(),w=42,h=35,color='FFC823',alpha=1),dict(name='mid-lime',cx=18+jitter(),cy=-1+jitter(),w=31,h=26,color='89DB43',alpha=.94),dict(name='mid-blue',cx=2+jitter(),cy=18+jitter(),w=32,h=27,color='3972E8',alpha=1),dict(name='mid-violet',cx=15+jitter(),cy=23+jitter(),w=31,h=31,color='854CE7',alpha=.98),dict(name='mid-magenta',cx=31+jitter(),cy=23+jitter(),w=34,h=31,color='E34CB5',alpha=.95),dict(name='warm-light',cx=31+jitter(),cy=-1+jitter(),w=23,h=22,color='FFE64D',alpha=.85)]
# Native probe: position0 is center; the ellipse perimeter samples70711 of the bounding-diagonal circle.
stops=[(0,1),(17678,.72),(35355,.35),(53033,.10),(70711,0),(100000,0)]

def stamp(i,c,role):
 c=dict(c);c['h']=c['w'] # Circular brush stamp; calibrated native circle gradient stops.
 add(i,c['name'],c['cx']-c['w']/2,c['cy']-c['h']/2,c['w'],c['h'],geometry='ellipse',gradient=f"radial:{c['color']}-{c['color']}-center",line='none');records.append(dict(slide=i,role=role,**c))
for i in [1,3]:
 for c in field:
  c=dict(c)
  if i==3 and c['name']=='mid-lime':c['color']='C7A2ED'
  stamp(i,c,'art')
 add(i,'type-card',1.6,1.6,12.4,6.4,fill='FFFFFF',opacity='.88',line='none')
 text(i,'eyebrow','NOVA / COLOR SYSTEM',2.25,2.2,11,.7,13)
 text(i,'headline','A spectrum\nof possibilities.',2.2,3.1,11.5,3.3,32)
 text(i,'caption','LOCAL COLOR. EDITABLE LAYERS.' if i==1 else 'ONE LOCAL EDIT / LIME TO LILAC',2.2,6.65,11.3,.6,11)
 text(i,'footer','NOVA 2026 / FORM & FEEL',1.7,17.6,25,.7,12)
text(2,'title','Soft fields obey layer order.',1.7,1.3,30,1.3,28);text(2,'subtitle','Independent single-layer controls / the same two fields in opposite order',1.8,3,30,.9,14)
centers=[5.0,13.0,21.0,29.0];kinds=['A','B','AB','BA']
for cx,kind in zip(centers,kinds):
 text(2,'label-'+kind,{'A':'CORAL ONLY','B':'CYAN ONLY','AB':'CORAL → CYAN','BA':'CYAN → CORAL'}[kind],cx-3.3,5.2,6.6,.9,14)
 for letter in kind:
  c=dict(name=kind+'-'+letter,cx=cx+(-.55 if letter=='A' else .55),cy=10+(.2 if letter=='A' else -.2),w=7.6,h=7.6,color='F06B63' if letter=='A' else '13B4C0',alpha=.70 if letter=='A' else .60);stamp(2,c,'diagnostic')
text(2,'footer','Editable ellipse shapes / radial alpha / normal source-over composition',1.8,17.5,30,.9,13)
fp=o/'foundation.json';fp.write_text(json.dumps(ops),encoding='utf-8');r=run('batch',f,'--input',fp,'--json');assert json.loads(r)['success'];run('close',f)
# L2 lacks per-stop alpha. Preserve native circle path fill and author alpha with L3 only.
raw=[]
for c in records:
 g=E.Element('{'+A+'}gradFill');gs=E.SubElement(g,'{'+A+'}gsLst')
 for pos,t in stops:
  color=E.SubElement(E.SubElement(gs,'{'+A+'}gs',pos=str(pos)),'{'+A+'}srgbClr',val=c['color']);E.SubElement(color,'{'+A+'}alpha',val=str(round(t*c['alpha']*100000)))
 pa=E.SubElement(g,'{'+A+'}path',path='circle');E.SubElement(pa,'{'+A+'}fillToRect',l='50000',t='50000',r='50000',b='50000')
 raw.append(dict(command='raw-set',part=f"/slide[{c['slide']}]",xpath=f"/p:sld/p:cSld/p:spTree/p:sp[p:nvSpPr/p:cNvPr[@name='{c['name']}']]/p:spPr/a:gradFill",action='replace',xml=E.tostring(g,encoding='unicode')))
fp=o/'gradient-alpha.json';fp.write_text(json.dumps(raw),encoding='utf-8');r=run('batch',f,'--input',fp,'--json');assert json.loads(r)['success'];run('close',f)
(o/'model.json').write_text(json.dumps({'seed':a.seed,'stops':stops,'stamps':records,'widthCm':33.87,'heightCm':19.05,'scope':'Original editable local fields; not a claim about Photoshop brush falloff'},indent=2),encoding='utf-8');(o/'validate.json').write_text(run('validate',f,'--json'),encoding='utf-8');print(json.dumps({'deck':str(f),'stamps':len(records)}))
