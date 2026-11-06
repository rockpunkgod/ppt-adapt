"""True master/layout timing, authored with OfficeCLI; no copied slide shapes."""
import argparse,json,subprocess,zipfile,copy,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--out',required=True);p.add_argument('--officecli',required=True);a=p.parse_args();o=Path(a.out)
if o.exists():raise FileExistsError('Refusing overwrite')
o.mkdir(parents=True);deck=o/'master-loop.pptx';carrier=o/'timing-carrier.pptx';P='http://schemas.openxmlformats.org/presentationml/2006/main';ns={'p':P};E.register_namespace('p',P)
def run(file,*args):
 r=subprocess.run([a.officecli,args[0],str(file),*map(str,args[1:])],capture_output=True,text=True,encoding='utf-8');assert r.returncode==0,r.stdout+r.stderr;return r.stdout
def shape(parent,name,x,y,w,h,color,text='',size='18',geometry='rect'):
 return dict(command='add',parent=parent,type='shape',props=dict(name=name,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',fill=color,line='none',text=text,size=size,font='Arial',color='FFFFFF',geometry=geometry,margin='0cm'))
def batch(file,ops,name):
 f=o/(name+'.json');f.write_text(json.dumps(ops),encoding='utf-8');r=run(file,'batch','--input',f,'--json');assert json.loads(r)['success'];(o/(name+'-log.json')).write_text(r,encoding='utf-8')
for f in [deck,carrier]:run(f,'create');run(f,'set','/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
run(deck,'set','/slidelayout[1]','--prop','name=Loop A');run(deck,'set','/slidelayout[2]','--prop','name=Loop B')
ops=[]
# Master marker persists on A/A/B; distinct layout markers exist only on their layout.
ops += [shape('/slidemaster[1]','master-track',2,5.4,8,0.05,'526576'),shape('/slidemaster[1]','master-orange',2,4.8,1.2,1.2,'E88735',geometry='ellipse')]
ops += [shape('/slidelayout[1]','layout-A-teal',14,8,1.2,1.2,'26BFAF'),shape('/slidelayout[2]','layout-B-purple',14,8,1.2,1.2,'B273E0')]
for i,layout in enumerate(['Loop A','Loop A','Loop B'],1):
 ops.append(dict(command='add',parent='/',type='slide',props={'layout':layout,'background':'102236'}))
 ops += [shape(f'/slide[{i}]','local-title',2,1.2,29,1.2,'none',f'{layout} / PAGE {i}','28'),shape(f'/slide[{i}]','local-note',2,2.7,29,0.8,'none','Orange: master. Color square: layout. Blue: local click.','16'),shape(f'/slide[{i}]','local-blue',24,12,4,1.5,'3867A8','LOCAL','18')]
 ops.append(dict(command='add',parent=f'/slide[{i}]/shape[3]',type='animation',props={'effect':'fade','trigger':'onClick','duration':'400'}))
batch(deck,ops,'deck-batch');run(deck,'close')
ops=[]
for i,path in enumerate(['M 0 0 L 0.2 0 E','M 0 0 L 0 0.15 E','M 0 0 L 0 -0.15 E'],1):
 ops.append(dict(command='add',parent='/',type='slide',props={'layout':'blank','background':'102236'}));ops.append(shape(f'/slide[{i}]','timing-source',2,4.8,1.2,1.2,'E88735'))
 ops.append(dict(command='add',parent=f'/slide[{i}]/shape[1]',type='animation',props={'class':'motion','path':'custom','d':path,'trigger':'withPrevious','duration':'1000','autoReverse':'true','repeat':'indefinite'}))
batch(carrier,ops,'carrier-batch');run(carrier,'close')
with zipfile.ZipFile(deck) as dz,zipfile.ZipFile(carrier) as cz:
 mappings=[('ppt/slideMasters/slideMaster1.xml','master-orange','/slideMaster[1]','/p:sldMaster'),('ppt/slideLayouts/slideLayout1.xml','layout-A-teal','/slideLayout[1]','/p:sldLayout'),('ppt/slideLayouts/slideLayout2.xml','layout-B-purple','/slideLayout[2]','/p:sldLayout')]
 edits=[]
 for i,(part,name,target,xpath) in enumerate(mappings,1):
  r=E.fromstring(dz.read(part));node=next(v for v in r.findall('.//p:cNvPr',ns) if v.get('name')==name);targetid=node.get('id')
  timing=copy.deepcopy(E.fromstring(cz.read(f'ppt/slides/slide{i}.xml')).find('p:timing',ns));assert timing is not None
  effect=timing.find(".//p:cTn[@presetClass='path']",ns);assert effect is not None
  effect.set('repeatCount','indefinite');effect.set('autoRev','1')
  for t in timing.iter():
   if 'spid' in t.attrib:t.set('spid',targetid)
  edits.append((target,xpath,E.tostring(timing,encoding='unicode')))
for target,xpath,xml in edits:run(deck,'raw-set',target,'--xpath',xpath,'--action','append','--xml',xml)
run(deck,'close');(o/'validation.json').write_text(run(deck,'validate','--json'),encoding='utf-8');print(deck)
