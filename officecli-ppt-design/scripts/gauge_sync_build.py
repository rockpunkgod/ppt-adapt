"""One numeric model drives a native doughnut, true pivot and digit strips."""
import argparse,json,subprocess,zipfile,copy,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);p.add_argument('--values',default='25,60');a=p.parse_args();o=Path(a.out);values=[int(v) for v in a.values.split(',')];assert len(values)==2 and all(1<=v<=99 for v in values),'Two integer values1..99, default tested25/60'
if o.exists():raise FileExistsError('Refusing overwrite')
o.mkdir(parents=True);W,H=33.87,19.05;cx,cy=16.935,10.;pitchPt=45.35;pitch=pitchPt/72*2.54;BG='F6F8FC';P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';C='http://schemas.openxmlformats.org/drawingml/2006/chart';N={'p':P,'a':A,'c':C}
E.register_namespace('p',P);E.register_namespace('a',A);E.register_namespace('c',C)
model=[{'value':v,'visible':v,'invisible':200-v,'needleDegrees':1.8*v,'tens':v//10,'ones':v%10,'digitPitchCm':pitch,'endSeconds':2} for v in values]
(o/'model.json').write_text(json.dumps(model,indent=2),encoding='utf-8')
def run(file,*args):
 r=subprocess.run([a.officecli,args[0],str(file),*map(str,args[1:])],capture_output=True,text=True,encoding='utf-8');assert r.returncode==0,r.stdout+r.stderr;return r.stdout
def geom(lower):
 # Two opposite arc directions form a half-annulus, not a rectangle over the hole.
 if lower:outer=(100200,50000,0,10800000);inner=(10200,50000,10800000,-10800000)
 else:outer=(0,50000,10800000,10800000);inner=(90000,50000,0,-10800000)
 ox,oy,st,sw=outer;ix,iy,ist,isw=inner;ro=50200 if lower else 50000;ri=39800 if lower else 40000
 return f'<a:custGeom xmlns:a="{A}"><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100000" h="100000"><a:moveTo><a:pt x="{ox}" y="{oy}"/></a:moveTo><a:arcTo wR="{ro}" hR="{ro}" stAng="{st}" swAng="{sw}"/><a:lnTo><a:pt x="{ix}" y="{iy}"/></a:lnTo><a:arcTo wR="{ri}" hR="{ri}" stAng="{ist}" swAng="{isw}"/><a:close/></a:path></a:pathLst></a:custGeom>'
for static in [False,True]:
 file=o/('final-reference.pptx' if static else 'gauge.pptx');run(file,'create');run(file,'set','/','--prop',f'slideWidth={W}cm','--prop',f'slideHeight={H}cm');ops=[]
 def shape(s,name,x,y,w,h,fill='none',text='',size=18,**extra):
  props=dict(name=name,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',fill=fill,line='none',text=text,size=str(size),font='Arial',color='17324E',margin='0cm');props.update({k:str(v) for k,v in extra.items()});ops.append(dict(command='add',parent=f'/slide[{s}]',type='shape',props=props))
 for i,m in enumerate(model,1):
  ops.append(dict(command='add',parent='/',type='slide',props={'layout':'blank','background':BG}))
  shape(i,'title',1.4,1.5,7.5,3,text='Product\nadoption',size=30,bold='true')
  shape(i,'subtitle',1.5,5,6.5,2,text='Active accounts\nScale 0–100%',size=17)
  shape(i,'footer',1.5,17.7,29,0.8,text='NOVA / PRODUCT INSIGHTS',size=12)
  # Single editable textbox per digit column; fixed paragraph pitch.
  for name,x,digit,seq in [('tens',15.3,m['tens'],list(range(10))),('ones',16.45,m['ones'],list(range(10))+[0])]:
   shape(i,name,x,11.2-(digit*pitch if static else 0),1.1,len(seq)*pitch,text='\n'.join(map(str,seq)),size=40,bold='true',align='center',lineSpacing=f'{pitchPt}pt',spaceBefore='0pt',spaceAfter='0pt',valign='top')
  shape(i,'number-cover-top',15.05,0,2.8,11.2,fill=BG);shape(i,'number-cover-bottom',15.05,11.2+pitch,2.8,H-11.2-pitch,fill=BG)
  shape(i,'remaining-track',cx-8.5,cy-8.5,17,17,fill='E2E7EE')
  ops.append(dict(command='add',parent=f'/slide[{i}]',type='chart',props=dict(name='gauge-data',chartType='doughnut',categories='Visible,Invisible',data=f'Gauge:{m["visible"]},{m["invisible"]}',x=f'{cx-8.5}cm',y=f'{cy-8.5}cm',width='17cm',height='17cm',firstSliceAngle='270' if static else '90',holeSize='80',colorrule='100:26BFAF:none',legend='none',title='none',chartareafill='none',chartborder='none',plotareafill='none',plotborder='none',seriesoutline='none')))
  shape(i,'chart-cover-bottom',cx-8.5,cy-8.5-0.02,17,17,fill=BG)
  shape(i,'needle-envelope',cx-7,cy-7,14,14,geometry='ellipse')
  shape(i,'needle-tip',cx-6.1,cy-0.14,6.1,0.28,fill='F26F4E')
  shape(i,'needle-hub',cx-.45,cy-.45,.9,.9,fill='F26F4E',geometry='ellipse')
  shape(i,'percent',17.75,11.5,1.2,0.8,text='%',size=23)
  shape(i,'unit',13,14,9,1,text='OF ACTIVE ACCOUNTS',size=15)
  if not static:
   for target,props in [('/chart[1]',{'effect':'spin','class':'emphasis','duration':'2000','trigger':'onClick'}),('/shape[11]',{'effect':'spin','class':'emphasis','duration':'2000','trigger':'withPrevious'}),('/shape[4]',{'class':'motion','path':'custom','d':f'M 0 0 L 0 {-m["tens"]*pitch/H} E','duration':'2000','trigger':'withPrevious'}),('/shape[5]',{'class':'motion','path':'custom','d':f'M 0 0 L 0 {-10*pitch/H} E','duration':'250','trigger':'withPrevious'}),('/shape[5]',{'class':'motion','path':'custom','d':f'M 0 0 L 0 {-m["ones"]*pitch/H} E','duration':'1000','delay':'1000','trigger':'withPrevious'})]:
    ops.append(dict(command='add',parent=f'/slide[{i}]'+target,type='animation',props=props))
 bf=o/('reference-batch.json' if static else 'gauge-batch.json');bf.write_text(json.dumps(ops),encoding='utf-8');r=run(file,'batch','--input',bf,'--json');assert json.loads(r)['success'];run(file,'close')
 # Locate actual chart relationship parts. CLI raw-set accepts explicit OPC part paths.
 with zipfile.ZipFile(file) as z:
  chartparts=[f for f in z.namelist() if '/charts/' in f and f.endswith('.xml')];assert len(chartparts)==2
  slides=[E.fromstring(z.read(f'ppt/slides/slide{i}.xml')) for i in [1,2]]
 for part in chartparts:
  layout=f'<c:layout xmlns:c="{C}"><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="factor"/><c:yMode val="factor"/><c:wMode val="factor"/><c:hMode val="factor"/><c:x val="0"/><c:y val="0"/><c:w val="1"/><c:h val="1"/></c:manualLayout></c:layout>'
  run(file,'raw-set','/'+part,'--xpath','//c:plotArea/c:layout','--action','replace','--xml',layout)
 for i,(m,r) in enumerate(zip(model,slides),1):
  for name,lower in [('remaining-track',False),('chart-cover-bottom',True)]:run(file,'raw-set',f'/slide[{i}]','--xpath',f"//p:sp[p:nvSpPr/p:cNvPr[@name='{name}']]/p:spPr/a:prstGeom",'--action','replace','--xml',geom(lower))
  triangle=f'<a:custGeom xmlns:a="{A}"><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100000" h="100000"><a:moveTo><a:pt x="0" y="50000"/></a:moveTo><a:lnTo><a:pt x="100000" y="0"/></a:lnTo><a:lnTo><a:pt x="100000" y="100000"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>'
  run(file,'raw-set',f'/slide[{i}]','--xpath',"//p:sp[p:nvSpPr/p:cNvPr[@name='needle-tip']]/p:spPr/a:prstGeom",'--action','replace','--xml',triangle)
  names={v.find('p:nvSpPr/p:cNvPr',N).get('name'):v.find('p:nvSpPr/p:cNvPr',N).get('id') for v in r.findall('p:cSld/p:spTree/p:sp',N)}
  ids=[names[q] for q in ['needle-envelope','needle-tip','needle-hub']]
  run(file,'add',f'/slide[{i}]','--type','group','--prop','name=needle-pivot','--prop','shapes='+','.join(f'/slide[{i}]/shape[@id={v}]' for v in ids))
 run(file,'close')
 with zipfile.ZipFile(file) as z:slides=[E.fromstring(z.read(f'ppt/slides/slide{i}.xml')) for i in [1,2]]
 for i,(m,r) in enumerate(zip(model,slides),1):
  group=r.find('p:cSld/p:spTree/p:grpSp',N);gid=group.find('p:nvGrpSpPr/p:cNvPr',N).get('id')
  if static:run(file,'raw-set',f'/slide[{i}]','--xpath','//p:grpSp/p:grpSpPr/a:xfrm','--action','setattr','--xml','rot='+str(round(m['needleDegrees']*60000)))
  else:
   timing=r.find('p:timing',N);targets={v.find('p:nvSpPr/p:cNvPr',N).get('name'):v.find('p:nvSpPr/p:cNvPr',N).get('id') for v in r.findall('.//p:sp',N)};tip=targets['needle-tip'];ones=targets['ones'];tens=targets['tens'];oneCount=0
   for v in timing.iter():
    if v.get('spid')==tip:v.set('spid',gid)
   for e in timing.findall('.//p:cTn[@presetClass]',N):
    target=e.find('.//p:spTgt',N).get('spid');rotation=e.find('.//p:animRot',N)
    if rotation is not None:rotation.set('by',str(round((m['needleDegrees'] if target==gid else 180)*60000)));e.set('decel','50000');e.set('accel','37500' if target==gid else '0')
    else:
     if target==ones:
      oneCount+=1
      if oneCount==1:e.set('repeatCount','4000');e.set('fill','remove')
      else:
       e.set('decel','100000');parents={c:p for p in timing.iter() for c in p};cur=parents.get(e)
       while cur is not None:
        condition=cur.find('p:stCondLst/p:cond',N)
        if condition is not None and condition.get('delay')=='1000':condition.set('delay','0')
        cur=parents.get(cur)
       e.find('p:stCondLst/p:cond',N).set('delay','1000')
     elif target==tens:e.set('decel','50000')
     child=e.find('.//p:animMotion/p:cBhvr/p:cTn',N)
     if child is not None:child.set('fill','remove' if target==ones and oneCount==1 else 'hold')
   run(file,'raw-set',f'/slide[{i}]','--xpath','/p:sld/p:timing','--action','replace','--xml',E.tostring(timing,encoding='unicode'))
 run(file,'close');(o/(file.stem+'-validation.json')).write_text(run(file,'validate','--json'),encoding='utf-8')
print(o/'gauge.pptx')
