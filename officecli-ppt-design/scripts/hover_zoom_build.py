"""Two branches with real holes and native Slide Zoom; OfficeCLI foundation."""
import argparse,subprocess,json,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():p.error('Fresh output directory required')
a.outdir.mkdir(parents=True);file=a.outdir/'base.pptx';log=[]
def cli(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace');log.append({'args':list(map(str,args)),'code':r.returncode,'stdout':r.stdout})
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
def shape(s,name,x,y,w,h,text='',**kw):
 props={'name':'!!'+name,'x':f'{x}cm','y':f'{y}cm','width':f'{w}cm','height':f'{h}cm','text':text,'font':'Arial','size':'22','color':'183D40','fill':'none','line':'none','margin':'0cm','valign':'middle'};props.update({k:str(v).lower() if isinstance(v,bool) else str(v) for k,v in kw.items()})
 args=['add',file,f'/slide[{s}]','--type','shape']
 for k,v in props.items():args+=['--prop',f'{k}={v}']
 cli(*args)
cli('create',file,'--locale','en-US');cli('set',file,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
for i in range(5):cli('add',file,'/','--type','slide','--prop','layout=blank','--prop','background=F3F1EB')
for s in range(1,4):
 active=s-1
 for j,x in [(1,3),(2,18.3)]:
  selected=active==j
  shape(s,'card-'+str(j),x,5.5,12.3,10.1,fill='D5E2DB' if not selected else 'BCD9D0')
  shape(s,'icon-'+str(j),x+1.0,6.2 if selected else 8.0,10.3,2.0 if selected else 3.6,'A' if j==1 else 'B',size=30 if selected else 78,bold=True,align='center')
  shape(s,'label-'+str(j),x+0.6,14.5 if selected else 16.5,11.1,0.8,'RESEARCH' if j==1 else 'DELIVERY',size=17,align='center')
  cli('add',file,f'/slide[{s}]','--type','zoom','--prop','name=!!zoom-'+str(j),'--prop','target='+str(j+3),'--prop',f'x={x+1.0}cm','--prop',f'y={8.5 if selected else 16.5}cm','--prop','width=10.3cm','--prop','height=5.79375cm','--prop','returnToParent=true','--prop','transitionDur=500')
 # One native editable custom geometry, outer perimeter + opposite-winding apertures.
 shape(s,'return-mask',0,0,33.87,19.05,fill='F3F1EB')
 ns='http://schemas.openxmlformats.org/drawingml/2006/main';q=lambda k:'{'+ns+'}'+k
 cg=E.Element(q('custGeom'));[E.SubElement(cg,q(k)) for k in ['avLst','gdLst','ahLst','cxnLst']];E.SubElement(cg,q('rect'),l='0',t='0',r='r',b='b');lst=E.SubElement(cg,q('pathLst'));path=E.SubElement(lst,q('path'),w='338700',h='190500',fill='norm',stroke='0')
 for pts in [[(0,0),(338700,0),(338700,190500),(0,190500)],[(30000,55000),(30000,156000),(153000,156000),(153000,55000)],[(183000,55000),(183000,156000),(306000,156000),(306000,55000)]]:
  for k,(x,y) in enumerate(pts):E.SubElement(E.SubElement(path,q('moveTo' if k==0 else 'lnTo')),q('pt'),x=str(x),y=str(y))
  E.SubElement(path,q('close'))
 cli('raw-set',file,f'/slide[{s}]','--xpath',"//p:sp[p:nvSpPr/p:cNvPr[@name='!!return-mask']]/p:spPr/a:prstGeom",'--action','replace','--xml',E.tostring(cg,encoding='unicode'))
 shape(s,'title',3,1.5,28,1.8,'TWO WAYS TO MAKE PROGRESS',size=31,bold=True)
 shape(s,'hint',3,3.4,28,0.9,'Hover to reveal. Click the preview to explore.',size=17)
 cli('set',file,f'/slide[{s}]','--prop','transition=morph')
for j in [1,2]:
 s=j+3;shape(s,'detail-title',3,1.7,28,2,'RESEARCH' if j==1 else 'DELIVERY',size=42,bold=True)
 shape(s,'detail-subtitle',3,4.1,28,1.2,'Observe the evidence.' if j==1 else 'Make the result dependable.',size=24)
 for k in range(3):shape(s,'detail-node-'+str(k),3+k*9.7,7.7,8.2,6.7,['ASK','COMPARE','DECIDE'][k] if j==1 else ['BUILD','CHECK','SHIP'][k],fill=['D9B190','ACB2D0','A8CDBF'][k],size=24,bold=True,align='center')
 shape(s,'detail-return',3,16.0,28,1,'Click blank space to return to the same preview.',size=17)
cli('close',file);(a.outdir/'build-log.json').write_text(json.dumps(log,indent=2),encoding='utf-8');print(file)
