"""OfficeCLI creates original path-clipped strokes and genuine Multiply fields."""
import argparse,json,subprocess,xml.etree.ElementTree as E
from pathlib import Path
from stroke_depth_geometry import STROKES,BRIDGE,CURVED_LEFT,PATCHES,PATCH_RECEIVERS
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--assets',type=Path,required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():p.error('Fresh output required')
a.outdir.mkdir(parents=True);deck=a.outdir/'stroke-depth.pptx';A='http://schemas.openxmlformats.org/drawingml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';q=lambda s:'{'+A+'}'+s;plan=[];logs=[]
def cli(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace');logs.append({'args':list(map(str,args)),'code':r.returncode,'stdout':r.stdout})
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
def text(s,name,value,x,y,w=30,h=1,size=18):
 cli('add',deck,f'/slide[{s}]','--type','shape',*[v for k,vv in {'name':name,'text':value,'font':'Microsoft YaHei','size':size,'x':f'{x}cm','y':f'{y}cm','width':f'{w}cm','height':f'{h}cm','fill':'none','line':'none','color':'D7F2E8'}.items() for v in ['--prop',str(k)+'='+str(vv)]])
def picture(s,name,src,x,y,size):
 cli('add',deck,f'/slide[{s}]','--type','picture','--prop','name='+name,'--prop','src='+str(a.assets/src),'--prop',f'x={x}cm','--prop',f'y={y}cm','--prop',f'width={size}cm','--prop',f'height={size}cm')
 return json.loads(cli('get',deck,f'/slide[{s}]/picture[@name={name}]','--json'))['data']['results'][0]['format']['relId']
def geometry(name):
 root=E.Element(q('custGeom'))
 for tag in ['avLst','gdLst','ahLst','cxnLst']:E.SubElement(root,q(tag))
 E.SubElement(root,q('rect'),l='0',t='0',r='600000',b='600000');path=E.SubElement(E.SubElement(root,q('pathLst')),q('path'),w='600000',h='600000',fill='norm',stroke='false',extrusionOk='false')
 commands=CURVED_LEFT if name=='left' else [('M' if i==0 else 'L',point) for i,point in enumerate(BRIDGE if name=='bridge' else PATCHES[name] if name in PATCHES else STROKES[name])]+[('Z',())]
 for kind,coords in commands:
  if kind=='Z':E.SubElement(path,q('close'));continue
  node=E.SubElement(path,q({'M':'moveTo','L':'lnTo','C':'cubicBezTo'}[kind]))
  for i in range(0,len(coords),2):E.SubElement(node,q('pt'),x=str(round(coords[i]*1000)),y=str(round(coords[i+1]*1000)))
 return E.tostring(root,encoding='unicode')
def glyph(s,prefix,x,y,size,back,detail,bridge,colored=False):
 order=['left']+(['patch-lv'] if detail else [])+['right','vertical']+(['patch-vh','patch-vr'] if detail else [])+['horizontal']+(['bridge'] if bridge else []);items=[]
 for name in order:
  receiver=PATCH_RECEIVERS.get(name,name);base=(('right' if name=='bridge' else receiver)+'.png') if colored else 'green.png';shape=prefix+'-'+name;picture(s,shape,base,x,y,size)
  cli('raw-set',deck,f'/slide[{s}]','--xpath',f"//p:pic[p:nvPicPr/p:cNvPr[@name='{shape}']]/p:spPr/a:prstGeom",'--action','replace','--xml',geometry(name))
  patch=name in PATCHES;effect=back or patch
  source=f'{"back" if back else "front"}-{"detail" if patch else "flat"}-{name}.png'
  if effect:
   rid=picture(s,shape+'-source',source,-30,-30,1);cli('raw-set',deck,f'/slide[{s}]','--xpath',f"//p:pic[p:nvPicPr/p:cNvPr[@name='{shape}-source']]/p:nvPicPr/p:cNvPr",'--action','setattr','--xml','hidden=1')
   node=E.Element(q('fillOverlay'),blend='mult');fill=E.SubElement(node,q('blipFill'));E.SubElement(fill,q('blip'),{'{'+R+'}embed':rid});E.SubElement(E.SubElement(fill,q('stretch')),q('fillRect'));cli('raw-set',deck,f'/slide[{s}]','--xpath',f"//p:pic[p:nvPicPr/p:cNvPr[@name='{shape}']]/p:blipFill/a:blip",'--action','append','--xml',E.tostring(node,encoding='unicode'))
  items.append({'name':name,'shape':shape,'base':base,'effect':effect,'source':source if effect else None,'receiver':receiver,'closedPatch':patch})
 plan.append({'slide':s,'prefix':prefix,'rectCm':[x,y,size,size],'back':back,'detail':detail,'bridge':bridge,'colored':colored,'order':items})
cli('create',deck);cli('set',deck,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
for s,title in [(1,'01 / LETTER ORDER ONLY'),(2,'02 / LOCAL STROKE DEPTH'),(3,'03 / SAME GLOBAL ORDER, LOCAL CROSSING CHANGED')]:
 cli('add',deck,'/','--type','slide','--prop','layout=blank','--prop','background=101E25');text(s,'title',title,1.4,0.8,size=25);text(s,'original-text','木木  /  原创四笔路径 · 原文保留',1.4,17.1,size=16)
 if s<3:
  glyph(s,'back-word',14,4,12,True,s==2,s==2);glyph(s,'front-word',6.4,4,12,False,s==2,s==2)
  text(s,'caption','前字覆盖后字；'+('局部补片与交叉桥片关闭' if s==1 else '横在竖前 / 竖在撇前 / 捺局部绕到竖前'),1.4,2.1,size=16)
 else:
  glyph(s,'without-bridge',2.5,4.2,11.5,False,True,False,True);glyph(s,'with-bridge',18,4.2,11.5,False,True,True,True)
  text(s,'before-label','无局部桥片：竖在捺前',2.5,2.4,w=12,size=18);text(s,'after-label','局部桥片置顶 · 2单位接缝余量',18,2.4,w=14,size=18)
cli('close',deck);(a.outdir/'plan.json').write_text(json.dumps(plan,indent=2,ensure_ascii=False),encoding='utf-8');(a.outdir/'build-log.json').write_text(json.dumps(logs,ensure_ascii=False),encoding='utf-8');print(deck)
