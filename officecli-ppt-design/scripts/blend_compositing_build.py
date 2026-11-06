"""Passed blend subset: CLI native mult/over/screen, explicitly external SoftLight."""
import argparse,json,subprocess,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--fields',type=Path,required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():p.error('Fresh output directory required')
a.outdir.mkdir(parents=True);deck=a.outdir/'blend-compositing.pptx';log=[];A='http://schemas.openxmlformats.org/drawingml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
def cli(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace');log.append({'args':list(map(str,args)),'code':r.returncode,'stdout':r.stdout})
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
def text(s,name,x,y,w,h,value,size=17):
 args=['add',deck,f'/slide[{s}]','--type','shape']
 for k,v in {'name':name,'x':f'{x}cm','y':f'{y}cm','width':f'{w}cm','height':f'{h}cm','text':value,'font':'Arial','size':size,'color':'173C40','fill':'none','line':'none'}.items():args+=['--prop',f'{k}={v}']
 cli(*args)
def picture(s,name,src,x,y,opacity=None):
 args=['add',deck,f'/slide[{s}]','--type','picture','--prop','name='+name,'--prop','src='+str(a.fields/src),'--prop',f'x={x}cm','--prop',f'y={y}cm','--prop','width=9.2cm','--prop','height=6.9cm']
 if opacity is not None:args+=['--prop','opacity='+str(opacity)]
 cli(*args)
 return json.loads(cli('get',deck,f'/slide[{s}]/picture[@name={name}]','--json'))['data']['results'][0]['format']['relId']
def effect(s,name,mode,rid=None,alpha=None,solid=False):
 q=lambda x:'{'+A+'}'+x
 node=E.Element(q('fillOverlay'),blend=mode)
 if solid:E.SubElement(E.SubElement(node,q('solidFill')),q('srgbClr'),val='808080')
 else:
  fill=E.SubElement(node,q('blipFill'),rotWithShape='1');blip=E.SubElement(fill,q('blip'),{'{'+R+'}embed':rid})
  if alpha is not None:E.SubElement(blip,q('alphaModFix'),amt=str(alpha))
  E.SubElement(E.SubElement(fill,q('stretch')),q('fillRect'))
 cli('raw-set',deck,f'/slide[{s}]','--xpath',f"//p:pic[p:nvPicPr/p:cNvPr[@name='{name}']]/p:blipFill/a:blip",'--action','append','--xml',E.tostring(node,encoding='unicode'))
cli('create',deck,'--locale','en-US');cli('set',deck,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
panels=[]
for s in [1,2]:
 cli('add',deck,'/','--type','slide','--prop','layout=blank','--prop','background=F3F1EB');text(s,'title',1.5,0.8,31,1.3,'COLOR / '+('TRUE BLEND MODES' if s==1 else 'OPACITY & LOCAL LIGHT'),25)
 rid=picture(s,'source-cache','source.png',-20,-20);localrid=picture(s,'local-cache','local-gradient-rgba.png',-20,-20)
 for name in ['source-cache','local-cache']:cli('raw-set',deck,f'/slide[{s}]','--xpath',f"//p:pic/p:nvPicPr/p:cNvPr[@name='{name}']",'--action','setattr','--xml','hidden=1')
 cases=[('base','BACKDROP / original RGB',None,None),('source','SOURCE / original RGB',None,None),('softlight-ref','SOFT LIGHT / external sRGB formula',None,None),('image-mult','MULTIPLY / native','mult',None),('image-over','OVERLAY / native, opaque source','over',None),('image-screen','SCREEN / native','screen',None)] if s==1 else [('base','BACKDROP / original RGB',None,None),('normal50','NORMAL / opacity 50%',None,None),('mult50','MULTIPLY / native alpha 50%','mult','half'),('local-softlight','LOCAL SOFT LIGHT / external sRGB',None,None),('local-mult','LOCAL MULTIPLY / native RGBA','mult','local'),('local-overlay','LOCAL OVERLAY / external sRGB',None,None)]
 for idx,(name,label,mode,kind) in enumerate(cases):
  x=1.5+(idx%3)*10.4;y=3.0+(idx//3)*8.0;text(s,'label-'+name,x,y-0.65,9.3,0.65,label,12)
  src={'source':'source.png','softlight-ref':'softlight-encoded.png','local-softlight':'local-softlight-encoded.png','local-overlay':'local-overlay-encoded.png'}.get(name,'backdrop.png');picture(s,name,src,x,y)
  if name=='normal50':picture(s,'normal50-top','source.png',x,y,0.5)
  if mode:effect(s,name,mode,rid=localrid if kind=='local' else rid,alpha=50000 if kind=='half' else None,solid=kind=='solid')
  panels.append({'slide':s,'name':name,'rectCm':[x,y,9.2,6.9],'mode':mode,'kind':kind,'backend':'external-render' if name in ['softlight-ref','local-softlight','local-overlay'] else 'officecli-native','sourceAsset':src})
cli('close',deck);(a.outdir/'panels.json').write_text(json.dumps(panels,indent=2),encoding='utf-8');(a.outdir/'build-log.json').write_text(json.dumps(log,indent=2),encoding='utf-8');print(deck)
