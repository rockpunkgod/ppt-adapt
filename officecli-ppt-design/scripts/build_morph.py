"""OfficeCLI reproductions of three tutorial mechanisms; no python-pptx.

The 3D test edits the author's public tutorial file and is labelled as such.
It is not a claim that OfficeCLI can synthesize stock-model animation scenes.
"""
import argparse, json, shutil, subprocess, sys, zipfile, hashlib
from pathlib import Path

BASE=Path(__file__).resolve().parent
W,H=33.87,19.05

class Deck:
 def __init__(self,cli,path):
  self.cli=str(cli);self.path=Path(path);self.ops=[];self.slide=0
 def call(self,*args):
  p=subprocess.run([self.cli,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace')
  if p.returncode: raise RuntimeError(p.stdout+'\n'+p.stderr)
  return p.stdout
 def start(self):
  if self.path.exists():raise FileExistsError(str(self.path)+' exists; choose a new output directory')
  self.call('create',self.path);self.call('open',self.path)
  self.call('set',self.path,'/','--prop',f'slideWidth={W}cm','--prop',f'slideHeight={H}cm')
 def batch(self):
  if not self.ops:return
  p=self.path.with_suffix('.batch.json');p.write_text(json.dumps(self.ops,ensure_ascii=False),encoding='utf-8')
  out=self.call('batch',self.path,'--input',p,'--json')
  data=json.loads(out)
  if '"success": false' in out or '"atomicRolledBack": true' in out:raise RuntimeError(out)
  self.ops=[]
 def add(self,parent,typ,**props):
  self.ops.append({'command':'add','parent':parent,'type':typ,'props':{k:str(v).lower() if isinstance(v,bool) else str(v) for k,v in props.items()}})
 def new(self,bg='000000',transition=None):
  self.batch();self.slide+=1
  props=dict(layout='blank',background=bg)
  if transition:props['transition']=transition
  self.add('/','slide',**props);self.batch();return self.slide
 def shape(self,name,x,y,w,h,fill='none',text='',size=22,color='FFFFFF',**props):
  defaults=dict(name=name,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',fill=fill,line='none',text=text,size=size,color=color,font='Arial',margin='0cm',valign='center');defaults.update(props)
  self.add(f'/slide[{self.slide}]','shape',**defaults)
 def picture(self,name,src,x,y,w,h,**props):
  self.add(f'/slide[{self.slide}]','picture',name=name,src=src,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',**props)
 def polygon(self,name,x,y,w,h,points,fill='FFFFFF',opacity=1,gradient=None):
  self.shape(name,x,y,w,h,fill=fill,opacity=opacity)
  self.batch()
  p=''.join((f'<a:moveTo><a:pt x="{a}" y="{b}"/></a:moveTo>' if i==0 else f'<a:lnTo><a:pt x="{a}" y="{b}"/></a:lnTo>') for i,(a,b) in enumerate(points))
  xml=f'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100000" h="100000">{p}<a:close/></a:path></a:pathLst></a:custGeom>'
  xp=f"//p:sp[p:nvSpPr/p:cNvPr[@name='{name}']]/p:spPr/a:prstGeom"
  self.call('raw-set',self.path,f'/slide[{self.slide}]','--xpath',xp,'--action','replace','--xml',xml)
  if gradient:
   xml='<a:gradFill rotWithShape="1"><a:gsLst>'+''.join(f'<a:gs pos="{pos}"><a:srgbClr val="{col}"><a:alpha val="{alpha}"/></a:srgbClr></a:gs>' for pos,col,alpha in gradient)+'</a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>'
   self.call('raw-set',self.path,f'/slide[{self.slide}]','--xpath',f"//p:sp[p:nvSpPr/p:cNvPr[@name='{name}']]/p:spPr/a:solidFill",'--action','replace','--xml',xml)
 def notes(self,text):self.add(f'/slide[{self.slide}]','notes',text=text)
 def finish(self):
  self.batch();self.call('close',self.path)
  valid=self.call('validate',self.path,'--json');(self.path.with_suffix('.validation.json')).write_text(valid,encoding='utf-8')
  html=self.path.with_suffix('.html');self.call('view',self.path,'html','-o',html)
  print(json.dumps({'artifact':str(self.path),'html':str(html),'validation':json.loads(valid)},ensure_ascii=False))

def modern(cli,out,landscape,portrait):
 d=Deck(cli,out/'modern-morph.pptx');d.start()
 for pair,src in enumerate([landscape,portrait]):
  for state in range(2):
   d.new('FFFFFF','morph-1800' if state else ('fade-400' if pair else None))
   # The photo is oversized on the second state but clipped by four native masks.
   if pair==0:
    x,y,w,h=(.4143,.8,33.0414,17.45) if not state else (-1.6,-.7,36.0,19.0125)
   else:
    x,y,w,h=(.8,-10.0,32.27,48.405) if not state else (.8,-24.0,32.27,48.405)
   d.picture('!!photo-'+str(pair),src,x,y,w,h)
   # Main reading area and contiguous bevels reproduce the split-mask mechanism.
   shift=-36 if state==0 else 0
   d.polygon('!!panel-main',.8+shift,.8,20,17.45,[(0,0),(100000,0),(66000,100000),(0,100000)],opacity=.86)
   d.polygon('!!panel-bevel1',14+shift,.8,8,17.45,[(60000,0),(100000,0),(40000,100000),(0,100000)],fill='000000',gradient=[(0,'000000',75000),(100000,'FFFFFF',10000)])
   d.polygon('!!panel-bevel2',17.2+shift,.8,8,17.45,[(60000,0),(100000,0),(40000,100000),(0,100000)],fill='000000',gradient=[(0,'000000',40000),(100000,'FFFFFF',5000)])
   for i,(label,value) in enumerate([('PHOTO STUDY','Light and reflection' if pair==0 else 'A path into the landscape'),('SOURCE','Unsplash / source listed in notes'),('METHOD','Fixed frame / moving image')]):
    d.shape(f'!!meta-{i}',2.2+shift,2.3+i*3.4,10.8,1.0,text=label,size=14,color='242424',bold=True,spacing=2)
    d.shape(f'!!value-{i}',2.2+shift,3.25+i*3.4,10.8,1.4,text=value,size=24 if i==0 else 18,color='111111',bold=i==0)
   d.shape('!!reticle-box',24.8,8.0,3.5 if state else .2,3.5 if state else .2,line='FFFFFF:1.5',lineOpacity=.3 if state else 0,fill='none')
   d.shape('!!reticle-h',24.0 if state else 36,9.4,1.5,.06,fill='FFFFFF')
   d.shape('!!reticle-v',24.72,8.7 if state else -5,.06,1.5,fill='FFFFFF')
   d.shape('!!caption',27 if state else 36,11.8,4,1.2,text='FRAME / 01',size=14,color='FFFFFF')
   # Fixed white frame covers the oversized photo without changing its crop window.
   for name,x,y,w,h in [('top',0,0,W,.8),('bottom',0,H-.8,W,.8),('left',0,0,.8,H),('right',W-.8,0,.8,H)]:d.shape('!!border-'+name,x,y,w,h,fill='FFFFFF')
   d.notes('Reconstructed from One Skill PPT Modern Morph tutorial 3wtwVE5injI, 02:44-06:16 and 11:19-14:08. Photos are two of the tutorial source Unsplash URLs; metadata is intentionally not fabricated. The movement is an implementation adaptation; runtime requires a Morph-capable player.')
 d.finish()

def cinematic(cli,out,landscape):
 d=Deck(cli,out/'cinematic-morph.pptx');d.start()
 # Cover -> image reveal. Cloud assets are omitted; this tests the reveal mechanism.
 for state in range(2):
  d.new('000000','morph-2000' if state else None)
  d.picture('!!cinema-photo',landscape,-6.1576,0 if not state else -5,46.1852,24.39,opacity=.3)
  for n,(word,y) in enumerate([('THE',4.0),('WATER',6.0),('AND LIGHT',8.0)]):
   d.shape('!!hero-'+str(n),2.5,y if not state else y-22,28.8,2.0,text=word,size=48 if n else 24,color='E8DEC5',font='Georgia',align='center',spacing=4)
  d.shape('!!reveal-copy',6,22 if not state else 10.8,21.8,4.0,text='One landscape, two views.\nA fixed viewport reveals a new part of the same image.',size=22,color='EFE5D2',align='center')
  d.notes('Luis Urrutia aQlTPUvVH5E, 00:13-04:23. Background starts at 30% opacity (70% transparency), title moves upward, next copy starts below the slide; second state uses Morph. Cloud sprites omitted intentionally, so atmospheric-asset match is not claimed.')
 # Two simultaneous event layers move in opposing directions, like the pro timeline.
 events=[('2019','The first view','Start with a quiet composition.'),('2024','A different perspective','Move the scene; preserve a clear reading order.')]
 for state in range(2):
  d.new('000000','morph-1800' if state else 'zoom-800')
  for j,(year,title,body) in enumerate(events):
   bgx=(j-state)*W
   d.picture(f'!!event-photo-{j}',landscape,bgx,0,W,H,opacity=.25,crop='3.05,0,3.05,0')
   d.shape(f'!!event-year-{j}',bgx+5,3.0,23.8,1.5,text=year,size=26,color='E8DEC5',align='center')
   d.shape(f'!!event-point-{j}',bgx+16.5,5.3,.5,.5,fill='E8DEC5',preset='ellipse')
   # Incoming copy begins left; outgoing copy leaves right: opposite to photos/date.
   copyx=(state-j)*W+5
   d.shape(f'!!event-title-{j}',copyx,7.0,23.8,2.5,text=title,size=38,color='E8DEC5',align='center',font='Georgia')
   d.shape(f'!!event-copy-{j}',copyx+2,10.1,19.8,2.0,text=body,size=22,color='FFFFFF',align='center')
  d.shape('!!timeline',0,5.54,W,.04,fill='E8DEC5')
  d.notes('Pro timeline from aQlTPUvVH5E 04:29-09:04: date, image and timeline move left; title/description move right. Timeline example dates are illustrative, not historical claims. Asset is a landscape rather than Roman history imagery.')
 # Editable, bounded-cost option: independent complete pages + Push from right.
 for i,(year,title,body) in enumerate(events):
  d.new('000000','push-right-900' if i else 'fade-400')
  d.picture('simple-bg',landscape,0,0,W,H,opacity=.25,crop='3.05,0,3.05,0')
  d.shape('simple-date',5,3,23.8,1.5,text=year,size=26,color='E8DEC5',align='center')
  d.shape('simple-title',5,7,23.8,2.5,text=title,size=38,color='E8DEC5',font='Georgia',align='center')
  d.shape('simple-copy',7,10.1,19.8,2,text=body,size=22,color='FFFFFF',align='center')
  d.shape('simple-axis',0,5.54,W,.04,fill='E8DEC5')
  d.shape('simple-point',16.5,5.3,.5,.5,fill='E8DEC5',preset='ellipse')
  d.notes('Simplified timeline from aQlTPUvVH5E 09:25-10:23: no offscreen duplicates, complete current-event page, Push from right. This trades individualized parallax for easier maintenance.')
 d.finish()

def template3d(cli,out,source):
 path=out/'3d-source-adaptation.pptx'
 if path.exists():raise FileExistsError(path)
 shutil.copy2(source,path)
 d=Deck(cli,path);d.call('open',path)
 replacements={'YOUR':'ORBITAL','AWESOME':'FLIGHT','TITLE':'STUDY','CINEMATIC':'CONTINUOUS','MORPH':'SCENE','TRANSITIONS':'MOTION','3D':'3D','MODEL':'MODEL','ANIMATIONS':'SCENES','SLIDE':'CLICK','SCROLL':'AND','MENU':'EXPLORE','Lorem ipsum is placeholder text commonly used in the graphic, print, and publishing industries for previewing layouts and visual mockups.':'A source-template adaptation tests preservation of native 3D scenes, cross-slide motion and linked navigation.'}
 for old,new in replacements.items():
  if old!=new:d.call('set',path,'/','--find',old,'--replace',new)
 for i in range(1,5):
  # Normalize an imported template's font rather than trusting fallback metrics.
  d.call('raw-set',path,f'/slide[{i}]','--xpath','//a:latin','--action','setattr','--xml','typeface=Arial')
  d.call('raw-set',path,f'/slide[{i}]','--xpath','//a:rPr','--action','setattr','--xml','kern=0')
  # A noFill ellipse only offers its thin outline to pointer hit-testing.
  # Give the linked circles a transparent native fill without changing their look.
  d.call('raw-set',path,f'/slide[{i}]','--xpath','//p:sp[p:nvSpPr/p:cNvPr/a:hlinkClick]/p:spPr/a:noFill','--action','replace','--xml','<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:solidFill>')
  d.call('add',path,f'/slide[{i}]','--type','notes','--prop','text=Adapted with OfficeCLI from One Skill PPT public tutorial source: https://www.patreon.com/posts/68201665. Preserves the original stock animated astronaut and native animation scene data; not generated from scratch. Original artwork and model belong to their respective rights holders. Do not bundle this source/model in a public skill package.')
 d.call('close',path)
 d.call('view',path,'html','-o',path.with_suffix('.html'))
 try:validation=d.call('validate',path,'--json')
 except RuntimeError as e:validation=str(e)
 path.with_suffix('.validation.json').write_text(validation,encoding='utf-8')
 with zipfile.ZipFile(source) as a,zipfile.ZipFile(path) as b:
  media=[n for n in a.namelist() if n.endswith('.glb')]
  result={n:hashlib.sha256(a.read(n)).hexdigest()==hashlib.sha256(b.read(n)).hexdigest() for n in media}
 (out/'3d-media-preservation.json').write_text(json.dumps(result,indent=2),encoding='utf-8')
 print(json.dumps({'artifact':str(path),'native_glb_preserved':result},ensure_ascii=False))

def main():
 p=argparse.ArgumentParser();p.add_argument('--officecli',default='officecli');p.add_argument('--mode',choices=['modern','cinematic','3d-template','all'],default='all');p.add_argument('--outdir',type=Path,required=True);p.add_argument('--landscape',type=Path);p.add_argument('--portrait',type=Path);p.add_argument('--source-3d',type=Path);a=p.parse_args();a.outdir.mkdir(parents=True,exist_ok=True)
 if a.mode in ('modern','cinematic','all') and a.landscape is None:p.error('--landscape is required for this mode')
 if a.mode in ('modern','all') and a.portrait is None:p.error('--portrait is required for this mode')
 if a.mode in ('3d-template','all') and a.source_3d is None:p.error('--source-3d is required: provide an authorized local source template')
 if a.mode in ('modern','all'):modern(a.officecli,a.outdir,a.landscape.resolve(),a.portrait.resolve())
 if a.mode in ('cinematic','all'):cinematic(a.officecli,a.outdir,a.landscape.resolve())
 if a.mode in ('3d-template','all'):template3d(a.officecli,a.outdir,a.source_3d.resolve())
if __name__=='__main__':main()
