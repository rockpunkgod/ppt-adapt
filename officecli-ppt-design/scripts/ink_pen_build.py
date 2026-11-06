"""Original native InkML plus an editable, tip-centered pencil; reuse Ink Replay tools."""
from pathlib import Path
import argparse,subprocess,sys,json,wave
HERE=Path(__file__).resolve().parent;sys.path.insert(0,str(HERE))
from build_components import Deck,W,H
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);d=Deck(a.outdir/'base.pptx',a.officecli);slide=d.slide('16251D')
d.shape(slide,'kicker',1.5,1.1,22,1,'TRACE / 笔尖与墨迹',font='Noto Sans SC',size=15,color='DAEF78')
d.shape(slide,'title',1.5,3,30,1.8,'每一笔，都跟得上。',font='Noto Serif SC',size=35,color='F2F3EC')
d.shape(slide,'footer',1.5,17,31,1,'单击慢放：5秒看清笔尖；下一页恢复1秒绘制。',font='Noto Sans SC',size=14,color='BCC7BB')
for name,x,text in [('nav-slow',27.4,'5秒'),('nav-fast',30,'1秒')]:d.shape(slide,name,x,1.1,2.2,.85,text,font='Noto Sans SC',size=14,color='BCC7BB',align='center')
for name,x,y,w,h,color in [('pen-pivot',5.15,11.15,3.2,3.2,'none'),('pen-body',6.87,11.35,.3,1.2,'F4CE6A'),('pen-nib',6.75,12.4,.42,.35,'D2A34C'),('tip-marker',6.68,12.68,.14,.14,'EF3C85')]:d.shape(slide,name,x,y,w,h,fill=color,geometry='ellipse' if name in ['pen-pivot','tip-marker'] else 'rect')
d.run('raw-set','/slide[1]','--xpath','//p:sp[p:nvSpPr/p:cNvPr[@name="pen-nib"]]/p:spPr/a:prstGeom','--action','replace','--xml','<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="w" b="h"/><a:pathLst><a:path w="420" h="350"><a:moveTo><a:pt x="0" y="350"/></a:moveTo><a:lnTo><a:pt x="120" y="0"/></a:lnTo><a:lnTo><a:pt x="420" y="150"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>')
d.run('add','/slide[1]','--type','group',props={'name':'tip-centered-pen','shapes':','.join('/slide[1]/shape[@name='+n+']' for n in ['pen-pivot','pen-body','pen-nib','tip-marker'])})
clock=a.outdir/'seed-clock.wav'
with wave.open(str(clock),'wb') as f:f.setnchannels(1);f.setsampwidth(2);f.setframerate(16000);f.writeframes(b'\0\0'*112000)
d.run('add','/slide[1]','--type','audio',props={'name':'audio','src':str(clock.resolve()),'x':'-2cm','y':'-2cm','width':'1cm','height':'1cm','autoPlay':True,'volume':0})
d.notes(slide,'Original native three-trace InkML. Tip-centered native pencil with no-fill/no-line circle. Slow diagnostic and one-second Replay/path demonstration. No bitmap or Wipe replacement.');d.run('close')
subprocess.run([sys.executable,str(HERE/'ink_pause_contentpart.py'),'--base',str(d.file),'--output',str(a.outdir/'with-ink.pptx')],check=True)
subprocess.run([sys.executable,str(HERE/'ink_pause_replay_seed.py'),'--base',str(a.outdir/'with-ink.pptx'),'--timing',str(HERE.parent/'assets/ink-replay-timing.xml'),'--output',str(a.outdir/'before-native.pptx')],check=True)
model=dict(path='M 0 0 '+' '.join(f'L {x/W:.9f} {y/H:.9f}' for x,y in [(7,-5),(14,0),(21,-5)])+' E',tipCenterCm=[6.75,12.75],traceCm=[[6.75,12.75],[13.75,7.75],[20.75,12.75],[27.75,7.75]])
(a.outdir/'model.json').write_text(json.dumps(model,indent=2),encoding='utf-8');print(a.outdir/'before-native.pptx')
