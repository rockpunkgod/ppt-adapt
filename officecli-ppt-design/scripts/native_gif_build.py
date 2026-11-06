"""Native vector animation source, or a two-page native GIF texture example."""
from pathlib import Path
import argparse,sys,xml.etree.ElementTree as E
sys.path.insert(0,str(Path(__file__).resolve().parent))
from build_components import Deck
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);p.add_argument('--stage',choices=['source','texture'],required=True);p.add_argument('--gif',type=Path);a=p.parse_args()
if a.stage=='texture' and (not a.gif or a.gif.read_bytes()[:6] not in [b'GIF87a',b'GIF89a']):raise ValueError('Use the native exported GIF')
d=Deck(a.out,a.officecli)
def shape(s,name,x,y,w,h,text='',**kw):
 props=dict(name=name,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',text=text,fill='none',line='none',font='Noto Sans SC',size='18',color='F2F3EC',margin='0cm');props.update(kw)
 return dict(command='add',parent=f'/slide[{s}]',type='shape',props=props)
if a.stage=='source':
 d.run('set','/',props={'slideWidth':'16cm','slideHeight':'9cm'});s=d.slide('101C2A');ops=[]
 for i,z in enumerate([1.2,2.1,3.1,4.2,5.5]):ops.append(shape(s,'square-'+str(i),8-z/2,4.5-z/2,z,z,geometry='rect',line='FFFFFF',lineWidth='3pt'))
 ops.append(shape(s,'phase-marker',10.8,4.3,.4,.4,geometry='ellipse',fill='39D5C1'));d.batch(ops)
 d.notes(s,'Original native vector texture source. Native Spin90/Grow130%, 2s auto-reverse; export only this slide as a transparent animated GIF. No input bitmap/video/GIF.')
else:
 s=d.slide('16251D');d.slide('16251D');d.batch([shape(s,'kicker',1.5,1.1,30,1,'TRACE / 原生纹理',size='15',color='DAEF78'),shape(s,'title',1.5,5,30,5,'纹理，也可以\n有自己的节奏。',size='42',font='Noto Serif SC'),shape(s,'open-button',13,12,8,2,'查看动态纹理',fill='DAEF78',color='16251D',align='center',valign='center'),shape(s,'open-hit',13,12,8,2,fill='FFFFFF',opacity='0',link='slide[2]')])
 s=2;d.batch([shape(s,'kicker',1.5,1.1,30,1,'TRACE / 原生纹理',size='15',color='DAEF78'),shape(s,'title',1.5,2.3,31,1.5,'让静态轮廓，容纳流动。',size='32',font='Noto Serif SC'),shape(s,'bare-label',1.5,4.5,14,1,'原生透明 GIF',size='16',color='BCC7BB'),shape(s,'clip-label',18,4.5,14,1,'同一 GIF，真实相交',size='16',color='BCC7BB'),shape(s,'footer',1.5,16.7,31,1,'同一段原生动画，同一份 GIF。轮廓不动，纹理循环。',size='14',color='BCC7BB'),shape(s,'vase-base',18,6,14,7.875,fill='4D6E66'),dict(command='add',parent='/slide[2]',type='picture',props=dict(name='bare-gif',src=str(a.gif.resolve()),x='1.5cm',y='6cm',width='14cm',height='7.875cm')),dict(command='add',parent='/slide[2]',type='picture',props=dict(name='texture-gif',src=str(a.gif.resolve()),x='18cm',y='6cm',width='14cm',height='7.875cm')),shape(s,'vase-cut',18,6,14,7.875,fill='DAEF78')])
 A='http://schemas.openxmlformats.org/drawingml/2006/main';q=lambda k:'{'+A+'}'+k;cg=E.Element(q('custGeom'))
 for k in ['avLst','gdLst','ahLst','cxnLst']:E.SubElement(cg,q(k))
 E.SubElement(cg,q('rect'),l='0',t='0',r='w',b='h');lst=E.SubElement(cg,q('pathLst'));path=E.SubElement(lst,q('path'),w='1000000',h='1000000')
 for kind,pts in [('moveTo',[(.42,.1)]),('lnTo',[(.58,.1)]),('cubicBezTo',[(.58,.24),(.56,.29),(.61,.39)]),('cubicBezTo',[(.77,.58),(.74,.83),(.62,.9)]),('cubicBezTo',[(.56,.95),(.44,.95),(.38,.9)]),('cubicBezTo',[(.26,.83),(.23,.58),(.39,.39)]),('cubicBezTo',[(.44,.29),(.42,.24),(.42,.1)]),('close',[])]:
  cmd=E.SubElement(path,q(kind))
  for x,y in pts:E.SubElement(cmd,q('pt'),x=str(round(x*1e6)),y=str(round(y*1e6)))
 for name in ['vase-base','vase-cut']:d.run('raw-set','/slide[2]','--xpath',f"//p:sp[p:nvSpPr/p:cNvPr[@name='{name}']]/p:spPr/a:prstGeom",'--action','replace','--xml',E.tostring(cg,encoding='unicode'))
 d.notes(s,'Original vase composition. Both pictures contain the same native-exported GIF. True Intersect, then equivalent fillRect repair; no outside mask, video or encoded GIF substitute.')
d.run('close');d.run('validate');print(d.file)
