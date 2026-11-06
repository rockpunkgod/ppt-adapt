"""Original editable wheel and a matching UntilEnd control, built with OfficeCLI."""
from pathlib import Path
import argparse,sys
sys.path.insert(0,str(Path(__file__).resolve().parent))
from build_components import Deck
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',type=Path,required=True);a=p.parse_args();d=Deck(a.out,a.officecli)
geometry='<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="w" b="h"/><a:pathLst><a:path w="1000" h="1000"><a:moveTo><a:pt x="500" y="500"/></a:moveTo><a:lnTo><a:pt x="1000" y="500"/></a:lnTo><a:arcTo wR="500" hR="500" stAng="0" swAng="5400000"/><a:close/></a:path></a:pathLst></a:custGeom>'
for i in [1,2]:
 s=d.slide('16251D');d.shape(s,'kicker',1.5,1.1,30,1,'TRACE / 点击与相位',font='Noto Sans SC',size=15,color='DAEF78')
 d.shape(s,'title',1.5,3,30,1.8,'让选择，停在这一刻。' if i==1 else '同一次点击，不同的结果。',font='Noto Serif SC',size=33,color='F2F3EC')
 names=['rotor-bounds'];d.shape(s,'rotor-bounds',5.5,5.5,10,10)
 for j,c in enumerate(['F4CE6A','4D6E66','BCC7BB','DAEF78']):
  name='sector-'+str(j);names.append(name);d.shape(s,name,5.5,5.5,10,10,fill=c,geometry='rect')
  d.run('raw-set',f'/slide[{s}]','--xpath',f'//p:sp[p:nvSpPr/p:cNvPr[@name="{name}"]]/p:spPr/a:prstGeom','--action','replace','--xml',geometry)
  d.run('set',f'/slide[{s}]/shape[@name={name}]',props={'rotation':str(j*90)})
 for name,x,y,text in [('label-a',11.1,12,'A'),('label-b',7.5,12,'B'),('label-c',7.5,7.5,'C'),('label-d',11.1,7.5,'D')]:
  names.append(name);d.shape(s,name,x,y,2.4,1.4,text,font='Arial',size=30,color='16251D',align='center')
 names.append('phase-marker');d.shape(s,'phase-marker',13.5,10.1,.8,.8,fill='58D9CE',geometry='ellipse')
 d.run('add',f'/slide[{s}]','--type','group',props={'name':'rotor','shapes':','.join(f'/slide[{s}]/shape[@name={name}]' for name in names)})
 d.shape(s,'fixed-centre',10.25,10.25,.5,.5,fill='16251D',geometry='ellipse')
 d.shape(s,'result',19,8,12,2,'已暂停' if i==1 else '继续旋转',font='Noto Serif SC',size=38,color='F48665')
 d.shape(s,'detail',19,10.5,12,3,'再次单击：停在当前角度。' if i==1 else '再次单击：显示文字，转盘继续。',font='Noto Sans SC',size=17,color='BCC7BB')
 d.shape(s,'footer',1.5,17,31,1,'单击开始 → 再次单击显示状态。左方向键两次回退后可重新开始。',font='Noto Sans SC',size=14,color='BCC7BB')
 d.notes(s,'Original four-sector wheel. Native Spin1s; the next OnClick Fade is independent. Page1 UntilNextClick, page2 UntilEnd control. No precomputed stopping angle.')
d.run('close');print(a.out)
