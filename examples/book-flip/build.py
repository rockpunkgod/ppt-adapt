"""Editable curved paper and spine-centered native Morph example."""
from pathlib import Path
import sys,argparse,xml.etree.ElementTree as E
ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT/'officecli-ppt-design/scripts'))
from build_components import Deck
parser=argparse.ArgumentParser();parser.add_argument('--officecli',required=True);parser.add_argument('--out',required=True);args=parser.parse_args()
p=Path(args.out);d=Deck(p,args.officecli);ops=[]
def sh(s,n,x,y,w,h,t='',size=20,**kw):
 props=dict(name=n,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',fill='none',line='none',font='Noto Sans SC',size=str(size),color='26372C',margin='0cm',valign='center')
 if t:props['text']=t
 props.update({k:str(v).lower() if isinstance(v,bool) else str(v) for k,v in kw.items()});ops.append(dict(command='add',parent=f'/slide[{s}]',type='shape',props=props))
def flush():
 if ops:d.batch(ops.copy());ops.clear()
def page(s,name,x,y,w,h,fill):
 sh(s,name,x,y,w,h,fill=fill);flush()
 ns='http://schemas.openxmlformats.org/drawingml/2006/main';q=lambda n:'{'+ns+'}'+n
 cg=E.Element(q('custGeom'));[E.SubElement(cg,q(n)) for n in ['avLst','gdLst','ahLst','cxnLst']];E.SubElement(cg,q('rect'),l='0',t='0',r='r',b='b');lst=E.SubElement(cg,q('pathLst'));path=E.SubElement(lst,q('path'),w='1000',h='1400',stroke='0')
 E.SubElement(E.SubElement(path,q('moveTo')),q('pt'),x='0',y='60')
 for points in [[(330,0),(680,0),(1000,38)]]:
  c=E.SubElement(path,q('cubicBezTo'));[E.SubElement(c,q('pt'),x=str(a),y=str(b)) for a,b in points]
 E.SubElement(E.SubElement(path,q('lnTo')),q('pt'),x='1000',y='1362')
 c=E.SubElement(path,q('cubicBezTo'));[E.SubElement(c,q('pt'),x=str(a),y=str(b)) for a,b in [(680,1330),(330,1330),(0,1400)]]
 E.SubElement(path,q('close'))
 d.run('raw-set',f'/slide[{s}]','--xpath',f"//p:sp[p:nvSpPr/p:cNvPr[@name='{name}']]/p:spPr/a:prstGeom",'--action','replace','--xml',E.tostring(cg,encoding='unicode'))
s=d.slide('16251D');sh(s,'label',1.5,.8,30,1,'TRACE',17,color='DAEF78')
sh(s,'book-cover',5.4,3.8,23.1,11.6,fill='697564',geometry='roundRect',shadow='000000-8-45-2-35')
sh(s,'book-stack',5.65,3.95,22.6,11.2,fill='CFCFBC')
sh(s,'left-paper',5.85,4.1,11.1,10.8,fill='F2F0E5');sh(s,'right-paper',16.95,4.1,11.1,10.8,fill='F2F0E5')
sh(s,'spine',16.85,4.1,.17,10.8,fill='B8BDA9')
sh(s,'left-title',7,5.1,8.6,2.2,'从资料，\n到证据。',33,font='Noto Serif SC')
sh(s,'left-body',7,9,8.4,3.8,'把散落的工作记录，\n整理成下一次判断\n能够复用的依据。',20,color='637066')
sh(s,'left-page',7,13.9,8,.8,'01 / 收集',13,color='637066')
sh(s,'right-title',18.1,5,8.8,1.5,'让答案接回工作。',26,font='Noto Serif SC')
for i,t in enumerate(['检索资料','核对来源','交接行动']):
 sh(s,f'node-{i}',18.1,7.2+i*2.05,8.5,1.4,t,fill='DAEF78',margin='.35cm',size=21)
sh(s,'right-page',18.1,13.9,8,.8,'02 / 接续',13,color='637066')
sh(s,'hint',1.5,17.5,30,1,'单击翻页；左方向键返回后，可再次翻页。',13,color='BCC7BB');flush()
page(s,'turning-right',16.95,4.1,11.1,10.8,'F7F4E8')
page(s,'axis-left-helper',5.85,4.1,11.1,10.8,'F7F4E8');flush()
d.run('add','/slide[1]','--type','group',props={'name':'!!turning-page','shapes':'/slide[1]/shape[@name=turning-right],/slide[1]/shape[@name=axis-left-helper]'})
s=d.clone(1);d.run('set','/slide[2]',props={'transition':'morph-2000','advanceClick':True})
d.run('close');d.run('validate');print(p)
