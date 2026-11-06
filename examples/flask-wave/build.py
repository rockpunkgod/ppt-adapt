from pathlib import Path
import sys,argparse,xml.etree.ElementTree as E
ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT/'officecli-ppt-design/scripts'))
from build_components import Deck
parser=argparse.ArgumentParser();parser.add_argument('--officecli',required=True);parser.add_argument('--out',required=True);args=parser.parse_args()
d=Deck(args.out,args.officecli);ops=[]
def sh(s,n,x,y,w,h,t='',size=20,**kw):
 props=dict(name=n,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',fill='none',line='none',font='Noto Sans SC',size=str(size),color='F2F3EC',margin='0cm',valign='center')
 if t:props['text']=t
 props.update({k:str(v).lower() if isinstance(v,bool) else str(v) for k,v in kw.items()});ops.append(dict(command='add',parent=f'/slide[{s}]',type='shape',props=props))
def flush():
 if ops:d.batch(ops.copy());ops.clear()
def curve(s,n,w,h,commands):
 flush();ns='http://schemas.openxmlformats.org/drawingml/2006/main';q=lambda n:'{'+ns+'}'+n
 cg=E.Element(q('custGeom'));[E.SubElement(cg,q(n)) for n in ['avLst','gdLst','ahLst','cxnLst']];E.SubElement(cg,q('rect'),l='0',t='0',r='r',b='b');lst=E.SubElement(cg,q('pathLst'));path=E.SubElement(lst,q('path'),w=str(w),h=str(h),stroke='1')
 for kind,points in commands:
  cmd=E.SubElement(path,q(kind))
  for x,y in points:E.SubElement(cmd,q('pt'),x=str(x),y=str(y))
 d.run('raw-set',f'/slide[{s}]','--xpath',f"//p:sp[p:nvSpPr/p:cNvPr[@name='{n}']]/p:spPr/a:prstGeom",'--action','replace','--xml',E.tostring(cg,encoding='unicode'))
flask=[('moveTo',[(360,0)]),('lnTo',[(640,0)]),('lnTo',[(640,440)]),('lnTo',[(970,1200)]),('cubicBezTo',[(1000,1300),(970,1400),(880,1400)]),('lnTo',[(120,1400)]),('cubicBezTo',[(30,1400),(0,1300),(30,1200)]),('lnTo',[(360,440)]),('close',[])]
for mode in [1,2]:
 s=d.slide('16251D')
 sh(s,'kicker',1.5,1.1,30,1,'TRACE / 视觉实验',15,color='DAEF78')
 sh(s,'title',1.5,4,17,4.5,'让变化发生，\n让边界保持清晰。',size=37,font='Noto Serif SC')
 sh(s,'body',1.5,10.2,16,3.5,'一条清楚的轮廓，\n承接内部的流动。\n外部信息保持稳定。',size=23,color='BCC7BB')
 sh(s,'caption',1.5,17.2,30,1,'单击：'+('液面上升与末段回弹。' if mode==1 else '两层波浪沿不同斜线升起。'),size=14,color='BCC7BB')
 if mode==1:sh(s,'water',19.5,8.3,12,12,fill='35A9C7')
 else:
  for name,x,y,col in [('wave-back',15,16.5,'196E8C'),('wave-front',17,16.8,'4BBAD0')]:
   sh(s,name,x,y,32,13,fill=col,opacity=.85 if name=='wave-front' else 1)
   commands=[('moveTo',[(0,50)])]
   for k in range(8):
    x0=k*400
    commands += [('cubicBezTo',[(x0+66,0),(x0+133,0),(x0+200,50)]),('cubicBezTo',[(x0+266,100),(x0+333,100),(x0+400,50)])]
   commands += [('lnTo',[(3200,1300)]),('lnTo',[(0,1300)]),('close',[])]
   curve(s,name,3200,1300,commands)
 sh(s,'flask-cut',21.5,4.2,8,11.8,fill='F2F3EC');curve(s,'flask-cut',1000,1400,flask)
 sh(s,'lip',24.18,4.15,2.64,.07,fill='F2F3EC');flush()
 d.notes(s,'原创瓶形液面样稿。页面是视觉动效演示，不表达真实体积或测量数据。')
d.run('close');d.run('validate');print(d.file)
