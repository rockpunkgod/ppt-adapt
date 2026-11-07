"""Original current showcase: editable paper architecture and a knowledge scene."""
from pathlib import Path
import sys, math, zipfile, xml.etree.ElementTree as ET
from xml.dom import minidom
SKILL=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(SKILL/'scripts'))
from build_components import Deck
from case_pages import Author
P='F5F2E9'; INK='182139'; BLUE='2549E8'; LIGHT='D5DFFF'; LIME='DFFF85'; MUTED='657089'
SERIF='Noto Serif SC'; SANS='Noto Sans SC'; LATIN='Bahnschrift SemiCondensed'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
N={'a':A,'p':'http://schemas.openxmlformats.org/presentationml/2006/main'}

class Scene(Author):
 def __init__(self,d):
  super().__init__(d,SKILL/'assets/demo/aura-device.png');self.polys=[]
 def text(self,s,n,x,y,w,h,t,size=25,c=INK,font=SANS,**kw):
  self.sh(s,n,x,y,w,h,t,size,c,font=font,**kw)
 def poly(self,s,n,pts,fill):
  xs,ys=zip(*pts);x,y=min(xs),min(ys);w,h=max(xs)-x,max(ys)-y
  if 'knowledge-link' in n:
   self.sh(s,n,x,y,w,h,fill='none',line=fill+':0.8')
   pts=pts[:2]
  else:self.sh(s,n,x,y,w,h,fill=fill)
  self.polys.append((s,n,[(round((a-x)/w*100000),round((b-y)/h*100000)) for a,b in pts]))
 def ribbon(self,s,state):
  # Same physical four facets retain identity while opening into a story strip.
  shapes=[[(0,2),(8,0),(12,3.2),(4,5.2)],[(4,5.2),(12,3.2),(12,8.2),(4,10.2)],[(4,10.2),(12,8.2),(16,11.2),(8,13.2)],[(8,13.2),(16,11.2),(16,12),(8,14)]]
  if state==0:trans=lambda x,y:(17.0+x*1.05,2.2+y*1.05)
  elif state==1:trans=lambda x,y:(-5+x*1.1,3.4+y*1.1)
  elif state==2:
   shapes=[[(0,0),(9,0),(9,4.2),(0,4.2)],[(9,0),(17.5,0),(17.5,4.2),(9,4.2)],[(17.5,0),(28.5,0),(28.5,4.2),(17.5,4.2)],[(0,4.2),(28.5,4.2),(28.5,4.5),(0,4.5)]]
   trans=lambda x,y:(2.1+x,8.1+y)
  elif state==3:trans=lambda x,y:(21+x*.95,5+y*.95)
  elif state==4:trans=lambda x,y:(17+x*1.05,3.3+y*1.05)
  else:trans=lambda x,y:(17+x*.85,3.1+y*.85)
  for j,pts in enumerate(shapes):
   color=['FFFFFF',BLUE,'CFD9FF',INK][j]
   self.poly(s,'!!paper-facet-'+str(j),[trans(x,y) for x,y in pts],color)
  if state in (0,3,4,5):
   self.poly(s,'!!paper-edge',[trans(x,y) for x,y in [(0,2),(4,5.2),(4,5.4),(0,2.2)]],LIGHT)
 def brand(self,s,dark=False):
  self.text(s,'brand',1.7,1.0,22,1.1,'PPT ADAPT',24,LIME if dark else BLUE,LATIN,bold=True)
 def head(self,s,t,c=INK):self.text(s,'headline',1.7,3.0,30.5,4.0,t,43,c,SERIF,bold=True)
 def cover(self):
  s=self.d.slide(P);self.ribbon(s,0);self.brand(s)
  self.text(s,'cover-title',1.7,5.0,17,6.5,'让想法，\n有自己的形状。',48,INK,SERIF,bold=True)
  self.text(s,'cover-body',1.8,14.2,15.7,2.6,'从一段内容出发，\n由 Agent 设计一整套演示。',25)
  self.text(s,'paper-word',23.2,4.35,7,1.4,'CONTENT',26,BLUE,LATIN)
  return s
 def brief(self):
  s=self.d.slide(P);self.ribbon(s,1);self.brand(s)
  self.text(s,'brief-head',13.5,3.0,18.5,3.8,'你只需要，\n把想说的说出来。',41,INK,SERIF,bold=True)
  self.sh(s,'!!brief-paper',13.5,8.0,18.3,8.4,fill='FFFFFF')
  self.text(s,'brief-question',14.5,9.0,16.2,6.4,'介绍一款团队知识工具。\n\n让产品团队理解：\n资料怎样变成判断，\n判断怎样接上行动。',27)
  return s
 def story(self):
  s=self.d.slide(P);self.ribbon(s,2);self.brand(s);self.head(s,'内容，先成为一条叙事。')
  for j,(x,w,num,h,b) in enumerate([(3.0,7.5,'01','为什么需要','散落的资料，\n难以支持判断。'),(12.0,7.2,'02','怎样改变','把答案与依据，\n重新连接起来。'),(20.6,10.0,'03','下一步是什么','让一次判断，\n接上下一步行动。')]):
   self.text(s,f'story-n{j}',x,7,5,1,num,22,BLUE,LATIN)
   self.text(s,f'story-h{j}',x,8.9,w,1.7,h,30,'FFFFFF' if j==1 else INK,SERIF,bold=True)
   self.text(s,f'story-b{j}',x,14.0,w,2.7,b,24)
  return s
 def knowledge(self,state):
  s=self.d.slide(INK)
  # Document constellation transforms into an evidence view, then a working page.
  base=[(20.0,5.7,8.3,6.0),(15.6,2.2,5.7,3.1),(27.0,2.0,5.1,3.0),(13.5,11.5,6.1,3.8),(27.0,12.3,6,3.8)]
  if state==0: poses=base
  elif state==1:poses=[(1.8,7.1,17.4,9.8),(22.2,7.0,9.2,2.4),(24.1,10.2,7.3,2.4),(22.2,13.4,9.2,2.4),(35,5,6,4)]
  else:poses=[(12.7,5.6,19.4,11.2),(14,7.5,4.7,2.2),(14,10.3,4.7,2.2),(14,13.1,4.7,2.2),(35,7,6,4)]
  if state==0:
   for j in range(1,5):
    x,y,w,h=base[j];self.poly(s,'!!knowledge-link'+str(j),[(24,8.7),(x+w/2,y+h/2),(x+w/2+.07,y+h/2+.1),(24.07,8.8)],'6B7BA8')
  else:
   for j in range(1,5):
    x,y,w,h=poses[j]
    start=(18,12) if j<4 else (35,10)
    if state==2 and j<4:start=(19.4,8.6+j*2)
    self.poly(s,'!!knowledge-link'+str(j),[start,(x,y+h/2),(x+.06,y+h/2+.07),(start[0]+.04,start[1]+.08)],'6B7BA8')
  for j,(x,y,w,h) in enumerate(poses):
   self.sh(s,'!!knowledge-doc'+str(j),x,y,w,h,fill=BLUE if j==0 else 'F5F2E9')
   if j>0:
    self.sh(s,'!!knowledge-tab'+str(j),x,y,.13,h,fill=LIME if j==3 else 'A5B8FF')
   if 0<j<4:
    self.text(s,'!!doc-label'+str(j),x+.5,y+(.25 if state==1 else .6),w-1,1.0,['','产品资料','风险清单','项目决策'][j],21,INK if state==0 else BLUE,bold=(state>0))
  self.text(s,'echo-brand',1.7,1.0,29,1.1,'回声 / 团队知识工作台',23,LIME)
  if state==0:
   self.text(s,'echo-head',1.7,4.2,13.5,6.0,'把散落的知识，\n连成下一次判断。',41,'FFFFFF',SERIF,bold=True)
   self.text(s,'echo-body',1.8,13.3,11.3,3.4,'从资料到依据，\n从依据到行动。',26,'C0CAE2')
   self.text(s,'answer-word0',20.8,7,7,2,'回声',42,'FFFFFF',SERIF,bold=True)
   for j in range(3):self.sh(s,'doc-line'+str(j),21,10+j*.38,4.8-j*.8,.06,fill='94A9FF')
   self.text(s,'doc-d',27.7,13.2,4.8,1.6,'会议记录',22)
  elif state==1:
   self.text(s,'evidence-head',1.7,2.8,30.5,3.4,'找到答案，也看见它从哪里来。',39,'FFFFFF',SERIF,bold=True)
   self.text(s,'answer-word1',2.7,8.0,15.5,1.9,'发布前，先确认什么？',31,'FFFFFF',SERIF,bold=True)
   for j,t in enumerate(['确认本次交付的范围','复核尚未解决的风险','让负责人确认下一步']):
    self.text(s,'answer-'+str(j),2.8,11.0+j*1.6,15,1.3,t,25,'FFFFFF')
   for j,(h,b) in enumerate([('产品资料','本期交付范围'),('风险清单','尚待复核事项'),('项目决策','负责人及行动')]):
    x,y,w,_=poses[j+1];self.text(s,f'source-b{j}',x+.5,y+1.2,w-1,.9,b,19)
  else:
   self.text(s,'action-head',1.7,4.1,10.3,6.3,'让每次判断，\n接得上下一步。',37,'FFFFFF',SERIF,bold=True)
   self.text(s,'action-body',1.8,13.3,9.7,2.8,'在同一处看清：\n结论、依据与行动。',24,'C0CAE2')
   self.text(s,'answer-word2',13.8,6.1,17.1,1.1,'项目 / 发布决策',24,'FFFFFF',bold=True)
   self.text(s,'ui-decision',20,8.0,11,2.2,'先做小范围试点',29,'FFFFFF',SERIF,bold=True)
   self.text(s,'ui-basis',20,10.6,10.6,2.2,'范围已确认\n风险项待复核',23,'FFFFFF')
   self.text(s,'ui-next',20,13.4,10.6,2.3,'补充试点反馈\n再决定推广范围',23,'FFFFFF')
  self.note(s,'原创回声概念产品发布页。为PPT Adapt能力演示设计的内容与界面，不是已上线软件、真实客户案例或性能实测。页面全部文字、文档、路径和界面由原生对象构成。')
  return s
 def panorama(self,work):
  s=self.d.slide(P);self.ribbon(s,3);self.brand(s);self.head(s,'独特的表达，贯穿一整个故事。')
  for j in range(3):self.pic(s,'story-page'+str(j),1.7+j*10.25,8.0,10.0,5.625,work/f'native-directions/slide-{j+5}.png')
  self.text(s,'panorama-line',1.8,15.0,28.6,2.2,'主张 → 依据 → 行动\n每页推进内容，也接住上一页的注意力。',25)
  return s
 def handoff(self):
  s=self.d.slide(P);self.ribbon(s,4);self.brand(s)
  self.text(s,'handoff-head',1.7,4.0,17,5.7,'好作品，\n也能继续生长。',47,INK,SERIF,bold=True)
  self.text(s,'handoff-body',1.8,12,14.2,4.2,'交付可编辑 PPTX。\n文字、图形与页面结构，\n留给你继续表达。',26)
  self.text(s,'handoff-word',23,4.8,7,2,'PPTX',43,BLUE,LATIN,bold=True)
  return s
 def end(self):
  s=self.d.slide(P);self.ribbon(s,5);self.brand(s)
  self.text(s,'end-head',1.7,4.3,18,6.2,'下一个想法，\n从你开始。',49,INK,SERIF,bold=True)
  self.text(s,'end-body',1.8,12.4,16,3.4,'安装 PPT Adapt，描述你的需求。\n收到一套演示，以及它的预览。',25)
  self.text(s,'end-url',1.8,17,30.2,1,'github.com/rockpunkgod/ppt-adapt',22,BLUE,LATIN)
  self.text(s,'end-paper-word',21.1,4.7,6.8,1.3,'YOUR IDEA',24,BLUE,LATIN)
  return s
 def alternate(self,k):
  s=self.d.slide('EDE5D4');red='A83D36'
  self.text(s,'alt-brand',1.7,1.1,25,1.2,'PPT ADAPT',25,red,LATIN)
  for j in range(4):self.sh(s,'alt-ring'+str(j),18+j*.8,2.5+j*1.5,17-j*1.8,17-j*1.8,fill='none',line=red+':2',preset='ellipse')
  self.text(s,'alt-title',1.7,4.1,20,6,'让想法，\n被看见，被记住。' if k==0 else '从片段，\n听见一个完整的故事。',46,INK,SERIF,bold=True)
  self.text(s,'alt-body',1.8,13.9,16,3,'内容形成节奏，\n表达留下回声。' if k==0 else '主张、证据与行动，\n围绕同一个问题展开。',26)
  return s
 def finish(self):
  self.flush();self.d.run('close')
  with zipfile.ZipFile(self.d.file) as z:parts={n:z.read(n) for n in z.namelist()}
  trees={}
  for s,name,pts in self.polys:
   key=f'ppt/slides/slide{s}.xml'
   if key not in trees:trees[key]=minidom.parseString(parts[key])
   root=trees[key]
   shape=next(v for v in root.getElementsByTagNameNS(N['p'],'sp') if v.getElementsByTagNameNS(N['p'],'cNvPr')[0].getAttribute('name').lstrip('!')==name.lstrip('!'))
   sp=shape.getElementsByTagNameNS(N['p'],'spPr')[0];geom=sp.getElementsByTagNameNS(A,'prstGeom')[0]
   def el(n,**attrs):return ET.Element('{'+A+'}'+n,{k:str(v) for k,v in attrs.items()})
   g=el('custGeom')
   for tag in ['avLst','gdLst','ahLst','cxnLst']:g.append(el(tag))
   g.append(el('rect',l='l',t='t',r='r',b='b'));lst=el('pathLst');path=el('path',w=100000,h=100000,**({'fill':'none'} if 'knowledge-link' in name else {}))
   for j,(x,y) in enumerate(pts):p=el('moveTo' if j==0 else 'lnTo');p.append(el('pt',x=x,y=y));path.append(p)
   if 'knowledge-link' not in name:path.append(el('close'))
   lst.append(path);g.append(lst)
   new=minidom.parseString(ET.tostring(g,encoding='utf-8')).documentElement
   sp.replaceChild(root.importNode(new,True),geom)
  # Preserve original namespace declarations, including prefixes used in mc values.
  for key,tree in trees.items():parts[key]=tree.toxml(encoding='utf-8')
  with zipfile.ZipFile(self.d.file,'w',zipfile.ZIP_DEFLATED) as z:
   for key,data in parts.items():z.writestr(key,data)
  self.d.run('validate');self.d.run('close')
