"""Current showcase: a native city scene, editorial typography and camera states."""
from pathlib import Path
from scene import Scene

PAPER='F7F1E7'; INK='302E35'; ORANGE='E65A36'; MINT='B7D3C5'
WATER='87B8BC'; LILAC='A7A0C5'; WHITE='FFFDF6'; GRAY='6F6866'
SERIF='Noto Serif SC'; SANS='Noto Sans SC'; LATIN='Bahnschrift SemiCondensed'

class Showcase(Scene):
 def label(self,s,t,c=ORANGE):
  self.text(s,'label-'+str(s),1.7,1.0,30.4,1.1,t,23,c,LATIN)
 def title(self,s,t,x=1.7,y=3.0,w=29,h=4.0,size=43,c=INK):
  self.text(s,'title-'+str(s),x,y,w,h,t,size,c,SERIF,bold=True)
 def city(self,s,ox=24.2,oy=6.0,k=.85,variant=False):
  """One consistent projected city; all surfaces remain editable native shapes."""
  def xy(x,y,z=0):return (ox+k*(x-y),oy+k*(.5*(x+y)-z))
  def face(n,points,c):self.poly(s,'!!city-'+n,[xy(*v) for v in points],c)
  def tile(n,x,y,w,d,c,z=0):face(n,[(x,y,z),(x+w,y,z),(x+w,y+d,z),(x,y+d,z)],c)
  # The same ground and cut edges anchor every camera state.
  face('shadow',[(0,.5,-.5),(10,.5,-.5),(10,10.5,-.5),(0,10.5,-.5)],'DED4C5')
  face('base-left',[(0,10,0),(10,10,0),(10,10,-.35),(0,10,-.35)],'BEB4A7')
  face('base-right',[(10,0,0),(10,10,0),(10,10,-.35),(10,0,-.35)],'CBBEAA')
  tile('ground',0,0,10,10,'E2DDCA')
  tile('river',0,4.2,10,1.5,WATER)
  tile('quay-n',0,4.03,10,.17,'EDE5D6')
  tile('quay-s',0,5.7,10,.17,'EDE5D6')
  tile('street-v',4.0,0,.85,10,WHITE)
  tile('street-h',0,7.35,10,.65,WHITE)
  tile('side-street',0,2.8,10,.5,WHITE)
  # Two bridges retain their identity when the camera approaches the bookshop.
  tile('bridge',3.93,4.05,1.0,1.85,'D8B59B',.1)
  tile('bridge-rail-a',3.93,4.05,.08,1.85,'7C6D60',.25)
  tile('bridge-rail-b',4.85,4.05,.08,1.85,'7C6D60',.25)
  tile('park',.4,6.1,2.9,2.95,MINT)
  tile('plaza',5.3,6.05,2.0,1.0,'EACDAF')
  # Each building consists of roof and two visible wall planes, not a bitmap.
  buildings=[('book',5.4,1.1,2.3,1.5,1.55,ORANGE,'A9412D','C64A2F'),
   ('gallery',.65,.55,2.6,1.55,1.0,LILAC,'817A9B','928BAE'),
   ('cafe',7.6,6.05,1.7,1.0,.75,'D6A46D','A77B4E','C28F58'),
   ('north',8.35,.5,1.1,2.1,2.05,'D7CDBD','AAA294','C0B7A6'),
   ('studio',.8,3.5,2.15,.45,.65,'D5BAA0','A48C76','B49C84'),
   ('south',5.3,8.4,1.7,1.0,1.7,'C4BAB0','9E9389','B1A79B'),
   ('homes',7.7,8.4,1.7,1.0,1.2,'D9B89E','B18D74','C7A287')]
  for name,x,y,w,d,h,roof,left,right in buildings:
   if variant and name=='book':roof,left,right=LILAC,'817A9B','928BAE'
   face(name+'-left',[(x,y+d,0),(x+w,y+d,0),(x+w,y+d,h),(x,y+d,h)],left)
   face(name+'-right',[(x+w,y,0),(x+w,y+d,0),(x+w,y+d,h),(x+w,y,h)],right)
   tile(name+'-roof',x,y,w,d,roof,h)
   # Light clerestory and entrance details make the same landmark recognizable.
   for j in range(3):
    a=x+.2+j*(w-.35)/3
    face(name+'-window'+str(j),[(a,y+d,.35),(a+.22,y+d,.35),(a+.22,y+d,.72),(a,y+d,.72)],'F6E4BE')
   if name=='book':
    tile('book-roof-light',x+.25,y+.25,w-.5,.25,'F3BC8F',h+.01)
    face('book-sign',[(x+.25,y+d,h-.5),(x+w-.25,y+d,h-.5),(x+w-.25,y+d,h-.17),(x+.25,y+d,h-.17)],WHITE)
   elif name=='gallery':tile('gallery-skylight',x+.45,y+.4,w-.9,.55,'7C8F9D',h+.01)
  # Small trees are editable crowns, trunks and projected ground shadows.
  for j,(x,y) in enumerate([(1,6.7),(2,6.5),(1.2,8.35),(2.5,8.6),(3.3,1.2),(9,3.55),(9,7),(6.1,6.5)]):
   tile('tree-shadow'+str(j),x-.12,y+.05,.75,.65,'91B29F')
   face('trunk'+str(j),[(x-.035,y,0),(x+.035,y,0),(x+.035,y,.7),(x-.035,y,.7)],'82755E')
   px,py=xy(x,y,.85)
   self.sh(s,'!!city-tree'+str(j),px-.28*k,py-.4*k,.64*k,.8*k,fill=('6F9C80' if j%2 else '87AF91'),preset='ellipse')
  # A genuine route connects three places in this fictional walking proposal.
  route=[(1.8,7),(4.42,7),(4.42,3.55),(6.55,3.55),(6.55,2.95)]
  for j,(a,b) in enumerate(zip(route,route[1:])):
   x,y=a;u,v=b;half=.065
   if x==u:tile('route'+str(j),x-half,min(y,v),half*2,abs(v-y),ORANGE,.035)
   else:tile('route'+str(j),min(x,u),y-half,abs(u-x),half*2,ORANGE,.035)
  for j,(x,y) in enumerate([route[0],(4.42,5.2),route[-1]]):
   px,py=xy(x,y,.06)
   self.sh(s,'!!city-stop'+str(j),px-.17*k,py-.17*k,.34*k,.34*k,fill=ORANGE,preset='ellipse')
  # Motion is a projected 2D camera, not a claim of an embedded 3D model.
  self.note(s,'原创漫游概念提案中的虚构街区。建筑、河道、步行路线、树木由可编辑原生形状和路径制作，不是实测地图、上线产品或真实导航。共享对象用相同名称与投影坐标保留身份，二维视角不称为真实三维摄像机。')
 def cover(self):
  s=self.d.slide(PAPER);self.city(s,25.4,5.8,.83)
  self.label(s,'PPT ADAPT / 让 Agent 自主设计整套演示')
  self.title(s,'让想法，\n抵达观众。',y=4.4,w=16.8,h=6.6,size=54)
  self.text(s,'cover-body',1.8,13.3,15.3,3.0,'从一句需求出发，\n走进一个完整的世界。',26,INK,SANS)
  self.text(s,'cover-credit',18.0,17.0,14.4,1.25,'made by codex with ppt adapt',22,ORANGE,LATIN)
  return s
 def brief(self):
  s=self.d.slide(PAPER);self.city(s,29.0,5.5,.92)
  self.label(s,'起点 / 你想让观众理解什么')
  self.title(s,'先说出，\n你想带来的改变。',w=20,y=3.0,h=5.4,size=43)
  self.text(s,'quote-mark',1.4,8.1,3,4.3,'“',92,ORANGE,SERIF)
  self.text(s,'brief-request',3.5,9.7,16.5,5.4,'介绍一款城市步行工具。\n让第一次来到这里的人，\n从目的地，走进街区。',29,INK,SERIF)
  return s
 def story(self):
  s=self.d.slide(INK);self.label(s,'构思 / 从意图到观看路线','D5C8B7')
  self.title(s,'先有一条故事，\n再有每一页。',w=26,y=3,h=5.4,size=46,c=PAPER)
  for j,(x,y,w,h,t,c) in enumerate([(2,11.2,8.6,2,'想去哪里',PAPER),(12.9,9.9,9.5,2,'沿途看见什么',PAPER),(25,8.6,7.2,2,'如何开始',PAPER)]):
   self.text(s,'story-word'+str(j),x,y,w,h,t,31,c,SERIF,bold=True)
   self.text(s,'story-step'+str(j),x,y+2.6,6,1.2,'0'+str(j+1),23,'E89B7E',LATIN)
  self.poly(s,'story-route',[(2,16.9),(2,16.78),(31.5,12.3),(31.5,12.42)],ORANGE)
  return s
 def roam(self,state):
  s=self.d.slide(PAPER)
  # Whole district -> route -> bookshop. Geometry is shared, texts deliberately differ.
  poses=[(23.7,5.4,.86),(25.1,3.4,1.35),(15.2,4.7,2.4)]
  self.city(s,*poses[state])
  # An opaque reading layer stabilizes the text while the city moves behind it.
  self.sh(s,'!!reading-plane',0,0,15.4,19.05,fill=PAPER)
  self.label(s,['漫游 / 城市步行计划','漫游 / 把路线连成体验','漫游 / 抵达一处细节'][state])
  if state==0:
   self.text(s,'roam-wordmark',1.55,3.2,14,2.9,'漫游',68,ORANGE,SERIF,bold=True)
   self.title(s,'把周末，\n还给一座城。',y=7.0,w=14,h=5.4,size=39)
   self.text(s,'roam-caption',1.8,14.3,12.5,2.7,'走过河岸，遇见街角，\n把目的地变成一段经历。',24,INK,SANS)
  elif state==1:
   self.title(s,'一条路，\n也是一段故事。',y=3.4,w=13.7,h=5.7,size=38)
   for j,t in enumerate(['从河岸出发','穿过街区','在书店停一停']):
    self.text(s,'walk-num'+str(j),1.8,10.3+j*2.2,1.8,1.2,str(j+1).zfill(2),22,ORANGE,LATIN)
    self.text(s,'walk-label'+str(j),4.0,10.3+j*2.2,10.4,1.4,t,25,INK,SANS)
  else:
   self.title(s,'值得停下的，\n不只是一站。',y=3.4,w=13.8,h=5.7,size=38)
   self.text(s,'book-body',1.8,11,12.2,4.5,'河岸书店\n\n给偶遇留一点时间，\n让路线容得下好奇心。',25,INK,SANS)
  return s
 def range(self,work):
  s=self.d.slide('DCE6DE');self.label(s,'PPT ADAPT / 内容决定表达')
  self.title(s,'换一个主题，\n就换一种进入世界的方式。',y=2.8,h=5.2,size=40)
  self.pic(s,'made-roam',1.8,9.0,16.3,9.16875,work/'native-directions/slide-1.png')
  self.pic(s,'made-robot',18.8,7.7,13.3,7.48125,Path(__file__).resolve().parents[2]/'assets/demo/robomaster-roles.jpg')
  self.text(s,'range-label',19,16.1,13.1,1.5,'城市提案 × 工程赛事',25,INK,SANS)
  self.note(s,'左侧为本次原创漫游演示的实际原生导图；右侧来自本仓库已完成的RoboMaster介绍，GMaster2026车型原照身份不变。两份作品用途与内容不同，不是竞品对照或客户业绩。照片来源见随附demo/robomaster-sources.md。')
  return s
 def handoff(self):
  s=self.d.slide(PAPER);self.city(s,24.5,5.6,.9,variant=True)
  self.label(s,'交付 / 作品还能由你继续')
  self.title(s,'好表达，\n不在交付时结束。',w=18,y=3.4,h=6.0,size=44)
  self.text(s,'edit-lines',1.8,11,14.5,4.8,'改一句话。\n换一种配色。\n把下一段故事接着讲。',28,INK,SANS)
  self.text(s,'editable-word',19,16.7,13,1.2,'EDITABLE · PPTX',27,ORANGE,LATIN)
  return s
 def end(self):
  s=self.d.slide(ORANGE);self.label(s,'PPT ADAPT',PAPER)
  self.title(s,'下一场表达，\n从你的想法开始。',y=3.2,w=30,h=6.2,size=52,c=PAPER)
  self.text(s,'end-flow',1.8,11.8,30.2,2.3,'安装技能 → 描述需求 → 收到 PPTX 与预览',28,PAPER,SANS)
  self.text(s,'end-word',1.4,13.8,24,2.8,'MAKE IT YOURS.',62,INK,LATIN,bold=True)
  self.text(s,'end-credit',1.8,16.8,30,1.0,'made by codex with ppt adapt',22,PAPER,LATIN)
  self.text(s,'end-url',1.8,18.0,30,.9,'github.com/rockpunkgod/ppt-adapt',20,PAPER,LATIN)
  return s
 def alternate(self,k):
  s=self.d.slide('E8ECF1');self.label(s,'PPT ADAPT / 编辑台方向','304E76')
  if k==0:
   self.text(s,'alt-large',1.5,2.7,30,7.3,'IDEAS\nIN FOCUS',78,'304E76',LATIN,bold=True)
   self.text(s,'alt-title',1.8,13.4,29,3.0,'把复杂，讲得清楚。',42,INK,SANS,bold=True)
  else:
   self.title(s,'从散落的信息，\n聚焦一个判断。',w=30)
   for j,(x,y,w,h) in enumerate([(2,9,7,6),(10,10,11,4),(22,7,9,9)]):
    self.sh(s,'alt-panel'+str(j),x,y,w,h,fill=['C9D4E0','A0B5CE','304E76'][j])
    self.text(s,'alt-label'+str(j),x+.4,y+.4,w-.8,2,['材料','关系','判断'][j],30,PAPER if j==2 else INK,SANS)
  return s
