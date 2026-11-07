from pathlib import Path
P='F4F5F7';I='111318';A='FF5B2A';M='A3ADBA';B='Noto Sans SC';H='Noto Sans SC Black';L='Bahnschrift SemiCondensed'
class Author:
 def __init__(self,d,asset):self.d=d;self.asset=Path(asset).resolve();self.q=[]
 def sh(self,s,n,x,y,w,h,t='',size=24,c=I,fill='none',font=B,**kw):
  p=dict(name=n,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',fill=fill,line='none')
  if t:p.update(text=t,font=font,size=size,color=c,margin='0cm',valign='center',autoFit='none',lineSpacing='1.06x')
  p.update(kw);self.q.append(dict(command='add',parent=f'/slide[{s}]',type='shape',props=p))
 def flush(self):
  if self.q:self.d.batch(self.q.copy());self.q.clear()
 def pic(self,s,n,x,y,w,h,src=None):
  self.flush();self.d.run('add',f'/slide[{s}]','--type','picture',props=dict(name=n,src=str(Path(src or self.asset).resolve()),x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm'))
 def note(self,s,t):self.flush();self.d.notes(s,t)
 def case(self,k):
  s=self.d.slide(I)
  if k==0:
   self.pic(s,'!!device',13.5,4.5,20.5,20.5*2/3)
   self.sh(s,'case-kicker',1.6,1,26,1,'作品示范 / AURA 边缘AI模块',21,A)
   self.sh(s,'case-wordmark',1.3,2.45,15.1,4.1,'AURA',103,P,font=L,bold=True)
   self.sh(s,'case-hero',1.6,8.05,14.7,5.1,'让智能，\n在设备这一端发生。',43,P,font=H,bold=True)
   self.sh(s,'case-promise',1.6,14.7,12.3,2.5,'从采集到推理，\n在设备侧完成判断。',25,'C2C9D2')
  elif k==1:
   self.pic(s,'!!device',.2,6.4,17,17*2/3)
   self.sh(s,'logic-kicker',1.6,1,27,1,'工作方式 / 本地处理，连接现场',21,A)
   self.sh(s,'logic-title',1.6,2.8,30.7,2.6,'离线工作，也能接上行动。',41,P,font=H,bold=True)
   self.sh(s,'logic-line',18.1,7.0,.045,8.1,fill='49505D')
   for j,(h,b) in enumerate([('采集','接入相机与设备信号。'),('推理','在本地完成信息处理。'),('响应','把判断交回现场设备。')]):
    y=6.25+j*3.8;self.sh(s,f'logic-node{j}',17.8,y+.42,.65,.65,fill=A,preset='ellipse')
    self.sh(s,f'logic-head{j}',19.3,y,12.3,1.5,h,32,P,font=H,bold=True)
    self.sh(s,f'logic-body{j}',19.3,y+1.8,12.3,1.4,b,23,'C2C9D2')
  else:
   self.pic(s,'!!device',-8,.3,29,29*2/3)
  self.sh(s,'!!use-panel',18 if k==2 else 33.87,0,15.87,19.05,fill=P)
  if k==2:
   self.sh(s,'use-kicker',19.2,1.15,12.5,.9,'应用场景 / AURA',21,A)
   self.sh(s,'use-title',19.2,3.0,13.2,3.9,'把判断，\n交给现场。',42,I,font=H,bold=True)
   for j,(h,b) in enumerate([('产线识别','在设备侧发现异常。'),('现场巡检','让采集连接到下一步。'),('边缘感知','就近处理，即时响应。')]):
    y=8.0+j*3.2;self.sh(s,f'use-head{j}',19.2,y,12.3,1.3,h,28,I,font=H,bold=True);self.sh(s,f'use-body{j}',19.2,y+1.5,12.3,1.15,b,21,'596273')
   self.sh(s,'aura-stamp',1.6,1.2,14,1,'AURA',24,A,font=L,bold=True)
  self.note(s,'本页为PPT Adapt本次制作的原创概念产品示范。AURA不是在售产品，未作任何实测性能承诺。主视觉是原创生成PNG；页面文字、结构和动画为PowerPoint原生对象。共享图片使用真实相同素材和比例，通过Morph表达连续变化。')
  if k:self.d.run('set',f'/slide[{s}]',props={'transition':'morph-1600','advanceClick':True})
  return s
 def alternate(self,kind):
  s=self.d.slide('FAFAF6')
  self.pic(s,'alt-device',.6,6,19.5,13)
  self.sh(s,'alt-kicker',1.6,1.1,30,1,'AURA / FIELD INTELLIGENCE',21,'213F73',font='Georgia')
  if kind==0:
   self.sh(s,'alt-title',1.6,2.55,30,3.2,'让智能走进现场。',49,'172747',font='Noto Serif SC',bold=True)
   self.sh(s,'alt-caption',21.1,8.0,11,5.5,'本地推理。\n连接设备。\n接上行动。',31,'172747',font='Noto Serif SC')
   self.sh(s,'alt-rule',21.1,14.75,10.6,.07,fill='213F73')
  else:
   self.sh(s,'alt-title',1.6,2.55,30,3.0,'从采集，到现场响应。',44,'172747',font='Noto Serif SC',bold=True)
   for j,(h,b) in enumerate([('输入','相机与设备信号'),('判断','本地推理与处理'),('行动','把结果交回设备')]):
    y=6.6+j*3.8;self.sh(s,f'alt-num{j}',21,y,2,1.5,f'0{j+1}',32,'213F73',font='Georgia');self.sh(s,f'alt-head{j}',24,y,8.5,1.5,h,29,'172747',font='Noto Serif SC');self.sh(s,f'alt-body{j}',21,y+1.8,11,1.2,b,22,'596273')
  self.note(s,'另一种艺术方向的代表页：浅色出版式布局与衬线字体，和深色产品发布方向使用同一概念产品与同一图片。')
  return s
