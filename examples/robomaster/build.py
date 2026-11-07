"""Original RoboMaster introduction, built with existing PPT Adapt primitives."""
from pathlib import Path
import argparse,json,sys
ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT/'officecli-ppt-design/scripts/demo_source'))
from scene import Scene,Deck
ASSET=Path(__file__).resolve().parent/'assets'
IMAGES=json.loads((ASSET/'images.json').read_text(encoding='utf-8'))
WHITE='FFFFFF';INK='15191F';RED='EF4438';CYAN='35A9C3';GRAY='727B85';LIGHT='E5E8EB'
HEAVY='Noto Sans SC Black';SANS='Noto Sans SC';LATIN='Bahnschrift SemiCondensed'
SOURCES='官方赛事介绍：https://www.robomaster.com/zh-CN/robo/overview；RMUC介绍：https://www.robomaster.com/zh-CN/robo/rm；2027规则：https://bbs.robomaster.com/wiki/20204847/1969149 。核验日期2026-11-07。'

class Robo(Scene):
 def txt(self,s,n,x,y,w,h,t,size=25,c=INK,font=SANS,**kw):self.text(s,n,x,y,w,h,t,size,c,font,**kw)
 def photo(self,s,n,key,x,y,w):
  info=IMAGES[key];self.pic(s,n,x,y,w,w*info['height']/info['width'],ASSET/info['file'])
 def tag(self,s,t,dark=False):self.txt(s,'chapter',1.7,1.0,30.2,1.15,t,22,WHITE if dark else RED,LATIN)
 def robot_photo(self,s,x,y,w):
  self.flush();info=IMAGES['arena-banner']
  h=w/(info['width']*.24/info['height'])
  self.d.run('add',f'/slide[{s}]','--type','picture',props={'name':'!!unit-image','src':str(ASSET/info['file']),'x':f'{x}cm','y':f'{y}cm','width':f'{w}cm','height':f'{h}cm','cropLeft':70,'cropRight':6})
 def heading(self,s,t,w=30,c=INK):self.txt(s,'heading',1.7,2.75,w,4.5,t,43,c,HEAVY)
 def line(self,s,n,x1,y1,x2,y2,c=GRAY):
  self.poly(s,'!!knowledge-link-'+n,[(x1,y1),(x2,y2),(x2+.01,y2+.02),(x1+.01,y1+.02)],c)
 def cover(self,alt=False):
  if alt:
   s=self.d.slide(INK);self.photo(s,'cover-photo','arena-banner',-5.0,7.0,52)
   self.tag(s,'ROBOMASTER / 机甲大师',True)
   self.txt(s,'alt-title',1.7,3.0,29,4.5,'让工程，成为现场。',56,WHITE,HEAVY)
   self.txt(s,'alt-name',1.7,14.4,31,3.2,'ROBOMASTER',67,WHITE,LATIN,bold=True)
   return s
  s=self.d.slide(WHITE);self.tag(s,'机甲大师 / 认识 RoboMaster')
  self.txt(s,'cover-robo',1.55,3.1,18,4.1,'ROBO',97,INK,LATIN,bold=True)
  self.txt(s,'cover-master',1.55,7.0,19,4.1,'MASTER',97,INK,LATIN,bold=True)
  self.txt(s,'cover-head',20.2,3.8,12.0,5.4,'把工程，\n带上赛场。',43,INK,HEAVY)
  self.txt(s,'cover-season',20.3,9.7,12.2,1.2,'2027 赛季入门',24,RED)
  self.sh(s,'photo-bed',0,12.2,33.87,6.85,fill=INK)
  self.photo(s,'!!competition-photo','arena-banner',0,12.2,33.87)
  self.txt(s,'cover-statement',1.7,17.8,30,1,'技术成为现场，团队成为答案。',22,WHITE)
  self.note(s,SOURCES+'封面照片来自官网RMUC页面的公开横幅，不把照片中的机器人指认为2027特定兵种。')
  return s
 def what(self):
  s=self.d.slide(INK);self.tag(s,'DJI 大疆发起 / 高校机器人赛事',True)
  self.heading(s,'学生造机甲，团队打配合。',c=WHITE)
  self.photo(s,'!!competition-photo','arena-banner',-7.0,7.1,47.9)
  for j,(h,b) in enumerate([('自主研发','高校战队设计并制作机器人'),('战术对抗','在同一赛场协作、交锋'),('工程实践','把多个学科接成可运行系统')]):
   x=1.7+j*10.6;self.txt(s,'what-h'+str(j),x,14.4,9.4,1.5,h,31,RED if j==0 else WHITE,HEAVY)
   self.txt(s,'what-b'+str(j),x,16.5,9.5,1.8,b,22,'C6CDD4')
  self.note(s,SOURCES+'RoboMaster由DJI大疆发起，介绍以高校系列赛中的超级对抗赛为主。其他入口包括高校联盟赛及高校人工智能挑战赛。')
 def arena(self,alt=False):
  s=self.d.slide(WHITE);self.tag(s,'RMUC 2027 / 观赛的第一张地图')
  self.heading(s,'先看目标，\n再看交火。',w=12)
  self.photo(s,'!!field','field-axon',10.0,5.0,23.5)
  for j,t in enumerate(['攻击对方基地','争夺关键阵地','建立全场控制']):self.txt(s,'goal'+str(j),1.8,10.0+j*2.1,11.4,1.6,t,27,RED if j==0 else INK,HEAVY)
  self.note(s,SOURCES+'官方2027规则V1.0.0，印刷页29–30，图4-3。依官网概述，击毁基地或在结束时取得更高全场控制程度可获胜；此页为观赛概述，不替代完整获胜判定。')
 def roles(self):
  s=self.d.slide(WHITE);self.tag(s,'地面机器人 / 同一目标，不同分工')
  self.txt(s,'role-head',15.2,2.8,17,4.5,'不同分工，\n合成一支战队。',39,INK,HEAVY)
  self.photo(s,'!!field','field-axon',-34.0,6.4,31.0)
  self.robot_photo(s,1.7,5.1,11.5)
  self.txt(s,'role-photo-caption',1.8,14.3,11.5,2.6,'一台机器人，\n接入整个团队。',25,INK,HEAVY)
  for j,(h,b) in enumerate([('重装','获取、装配能量单元，提供团队增益'),('步兵','地面机动，参与弹丸对抗'),('哨兵','全自动或半自动运行')]):
   y=8.0+j*2.7;self.txt(s,'role-h'+str(j),17.0,y,14.6,1.35,h,29,RED,HEAVY);self.txt(s,'role-b'+str(j),17.0,y+1.4,14.6,1.15,b,21)
  self.txt(s,'other-systems',1.8,17.3,30.5,1.0,'组合系统：空中系统 / 飞镖系统 / 雷达系统',22,GRAY)
  self.note(s,SOURCES+'2027规则印刷页13、20–28：英雄与工程合并为重装；地面机器人为重装、步兵和哨兵。插画为原创通用地面机器人示意，不代表具体赛队结构。')
 def engineering(self,alt=False):
  s=self.d.slide(INK if alt else WHITE);self.tag(s,'工程视角 / 一个动作背后的完整链路',alt)
  self.heading(s,'看见之后，\n还要做出动作。',w=16,c=WHITE if alt else INK)
  self.robot_photo(s,17.8,3.6,14.3)
  for j,(h,b) in enumerate([('感知','获取环境信息'),('决策','判断下一步动作'),('控制','让运动跟上预期'),('执行','把指令变成动作')]):
   x=1.8+j*8.1;self.txt(s,'chain-h'+str(j),x,14.8,7.4,1.4,h,30,RED,HEAVY)
   self.txt(s,'chain-b'+str(j),x,16.6,7.3,1.4,b,21,WHITE if alt else INK)
   if j<3:self.line(s,'chain'+str(j),x+2.5,14.2,x+8.1,14.2,GRAY)
  self.note(s,SOURCES+'此页将赛事涉及的感知、控制、嵌入式与机械工程整理为教学用通用工程链路，不声称所有机器人的软硬件完全相同。插画为原创原生形状。')
 def teamwork(self):
  s=self.d.slide(WHITE);self.tag(s,'团队视角 / 单机之外，才是比赛')
  self.heading(s,'把单机能力，变成团队优势。')
  self.photo(s,'battlefield','field-top',13.5,7.3,19)
  self.line(s,'team-a',8.1,12.7,18.7,9.0,RED)
  self.line(s,'team-b',8.1,12.7,25.2,14.3,CYAN)
  self.robot_photo(s,2.4,9.4,8.1)
  for j,(x,y) in enumerate([(17,9),(26,14)]):self.sh(s,'partner-marker'+str(j),x,y,2,2,'队友',20,WHITE,fill=CYAN,preset='ellipse',align='center',bold=True)
  for j,(x,y,t) in enumerate([(2,15.7,'共享信息'),(14.4,6.0,'协调位置'),(24.4,6.0,'把握时机')]):self.txt(s,'team-label'+str(j),x,y,8.5,1.7,t,28,INK,HEAVY)
  self.note(s,SOURCES+'团队关系为原创教学示意，不是实际对局路线或推荐战术，也不表示2027机器人实际部署位置。底图为规则手册印刷页29图4-1。')
 def iterate(self):
  s=self.d.slide(WHITE);self.tag(s,'赛场之外 / 一整年的工程实践')
  self.heading(s,'让图纸，经得起反复验证。')
  for j,t in enumerate(['设计','制作','联调','复盘']):
   x=1.8+j*8.1;self.txt(s,'loop'+str(j),x,7.3,7.0,1.7,t,31,RED if j==2 else INK,HEAVY)
   if j<3:self.line(s,'loop'+str(j),x+3.2,8.1,x+7.2,8.1,GRAY)
  self.photo(s,'!!team-photo','team-culture',0,10.0,33.87)
  self.note(s,SOURCES+'官方赛事介绍说明团队备赛、综合技术实践与工程师文化。照片直接来自官网赛事文化图片，权利归原作者；四步为本演示的工程实践概括，不是官方固定赛程。')
 def people(self):
  s=self.d.slide(WHITE);self.tag(s,'找到自己的位置 / 从一个模块开始')
  self.heading(s,'每一种专长，\n都有上场的位置。')
  for j,(h,b) in enumerate([('机械','机构与结构'),('电控','硬件与驱动'),('算法','感知与决策'),('运营','组织与传播')]):
   x=1.8+j*8.1;self.txt(s,'people-h'+str(j),x,9.2,7.0,1.7,h,35,RED if j==2 else INK,HEAVY)
   self.txt(s,'people-b'+str(j),x,11.5,7.1,1.5,b,24)
   self.line(s,'people'+str(j),x+2.5,13.8,16.8,16.4,GRAY)
  self.txt(s,'people-summary',4.2,16.8,25.5,1.2,'不同能力，共同完成一台能上场的机器人。',24,INK)
  self.sh(s,'team-joint',16.55,16.15,.5,.5,fill=RED,preset='ellipse')
  self.note(s,SOURCES+'赛事涉及机械、嵌入式、电气、控制、感知、人机交互及团队统筹管理。本页为帮助新成员理解的常见分工概括，不是具体学校的招生承诺。')
 def close(self):
  s=self.d.slide(INK);self.photo(s,'trophy','arena-poster',-6.0,0,28.5)
  self.sh(s,'closing-reading-plane',17,0,16.87,19.05,fill=INK)
  self.txt(s,'closing-brand',18.2,1.2,14.0,1.2,'ROBOMASTER',27,RED,LATIN,bold=True)
  self.txt(s,'closing-head',18.2,4.4,14.0,5.6,'把热爱，\n做成能上场的作品。',37,WHITE,HEAVY)
  self.txt(s,'closing-body',18.3,11.5,13.7,4.3,'看一场完整比赛。\n找到本校战队。\n从一个模块开始。',26,WHITE)
  self.txt(s,'closing-url',18.3,17.0,13.8,1.2,'robomaster.com',27,WHITE,LATIN)
  self.note(s,SOURCES+'奖杯照片来自RoboMaster官网首页视频海报。了解和报名请通过官网当前入口查询，不提供已过期的报名承诺。')

def main():
 p=argparse.ArgumentParser();p.add_argument('--out',type=Path,required=True);p.add_argument('--officecli',required=True);p.add_argument('--pilot',action='store_true');args=p.parse_args()
 out=args.out.resolve()
 if out.exists():raise FileExistsError('Choose a new output directory: '+str(out))
 out.mkdir(parents=True);d=Deck(out/'robomaster.pptx',args.officecli);a=Robo(d)
 if args.pilot:a.cover();a.engineering();a.cover(True);a.engineering(True)
 else:
  a.cover();a.what();a.arena();a.roles();a.engineering();a.teamwork();a.iterate();a.people();a.close()
  a.flush()
  for i in range(1,10):d.run('set',f'/slide[{i}]',props={'transition':'morph-1500' if i in [2,4,5,6] else 'fade-750','advanceClick':True})
 a.finish();print(d.file)

if __name__=='__main__':main()
