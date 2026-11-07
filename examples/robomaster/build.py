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
 def robot_photo(self,s,x,y,w,key,name):
  self.photo(s,name,key,x,y,w)
 def heading(self,s,t,w=30,c=INK):self.txt(s,'heading',1.7,2.75,w,4.5,t,43,c,HEAVY)
 def line(self,s,n,x1,y1,x2,y2,c=GRAY):
  self.poly(s,'!!knowledge-link-'+n,[(x1,y1),(x2,y2),(x2+.01,y2+.02),(x1+.01,y1+.02)],c)
 def cover(self,alt=False):
  if alt:
   s=self.d.slide(INK);self.photo(s,'cover-photo','competition-scene',0,2.3,33.87)
   self.tag(s,'ROBOMASTER / 机甲大师',True)
   self.txt(s,'alt-title',1.7,3.0,29,4.5,'让工程，成为现场。',56,WHITE,HEAVY)
   self.txt(s,'alt-name',1.7,14.4,31,3.2,'ROBOMASTER',67,WHITE,LATIN,bold=True)
   return s
  s=self.d.slide(WHITE);self.tag(s,'机甲大师 / 认识 RoboMaster')
  self.txt(s,'cover-robo',1.55,3.2,15,4.1,'ROBO',90,INK,LATIN,bold=True)
  self.txt(s,'cover-master',1.55,7.1,15,4.1,'MASTER',90,INK,LATIN,bold=True)
  self.photo(s,'studio-hero','robot-studio',16.3,3.5,17.57)
  self.txt(s,'cover-head',1.75,12.0,14.0,4.0,'把工程，\n带上赛场。',35,INK,HEAVY)
  self.txt(s,'cover-season',1.8,17.0,14.0,1.2,'高校机器人赛事入门',24,RED)
  self.txt(s,'cover-subject',17.3,16.3,15.0,1.3,'2026 · GMaster 步兵机器人',22,INK)
  self.note(s,SOURCES+'主视觉为GMaster战队公开的2026区域赛3号步兵机器人“大哥”定妆原照，6000×4000，https://gmaster.team/robots/infantry-robot/2026-infantry-3 。2027是规则讲解版本，不将历史照片声称为2027已实测车型。')
  return s
 def what(self):
  s=self.d.slide(INK);self.tag(s,'DJI 大疆发起 / 高校机器人赛事',True)
  self.heading(s,'学生造机甲，团队打配合。',c=WHITE)
  self.photo(s,'competition-scene','competition-scene',1.7,6.1,19.35)
  for j,(h,b) in enumerate([('自主研发','高校战队设计并制作机器人'),('战术对抗','在同一赛场协作、交锋'),('工程实践','把多个学科接成可运行系统')]):
   y=7.0+j*3.8;self.txt(s,'what-h'+str(j),22.5,y,9.5,1.5,h,30,RED if j==0 else WHITE,HEAVY)
   text=['高校战队\n设计并制作机器人','在同一赛场\n协作、交锋','把多个学科\n接成可运行系统'][j]
   self.txt(s,'what-b'+str(j),22.5,y+1.6,9.5,2.15,text,22,'C6CDD4')
  self.note(s,SOURCES+'RoboMaster由DJI大疆发起，介绍以高校系列赛中的超级对抗赛为主。其他入口包括高校联盟赛及高校人工智能挑战赛。')
 def arena(self,alt=False):
  s=self.d.slide(WHITE);self.tag(s,'RMUC 2027 / 观赛的第一张地图')
  self.heading(s,'先看目标，\n再看交火。',w=12)
  self.photo(s,'!!field','field-axon',10.0,5.0,23.5)
  for j,t in enumerate(['攻击对方基地','争夺关键阵地','建立全场控制']):self.txt(s,'goal'+str(j),1.8,10.0+j*2.1,11.4,1.6,t,27,RED if j==0 else INK,HEAVY)
  self.note(s,SOURCES+'官方2027规则V1.0.0，印刷页29–30，图4-3。依官网概述，击毁基地或在结束时取得更高全场控制程度可获胜；此页为观赛概述，不替代完整获胜判定。')
 def roles(self):
  s=self.d.slide(WHITE);self.tag(s,'兵种图鉴 / GMaster 2026 区域赛公开车型')
  self.heading(s,'四种分工，四种真实的机器。')
  items=[('英雄','hero-studio','!!role-hero','远程火力支援'),('工程','engineer-studio','!!role-engineer','机构操作与资源获取'),('步兵','robot-studio','role-infantry','机动对抗与协作'),('哨兵','sentry-studio','role-sentry','自动运行与场地协同')]
  for j,(h,key,name,b) in enumerate(items):
   x=1.7+j*8.1;self.photo(s,name,key,x,7.3,7.5)
   self.txt(s,'role-h'+str(j),x,13.2,7.5,1.6,h,31,RED,HEAVY)
   self.txt(s,'role-b'+str(j),x,15.5,7.5,2.0,b,22)
  self.note(s,SOURCES+'本页是2026公开车型图鉴，照片按GMaster档案逐一对应英雄、工程、步兵、哨兵，全部采用整机原图。不是2027出场阵容；下一页单独解释2027规则变化。来源：https://gmaster.team/#robots 。')
 def evolution(self):
  s=self.d.slide(WHITE);self.tag(s,'规则演进 / RMUC 2027')
  self.heading(s,'从英雄、工程，到重装。')
  self.photo(s,'!!role-hero','hero-studio',1.7,7.2,9.0)
  self.photo(s,'!!role-engineer','engineer-studio',11.6,7.2,9.0)
  self.txt(s,'old-hero',1.7,13.6,9,1.5,'2026 · 英雄',25,INK,HEAVY)
  self.txt(s,'old-engineer',11.6,13.6,9,1.5,'2026 · 工程',25,INK,HEAVY)
  self.line(s,'merge-a',6.2,15.1,6.2,15.7,GRAY)
  self.line(s,'merge-b',16.1,15.1,16.1,15.7,GRAY)
  self.line(s,'merge-rail',6.2,15.7,22.0,15.7,GRAY)
  self.line(s,'merge-out',22.0,15.7,23.2,10.3,GRAY)
  self.txt(s,'heavy-role',23.7,7.4,8.6,2.8,'重装',47,RED,HEAVY)
  self.txt(s,'heavy-functions',23.8,11.0,8.5,4.2,'能量单元装配\n团队增益\n模块化发射',24,INK)
  self.txt(s,'current-ground',1.8,17.0,30.2,1.2,'2027 地面机器人：重装 / 步兵 / 哨兵',24,INK)
  self.note(s,SOURCES+'2027规则印刷页13、20–28明确英雄与工程合并为重装。左侧照片仍标识为2026英雄和工程来源车型，用规则演进关系解释新角色，不将旧照片改名冒充2027重装实物。')
 def systems(self):
  s=self.d.slide(WHITE);self.tag(s,'更多角色 / GMaster 2026 公开实物')
  self.heading(s,'地面之外，还有另一套协作。')
  self.photo(s,'aerial-photo','aerial-studio',1.7,8.1,9.6)
  self.photo(s,'dart-photo','dart-studio',22.5,8.1,9.6)
  # Original radar JPEG carries EXIF orientation 8. Keep the pixels unchanged;
  # apply the corresponding CCW quarter-turn as a native PowerPoint transform.
  self.flush();info=IMAGES['radar-studio']
  self.d.run('add',f'/slide[{s}]','--type','picture',props={'name':'radar-photo','src':str(ASSET/info['file']),'x':'12.6cm','y':'7.433333cm','width':'8.6cm','height':'5.733333cm','rotation':270})
  for j,(h,b) in enumerate([('空中机器人','空中支援与侦察'),('雷达系统','全场感知与信息共享'),('飞镖系统','发射装置与目标打击')]):
   x=1.7+j*10.4;self.txt(s,'system-h'+str(j),x,15.2,9.6,1.7,h,29,RED,HEAVY);self.txt(s,'system-b'+str(j),x,17.2,9.6,1.2,b,21)
  self.note(s,SOURCES+'三张照片分别来自GMaster2026公开档案的空中机器人、雷达机器人与飞镖系统，按主体身份对应使用。2027规则将空中、飞镖、雷达归为组合系统；2026照片不冒充2027完整系统配置。雷达原JPEG带EXIF方向8，仅通过PPT对象旋转显示正立，未改动图像内容。')
 def engineering(self,alt=False):
  s=self.d.slide(INK if alt else WHITE);self.tag(s,'工程视角 / 一个动作背后的完整链路',alt)
  self.heading(s,'看见之后，\n还要做出动作。',w=16,c=WHITE if alt else INK)
  self.robot_photo(s,17.8,3.6,14.3,'sentry-studio','sentry-engineering')
  for j,(h,b) in enumerate([('感知','获取环境信息'),('决策','判断下一步动作'),('控制','让运动跟上预期'),('执行','把指令变成动作')]):
   x=1.8+j*8.1;self.txt(s,'chain-h'+str(j),x,14.8,7.4,1.4,h,30,RED,HEAVY)
   self.txt(s,'chain-b'+str(j),x,16.6,7.3,1.4,b,21,WHITE if alt else INK)
   if j<3:self.line(s,'chain'+str(j),x+2.5,14.2,x+8.1,14.2,GRAY)
  self.note(s,SOURCES+'本页采用GMaster2026哨兵整机原照，介绍赛事中的感知、决策、控制与执行链路，不声称所有机器人的软硬件完全相同。未生成或修改机器人结构。')
 def teamwork(self):
  s=self.d.slide(WHITE);self.tag(s,'团队视角 / 单机之外，才是比赛')
  self.heading(s,'把单机能力，变成团队优势。')
  self.photo(s,'battlefield','field-top',13.5,7.3,19)
  self.line(s,'team-a',8.1,12.7,18.7,9.0,RED)
  self.line(s,'team-b',8.1,12.7,25.2,14.3,CYAN)
  self.robot_photo(s,1.7,8.4,11.0,'crew-photo','team-crew-photo')
  for j,(x,y) in enumerate([(17,9),(26,14)]):self.sh(s,'partner-marker'+str(j),x,y,2,2,'队友',20,WHITE,fill=CYAN,preset='ellipse',align='center',bold=True)
  for j,(x,y,t) in enumerate([(2,15.7,'共享信息'),(14.4,6.0,'协调位置'),(24.4,6.0,'把握时机')]):self.txt(s,'team-label'+str(j),x,y,8.5,1.7,t,28,INK,HEAVY)
  self.note(s,SOURCES+'团队照片来自BirdieBot公开图库“冠军团队与机器人”，照片中为2023活动，不将其指认为2027赛队或比赛成绩。关系线为教学示意，不是实际对局路线；底图为规则手册印刷页29图4-1。')
 def iterate(self):
  s=self.d.slide(WHITE);self.tag(s,'赛场之外 / 一整年的工程实践')
  self.heading(s,'让图纸，经得起反复验证。')
  self.photo(s,'workshop-wide','workshop-wide',1.7,8.7,14.7)
  self.photo(s,'workshop-portrait','workshop-portrait',24.2,6.5,8.0)
  for j,t in enumerate(['设计','制作','联调','复盘']):
   self.txt(s,'loop'+str(j),18.1,7.4+j*2.8,5.5,1.7,t,30,RED if j==2 else INK,HEAVY)
  self.note(s,SOURCES+'两张原图来自木鸢BirdieBot战队公开图库：https://birdiebot.org/photo/ ，标题为准备阶段·算法调试、机械组·检修现场。保留原图全部边界和水印。四步为本演示的工程实践概括，不是官方固定赛程。')
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
  a.cover();a.what();a.arena();a.roles();a.evolution();a.systems();a.engineering();a.teamwork();a.iterate();a.people();a.close()
  a.flush()
  for i in range(1,d.count+1):d.run('set',f'/slide[{i}]',props={'transition':'morph-1500' if i==5 else 'fade-750','advanceClick':True})
 a.finish();print(d.file)

if __name__=='__main__':main()
