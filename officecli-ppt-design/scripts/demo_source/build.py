from pathlib import Path
import sys
import argparse
parser=argparse.ArgumentParser()
parser.add_argument('--work',type=Path,required=True)
parser.add_argument('--officecli',required=True)
args=parser.parse_args()
SKILL=Path(__file__).resolve().parents[2]
O=args.work.resolve()
sys.path.insert(0,str(SKILL/'scripts'))
from build_components import Deck
from case_pages import Author,P,I,A,M,B,H,L
d=Deck(O/'main.pptx',args.officecli);a=Author(d,SKILL/'assets/demo/aura-device.png');sh=a.sh
def title(s,t,dark=False):sh(s,'title',1.6,1.35,30.6,3.0,t,37,P if dark else I,font=H,bold=True)
def note(s,t):a.note(s,t)
# Project cover: actual slide outputs are the product visual.
s=d.slide(I)
a.pic(s,'cover-back',21.1,2.7,13.4,7.5375,O/'native-directions/slide-2.png')
a.pic(s,'cover-front',16.1,6.1,17.7,9.95625,O/'native-directions/slide-1.png')
sh(s,'project-brand',1.6,1.05,24,1.3,'PPT ADAPT',27,A,font=L,bold=True)
sh(s,'cover-headline',1.6,5.2,15.2,7.1,'把一段需求，\n做成完整发布。',50,P,font=H,bold=True)
sh(s,'cover-body',1.6,13.7,14.3,3.0,'让 Agent 组织内容、决定设计，\n并交付能继续修改的演示。',25,'C3CBD6')
note(s,'本演示介绍PPT Adapt，以本次原创概念产品任务的实际成稿作为主视觉和能力证据。右侧图片为本次PowerPoint真实导图，不是虚构的软件界面。')
# Input is explicit and honest, rather than an invented AI baseline.
s=d.slide(P);title(s,'用一次完整任务，看它能交出什么。')
sh(s,'input-label',1.6,4.8,21,1.0,'本次演示任务',23,A,bold=True)
sh(s,'brief-rule',1.6,6.1,.1,10.2,fill=A)
sh(s,'brief',2.3,6.15,21.2,10.3,'为一款离线工作的边缘AI模块，\n做三页产品发布。\n\n面向产品团队，依次讲清：\n它是什么、怎样工作、用在哪里。\n\n产品要突出，画面要连贯，\n内容还要能继续修改。',26,I)
sh(s,'brief-count',26,5.6,6.2,4.5,'03',97,I,font=L)
sh(s,'brief-output',26,10.5,6.3,5.5,'页产品发布\n\n同一条叙事\n同一套视觉',23,I)
note(s,'AURA及这段需求均为本次演示创建的概念任务，用于实际制作三页连贯发布。没有伪造客户、商业效果或其他AI的对照结果。')
# Three consecutive, usable product-presentation pages.
for k in range(3):a.case(k)
# Same material, different art direction: actual candidate output.
s=d.slide(P);title(s,'同样的资料，可以有不同的设计答案。')
a.pic(s,'style-dark',1.6,5.3,14.8,8.325,O/'native-directions/slide-1.png')
a.pic(s,'style-light',17.4,5.3,14.8,8.325,O/'native-directions/slide-3.png')
sh(s,'style-dark-caption',1.6,14.4,14.7,1.5,'聚焦产品的深色发布',27,I,font=H,bold=True)
sh(s,'style-light-caption',17.4,14.4,14.7,1.5,'强调信息的浅色出版',27,I,font='Noto Serif SC',bold=True)
sh(s,'style-conclusion',1.6,16.6,30.5,1.3,'字体、构图与素材位置，都按这次表达重新决定。',24,I)
note(s,'两张代表页是本次实际制作并由PowerPoint导出的不同艺术方向，使用同一概念产品与同一原创素材。候选不是固定模板库。')
# A real, editable chart is concrete evidence of handoff quality.
s=d.slide(P);title(s,'图表交付后，数据还能继续改。')
sh(s,'chart-subtitle',1.6,4.4,29.9,1.25,'PPT Adapt 配方按执行方式分布',24,'596273')
a.flush();d.run('add','/slide[7]','--type','chart',props=dict(name='editable-chart',chartType='bar',categories='OfficeCLI,PowerPoint COM,Hybrid,External',data='配方条目:33,1,47,5',x='1.3cm',y='6.0cm',width='22.0cm',height='10.1cm',legend='none',title='none',colors=A,chartareafill='none',chartborder='none',plotareafill='none',axisfont='18:'+I+':Noto Sans SC',axismin=0,gridlines='D9DDE3:0.6'))
sh(s,'xlsx-badge',25,7.0,6.9,1.7,'XLSX',28,P,fill=A,font=L,align='center',bold=True)
sh(s,'workbook-message',25,10.0,7.0,5.0,'原生图表\n真实数据源\n工作簿随文件保存',26,I,font=H)
note(s,'使用仓库recipe-index的真实四类条目计数33/1/47/5。图表为原生bar，后续沿用已有工作簿补齐机制绑定真实嵌入xlsx；这里证明可编辑交付，不把配方数量等同作品质量。')
# Explain why the demonstrated result is repeatable, using the real project chain.
s=d.slide(I);title(s,'这套成果，有一条能跑完的制作链。',True)
sh(s,'pipeline-line',3.0,9.0,27,.07,fill='4D5564')
for j,(en,h,b,tool) in enumerate([
 ('PLAN','先设计整套','读懂受众与内容，\n确定叙事、视觉与页面职责。','设计总控'),
 ('BUILD','再组合制作','调用已有机制，\n把页面与动效落实成文件。','OfficeCLI + COM'),
 ('PLAY','预览后修正','查看原生画面和播放，\n修正问题，再交出作品。','PowerPoint')]):
 x=1.6+j*10.6;sh(s,f'pipe-word{j}',x,5.4,9.4,3,en,54,P,font=L,bold=True)
 sh(s,f'pipe-dot{j}',x+4.25,8.65,.75,.75,fill=A,preset='ellipse')
 sh(s,f'pipe-head{j}',x,10.4,9.4,1.6,h,29,P,font=H,bold=True)
 sh(s,f'pipe-body{j}',x,12.6,9.4,3.0,b,22,'C3CBD6')
 sh(s,f'pipe-tool{j}',x,16.0,9.4,1.3,tool,23,A,font=L)
note(s,'依据现有项目真实工作链：全稿设计、已有配方组合、OfficeCLI主体制作、必要COM精修、PowerPoint原生预览。没有新增引擎或改写既有配方的验收边界。')
# A deliverable is a usable file, not a slogan about creativity.
s=d.slide(P);title(s,'最终拿走的，是你能继续改的整套 PPT。')
sh(s,'file-paper',2.0,5.9,11.8,10.6,fill='FFFFFF',line='D1D6DF:1')
sh(s,'file-fold',10.6,5.9,3.2,3.2,fill='E6EAF0',preset='rtTriangle',rotation=180)
sh(s,'file-word',3.0,9.2,9.5,3.5,'PPTX',69,A,font=L,bold=True)
sh(s,'file-sub',3.0,13.2,9.7,1.7,'完整演示 · 可以接手',24,I)
for j,(h,b) in enumerate([('改文字与版式','按自己的内容增删、替换、重排。'),('改图表与数据','在文件内继续编辑真实图表。'),('改图片与动效','调整素材、裁切和原生时间轴。')]):
 y=6.1+j*3.9;sh(s,f'edit-head{j}',17.2,y,14.9,1.5,h,30,I,font=H,bold=True);sh(s,f'edit-body{j}',17.2,y+1.9,14.7,1.4,b,23,'596273')
note(s,'本文件保留原生文字、形状、图表、嵌入工作簿和时间轴。生成产品图是PNG，可替换和裁切，不把它称为可直接编辑的3D模型。')
# Return to the user's next task with a concrete project entry point.
s=d.slide(I);sh(s,'closing-brand',1.6,1.1,25,1.3,'PPT ADAPT',28,A,font=L,bold=True)
a.pic(s,'closing-work',18.7,5.6,15.5,8.71875,O/'native-directions/slide-1.png')
sh(s,'closing-head',1.6,5.2,18.0,6.6,'下一次发布，\n从你的内容开始。',47,P,font=H,bold=True)
sh(s,'closing-action',1.6,13.9,27,1.7,'给出主题、受众、场景与素材，开始制作。',26,'C3CBD6')
sh(s,'project-url',1.6,16.4,30.3,1.2,'github.com/rockpunkgod/ppt-adapt',24,A,font=L)
note(s,'交付与行动页。真实项目地址为https://github.com/rockpunkgod/ppt-adapt。核心价值由前面的完整任务、实际作品、可编辑图表和真实制作链支持。')
a.flush()
for n in [1,2,3,6,7,8,9,10]:d.run('set',f'/slide[{n}]',props={'transition':'fade-650','advanceClick':True})
d.run('close');d.run('validate');d.run('close');print('10-page project presentation with a complete three-page case ready')
