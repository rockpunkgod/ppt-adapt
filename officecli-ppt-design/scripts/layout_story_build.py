"""Standard-library OfficeCLI builder; creates new files only, no PPT libraries."""
import argparse,json,subprocess
from pathlib import Path
W,H=33.87,19.05
INK='16313A';TEAL='117C80';MUTED='64757A';PAPER='F6F4EE';GOLD='DDAC57'
class Deck:
 def __init__(self,exe,path):
  self.exe,self.path,self.n,self.log=exe,path,0,[]
  if path.exists():raise FileExistsError('Refusing to overwrite '+str(path))
  self.call('create',path,'--locale','zh-CN');self.call('set',path,'/','--prop',f'slideWidth={W}cm','--prop',f'slideHeight={H}cm')
 def call(self,*args):
  p=subprocess.run([self.exe,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace');self.log.append({'args':list(map(str,args)),'code':p.returncode,'out':p.stdout,'err':p.stderr})
  if p.returncode:raise RuntimeError(p.stdout+p.stderr)
  return p.stdout
 def add(self,t,s,**p):return {'command':'add','parent':s,'type':t,'props':{k:str(v).lower()if isinstance(v,bool)else str(v)for k,v in p.items()}}
 def batch(self,ops):
  f=self.path.with_suffix('.batch.json');f.write_text(json.dumps(ops,ensure_ascii=False),encoding='utf-8');r=json.loads(self.call('batch',self.path,'--input',f,'--json'))
  if r.get('success') is False:raise RuntimeError(str(r))
 def shape(self,s,n,x,y,w,h,t='',size=22,fill='none',color=INK,**kw):
  props=dict(name=n,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',text=t,font='Microsoft YaHei',size=size,fill=fill,color=color,line='none',margin='0cm');props.update(kw)
  return self.add('shape',s,**props)
 def slide(self,title,tag,bg=PAPER):
  self.n+=1;s=f'/slide[{self.n}]';self.batch([self.add('slide','/',layout='blank',background=bg),self.shape(s,'eyebrow',1.7,.7,29,.7,tag,12,color=TEAL),self.shape(s,'title',1.7,1.9,30,2.2,title,32,bold=True),self.shape(s,'footer',1.7,17.9,30,.7,'原创演示内容 · 非原作者素材／非真实经营数据',10,color=MUTED)])
  return s
 def note(self,s,t):self.batch([self.add('notes',s,text=t)])
 def finish(self):
  self.call('close',self.path)
  self.path.with_suffix('.validation.txt').write_text(self.call('validate',self.path)+'\n'+self.call('view',self.path,'issues','--json'),encoding='utf-8')
  self.path.with_suffix('.build-log.json').write_text(json.dumps(self.log,ensure_ascii=False,indent=2),encoding='utf-8')

def alignment(exe,out):
 d=Deck(exe,out/'alignment.pptx')
 s=d.slide('同一排信息，共用底边与等距间隔','01 / 选择对象对齐');ops=[]
 widths=[7,9,6];bottom=14.8;left=2;right=31.8;gap=(right-left-sum(widths))/2;x=left
 for i,(w,h)in enumerate(zip(widths,[6.8,9,5.4])):
  ops += [d.shape(s,f'card-{i}',x,bottom-h,w,h,f'0{i+1}\n模块',30,fill=[TEAL,INK,GOLD][i],color='FFFFFF',bold=True,align='center',valign='center')];x+=w+gap
 ops+=[d.shape(s,'baseline',2,bottom,29.8,.035,fill=TEAL),d.shape(s,'explain',2,15.4,29.8,1.3,'不同高度保留差异；底边一致，间隙一致。',20,color=MUTED)];d.batch(ops);d.note(s,'Kevin 03:27–05:23: Align Selected Objects vs Align to Slide, bottom align, distribute horizontally. Here uses bounded content area, equal gaps, not equal center distances.')
 s=d.slide('页面居中与组内布局，分开处理','02 / 先组合，再整体移动');ops=[d.shape(s,'logo-bar',4,6,9,1.6,'青屿工作室',26,fill=INK,color='FFFFFF',bold=True,align='center'),d.shape(s,'logo-disc',7,8.1,3,3,preset='ellipse',fill=GOLD),d.shape(s,'logo-tag',3.5,11.5,10,1,'空间 · 服务 · 体验',19,align='center')];d.batch(ops)
 d.batch([d.add('group',s,name='logo-unit',shapes=','.join(f'{s}/shape[@name={n}]'for n in ['logo-bar','logo-disc','logo-tag']))]);d.call('set',d.path,s+'/group[1]','--prop',f'x={(W-10)/2}cm','--prop',f'y={(H-6.5)/2}cm');d.note(s,'Kevin05:25–06:33: group before middle alignment; individual components must not collapse to same center. Group bounding box10×6.5cm moves centered while relative offsets remain.')
 s=d.slide('用固定坐标复用网格，而非逐次目测','03 / 对齐与网格');ops=[]
 for row in range(2):
  for col in range(3):
   x=2+col*10.05;y=5.2+row*5.4;ops+=[d.shape(s,f'grid-{row}-{col}',x,y,9.1,4.5,f'{row*3+col+1:02d}',34,fill='DFEAE5',color=TEAL,bold=True,align='center',valign='center')]
 d.batch(ops);d.note(s,'Kevin06:33–07:53: ruler/grid/guides and guide settings. CLI equivalent is exact shared coordinates; does not claim to toggle PowerPoint View>Guides GUI. Margins2cm; horizontal gap.95cm, vertical gap.9cm.')
 d.finish()

def visual(exe,out):
 d=Deck(exe,out/'visual-design.pptx');s=d.slide('现货可得性是首要购买原因','01 / 结论标题 + 原生可编辑图表')
 d.batch([d.add('chart',s,name='survey-chart',chartType='bar',x='2cm',y='5cm',width='26cm',height='10.8cm',categories='技术支持,线上购买,客户服务,价格,现货',data='受访者占比:25,33,63,70,84',colors=TEAL,legend='none',axismin=0,axismax=100,axisnumfmt='0"%"',dataLabels='outsideEnd',axisfont='15:64757A:Microsoft YaHei',chartborder='none',chartareafill=PAPER,plotareafill=PAPER),d.shape(s,'unit',28,6,4,4,'84%',34,color=TEAL,bold=True)])
 d.note(s,'Leila00:52–01:43: list/data→visual chart with conclusion headline, focus eyes. Values follow video example25/33/63/70/84 as tutorial demonstration, not live survey. Native chart not flattened image.')
 s=d.slide('让主张拥有呼吸空间','02 / 留白可以是任何纯色',INK)
 d.batch([d.shape(s,'hero',3,6.0,23,4.2,'少一点装饰\n多一点重点',44,color='FFFFFF',bold=True),d.shape(s,'sub',3,12,20,2,'一页保留一个可复述的主张。',24,color='C5D9D6'),d.shape(s,'accent',27,7.1,2.5,2.5,fill=GOLD,preset='ellipse')]);d.note(s,'Leila01:46–03:53: negative space any solid/blurred background, clean background, omit unnecessary logos, big text and generous spacing; split excessive content rather than shrink all type.')
 s=d.slide('时间线先理清事件，再安排视觉节奏','03 / 重建布局，不冒充调用Designer');ops=[]
 for i,(yr,t)in enumerate([('2024','验证服务'),('2025','建立团队'),('2026','试点扩张')]):
  x=2.3+i*10;ops +=[d.shape(s,f'line-{i}',x,8.6,9.2,.05,fill='B5CDCA'),d.shape(s,f'dot-{i}',x,8.2,.8,.8,preset='ellipse',fill=TEAL),d.shape(s,f'year-{i}',x,5.8,8,1.5,yr,30,bold=True),d.shape(s,f'event-{i}',x,10,8,2,t,26)]
 d.batch(ops);d.note(s,'Leila03:54–05:14: Home>Design Ideas suggests timeline/collage/bullet redesign; background-baked components may not be editable. This manual editable timeline reconstructs one output principle; no Microsoft Designer service invoked.')
 svg=out/'symbol.svg';svg.write_text('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><path d="M50 4L92 28V73L50 96L8 73V28Z" fill="#117C80"/><path d="M29 52L44 67L73 34" fill="none" stroke="#fff" stroke-width="9" stroke-linecap="round" stroke-linejoin="round"/></svg>',encoding='utf-8')
 s=d.slide('同一矢量符号，在两种尺度保持清晰','04 / SVG资源保留矢量')
 d.batch([d.add('picture',s,name='svg-small',src=str(svg),x='4cm',y='8cm',width='2cm',height='2cm'),d.add('picture',s,name='svg-large',src=str(svg),x='13cm',y='5cm',width='10cm',height='10cm'),d.shape(s,'svg-note',25,8,6,3.5,'2cm → 10cm\n同一SVG源',20,color=MUTED)]);d.note(s,'Leila05:14–06:43: SVG scales mathematically; source also demonstrates convert to Office shape for individual recolor. This validates SVG embedding/scaling only, not SVG→native shape conversion.')
 s=d.slide('让视觉焦点落在三分交点附近','05 / 三分法是一种构图选择');ops=[]
 for i,x in enumerate([W/3,W*2/3]):ops.append(d.shape(s,f'third-v-{i}',x,4.8,.025,11.7,fill='D5DDD6'))
 for i,y in enumerate([H/3,H*2/3]):ops.append(d.shape(s,f'third-h-{i}',1.8,y,30,.025,fill='D5DDD6'))
 ops +=[d.shape(s,'subject',W/3-2,H/3-2,4,4,fill=GOLD,preset='ellipse'),d.shape(s,'quote',W*2/3-3,H/3,10,3,'给重要信息\n一个明确的位置',29,bold=True),d.shape(s,'quote-body',W*2/3-3,H*2/3,10,2,'沿线组织文字，\n保留视线方向的空间。',20,color=MUTED)];d.batch(ops);d.note(s,'Leila06:44–08:04: canvas thirds, subject head near intersection, text along lines, View>Guides. Original abstract subject replaces photo. Grid is visible here as teaching diagram, normally remove from final audience slide.')
 d.finish()

def story(exe,out):
 d=Deck(exe,out/'storytelling.pptx')
 s=d.slide('扩张之前，先明确要解决的决策','01 / SCQ 给答案提供上下文');ops=[]
 for i,(k,head,body)in enumerate([('S','当前情况','青屿服务的预约量持续增长。'),('C','出现变化','现有场地高峰时段已接近满负荷。'),('Q','需要回答','下个季度应先扩建，还是先优化排期？')]):
  y=5+i*3.6;ops +=[d.shape(s,'scq-'+k,2,y,1.5,1.5,k,26,fill=TEAL,color='FFFFFF',bold=True,align='center',valign='center'),d.shape(s,'scq-head-'+k,4.4,y,25,1.1,head,23,bold=True),d.shape(s,'scq-body-'+k,4.4,y+1.3,26,1.4,body,23,color=MUTED)]
 d.batch(ops);d.note(s,'Analyst01:24–02:58: Situation=current state, Complication=change/reason presentation needed, Question follows naturally; introduction sets SCQ, main deck answers. Fictional scenario is adaptation, no real recommendation.')
 s=d.slide('建议先以北区小型试点验证扩张','02 / 答案在前，理由随后');ops=[d.shape(s,'answer',2,5,29.8,3,'先试点，再决定长期扩建',32,fill=INK,color='FFFFFF',bold=True,align='center',valign='center')]
 for i,(h,t)in enumerate([('需求有支撑','北区需求增长更快，\n值得优先验证。'),('运营可控制','限定场次与周期，\n把验证成本留在可承受范围。')]):
  x=2+i*15.3;ops +=[d.shape(s,f'reason-head-{i}',x,10,14.5,1.5,h,28,color=TEAL,bold=True),d.shape(s,f'reason-body-{i}',x,12,14,3,t,24)]
 d.batch(ops);d.note(s,'Analyst04:50–05:58: pyramid main idea first, reasons beneath, granular data beneath reasons. Next two slides expand each reason rather than show disconnected facts.')
 s=d.slide('北区预约增长更快，值得优先验证','03 / 理由一 → 支持数据')
 d.batch([d.add('chart',s,name='north-demand',chartType='column',x='2cm',y='5cm',width='21cm',height='10.6cm',categories='1月,2月,3月,4月',data='北区:80,94,115,138;其他地区:82,86,88,91',colors=TEAL+',C9D2CD',legend='bottom',axismin=0,axismax=160,dataLabels='none',axisfont='14:64757A:Microsoft YaHei',legendFont='14:16313A:Microsoft YaHei',chartareafill=PAPER,plotareafill=PAPER,chartborder='none'),d.shape(s,'north-insight',25,6,6,6,'+73%\n北区预约量\n1月→4月',26,color=TEAL,bold=True)]);d.note(s,'Analyst05:34–06:18: supporting data expands each reason in depth-first slide order. All data invented for structural demonstration; not source tourism data or actual business evidence.')
 s=d.slide('限定规模与周期，让试点成本可控','04 / 理由二 → 可执行的验证边界');ops=[]
 for i,(num,head,body)in enumerate([('01','限定规模','只开放两类高需求服务'),('02','限定周期','连续四周观察预约与复购'),('03','明确门槛','达到利用率与满意度门槛后再扩建')]):
  x=2+i*10;ops +=[d.shape(s,'guard-num-'+num,x,5.4,8,1.8,num,38,color=GOLD,bold=True),d.shape(s,'guard-head-'+num,x,8.0,8.8,1.6,head,28,bold=True),d.shape(s,'guard-body-'+num,x,10.5,8.6,4.4,body,23,color=MUTED)]
 d.batch(ops);d.note(s,'Analyst05:59–06:33: highest idea first then supporting details after each. Closing page is fictional operating detail supporting reason2, not attributed as source-specific prescription.')
 d.finish()

if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--officecli',default='officecli');p.add_argument('--outdir',type=Path,required=True);p.add_argument('--recipe',choices=['alignment','visual','story','all'],default='all');a=p.parse_args()
 if a.outdir.exists():p.error('Output directory already exists; choose a new --outdir')
 a.outdir.mkdir(parents=True)
 recipes={'alignment':alignment,'visual':visual,'story':story};selected=list(recipes.values())if a.recipe=='all'else[recipes[a.recipe]];names={'alignment':'alignment.pptx','visual':'visual-design.pptx','story':'storytelling.pptx'};existing=[str(a.outdir/names[k])for k,v in recipes.items()if v in selected and(a.outdir/names[k]).exists()]
 if existing:p.error('Refusing to overwrite: '+', '.join(existing))
 for f in selected:f(a.officecli,a.outdir);print(f.__name__+' built',flush=True)
