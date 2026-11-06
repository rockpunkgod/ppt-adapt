"""OfficeCLI-only theme slot/reference/direct-override experiment; stdlib only."""
import argparse,json,subprocess,shutil
from pathlib import Path

p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--outdir',required=True);a=p.parse_args();out=Path(a.outdir)
if out.exists():p.error('Fresh output directory required')
out.mkdir(parents=True);base=out/'theme-base.pptx';log=[]
def call(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace');log.append({'args':list(map(str,args)),'code':r.returncode,'stdout':r.stdout,'stderr':r.stderr})
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 try:
  v=json.loads(r.stdout)
  if isinstance(v,dict) and v.get('success') is False:raise RuntimeError(r.stdout)
 except json.JSONDecodeError:pass
 return r.stdout
def setprops(file,path,**props):
 args=[]
 for k,v in props.items():args+=['--prop',f'{k}={v}']
 return call('set',file,path,*args)
def add(typ,parent,**props):return {'command':'add','parent':parent,'type':typ,'props':{k:str(v).lower() if isinstance(v,bool) else str(v) for k,v in props.items()}}
def batch(ops):
 f=out/'commands.json';f.write_text(json.dumps(ops,ensure_ascii=False),encoding='utf-8');call('batch',base,'--input',f,'--json')
def shape(slide,name,x,y,w,h,text='',size=22,color='193C40',fill='none',font_ea='Microsoft YaHei',font_latin='Arial',**props):
 q=dict(name=name,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',text=text,size=size,color=color,fill=fill,line='none',margin='0cm',lang='zh-CN',valign='middle')
 q.update({'font.ea':font_ea,'font.latin':font_latin});q.update(props);return add('shape',slide,**q)
call('create',base,'--locale','zh-CN');setprops(base,'/',slideWidth='33.87cm',slideHeight='19.05cm',**{'theme.color.accent1':'1C6561','theme.color.accent2':'607F9A','theme.font.major.eastAsia':'Microsoft YaHei','theme.font.minor.eastAsia':'Microsoft YaHei','theme.font.major.latin':'Arial','theme.font.minor.latin':'Arial'})
batch([add('slide','/',layout='blank',background='F1EEE5')]);s='/slide[1]'
batch([
 shape(s,'heading',1.5,1.1,30,2,'让主题改变整套视觉',size=34,bold=True),
 shape(s,'intro',1.5,3.2,30,1,'统一色彩与字体，保留局部格式。',size=19,color='536B70'),
 shape(s,'theme-card',1.7,6,14.4,7.7,fill='accent1',preset='roundRect'),
 shape(s,'direct-card',17.7,6,14.4,7.7,fill='1C6561',preset='roundRect'),
 shape(s,'theme-card-label',2.6,6.6,12.5,1,'随主题变更',size=20,color='FFFFFF'),
 shape(s,'direct-card-label',18.6,6.6,12.5,1,'局部格式保留',size=20,color='FFFFFF'),
 shape(s,'theme-card-text',2.6,8.4,12.5,2.8,'形与色，让信息清楚',size=29,color='FFFFFF',font_ea='+mj-ea',font_latin='+mj-lt'),
 shape(s,'direct-card-text',18.6,8.4,12.5,2.8,'形与色，让信息清楚',size=29,color='FFFFFF'),
 shape(s,'theme-footer',1.7,15.4,14.5,1.2,'主题引用，更新一次即可联动。',size=19,color='accent1'),
 shape(s,'direct-footer',17.7,15.4,14.5,1.2,'直接格式，维持指定的表达。',size=19),
])
batch([add('slide','/',layout='blank',background='F1EEE5')]);s='/slide[2]'
ops=[shape(s,'heading',1.5,1.1,30,2,'字体的来源，决定修改的范围',size=33,bold=True),shape(s,'intro',1.5,3.2,30,1,'只改标题的东亚主题字体，正文与局部格式保持。',size=19,color='536B70'),shape(s,'theme-rule',1.7,4.7,14.4,0.12,fill='accent1'),shape(s,'direct-rule',17.7,4.7,14.4,0.12,fill='1C6561')]
for label,name,y,ea,latin,text,lang in [('标题主题','major',6.1,'+mj-ea','+mj-lt','让信息清楚可读','zh-CN'),('正文主题','minor',10.2,'+mn-ea','+mn-lt','让信息清楚可读','zh-CN'),('Latin 独立','latin',14.3,'+mj-ea','+mj-lt','Design with purpose','en-US')]:
 ops += [shape(s,name+'-caption',1.7,y,13,0.7,label,size=16,color='536B70'),shape(s,name+'-reference',1.7,y+0.9,14.4,1.8,text,size=29,font_ea=ea,font_latin=latin,lang=lang),shape(s,name+'-direct',17.7,y+0.9,14.4,1.8,text,size=29,lang=lang)]
batch(ops);call('close',base)
# Copy only closed packages; all authoring mutations stay inside OfficeCLI.
color=out/'theme-color.pptx';shutil.copyfile(base,color);setprops(color,'/',**{'theme.color.accent1':'AE4E32'});call('close',color)
font=out/'theme-font.pptx';shutil.copyfile(color,font);setprops(font,'/',**{'theme.font.major.eastAsia':'SimSun'});call('close',font)
for file in [base,color,font]:
 call('validate',file);call('view',file,'issues','--json');call('close',file)
(out/'build-log.json').write_text(json.dumps(log,ensure_ascii=False,indent=2),encoding='utf-8');print(out)
