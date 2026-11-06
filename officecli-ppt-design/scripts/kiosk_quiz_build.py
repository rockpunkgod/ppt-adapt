"""Original two-slide quiz foundation, OfficeCLI only, stdlib only."""
import argparse,subprocess,json
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--outdir',required=True);a=p.parse_args();out=Path(a.outdir)
if out.exists():p.error('Fresh output directory required')
out.mkdir(parents=True);file=out/'quiz-base.pptx';log=[]
def call(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace');log.append({'args':list(map(str,args)),'code':r.returncode,'stdout':r.stdout,'stderr':r.stderr})
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 try:
  value=json.loads(r.stdout)
  if isinstance(value,dict) and value.get('success') is False:raise RuntimeError(r.stdout)
 except json.JSONDecodeError:pass
 return r.stdout
def add(typ,parent,**props):return {'command':'add','type':typ,'parent':parent,'props':{k:str(v).lower() if isinstance(v,bool) else str(v) for k,v in props.items()}}
def shape(s,name,x,y,w,h,text='',size=24,color='193C40',fill='none',**props):
 q=dict(name=name,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',text=text,size=size,color=color,fill=fill,line='none',margin='0cm',font='Microsoft YaHei',valign='middle',lang='zh-CN');q.update(props);return add('shape',s,**q)
def batch(ops):
 f=out/'commands.json';f.write_text(json.dumps(ops,ensure_ascii=False),encoding='utf-8');call('batch',file,'--input',f,'--json')
call('create',file,'--locale','zh-CN');call('set',file,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
batch([add('slide','/',layout='blank',background='F1EEE5'),add('slide','/',layout='blank',background='F1EEE5')])
s='/slide[1]';ops=[shape(s,'heading',1.5,1.1,30,1.9,'先判断，再继续',size=34,bold=True),shape(s,'intro',1.5,3.1,30,1,'一道算式，三个选择。',size=19,color='536B70'),shape(s,'question',1.7,4.7,30.5,2.5,'2 + 3 = ?',size=56,align='center')]
for i,value in enumerate([4,5,6]):ops.append(shape(s,f'answer-{value}',2.2+i*10.1,8.4,9.2,3.3,str(value),size=36,color='FFFFFF',fill='193C40',preset='roundRect',align='center'))
ops += [shape(s,'instruction',2.2,12.8,29,1.4,'选错可重试；答对后才能继续。',size=20,color='536B70'),shape(s,'next',25.6,15.2,6,2.2,'继续',size=24,color='FFFFFF',fill='21645E',preset='roundRect',align='center',link='slide[2]')]
batch(ops);s='/slide[2]';batch([shape(s,'result-title',1.5,1.2,30,2,'判断完成',size=34,bold=True),shape(s,'result-math',2,5.6,29.8,3.2,'2 + 3 = 5',size=58,align='center',color='21645E'),shape(s,'result-note',3.4,10.4,27,3,'错误反馈留在题目页。\n正确反馈解锁下一步。',size=27,align='center'),shape(s,'result-footer',3.4,15.7,27,1,'按 Esc 退出演示。',size=18,color='536B70',align='center')])
# Keep the hyperlink on the button shape, not duplicated on its text run (blue/underline override).
call('raw-set',file,'/slide[1]','--xpath',"//p:sp[p:nvSpPr/p:cNvPr[@name='next']]/p:txBody//a:hlinkClick",'--action','remove','--json')
for s in ['/slide[1]','/slide[2]']:call('set',file,s,'--prop','advanceClick=false','--prop','advanceTime=0')
call('close',file);call('validate',file);call('view',file,'issues','--json');call('close',file)
(out/'build-log.json').write_text(json.dumps(log,ensure_ascii=False,indent=2),encoding='utf-8');print(file)
