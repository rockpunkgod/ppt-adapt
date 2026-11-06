"""Build a real editable line+area chart with OfficeCLI; stdlib only."""
import argparse,subprocess,json
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args()
if a.output.exists():raise FileExistsError(a.output)
a.output.parent.mkdir(parents=True,exist_ok=True)
def call(*x):
 r=subprocess.run([a.officecli,*map(str,x)],capture_output=True,text=True,encoding='utf-8',errors='replace')
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
call('create',a.output);call('open',a.output)
call('set',a.output,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
ops=[{'command':'add','parent':'/','type':'slide','props':{'layout':'blank','background':'F6F8FC'}}]
def txt(name,text,x,y,w,h,size,color,bold=False):
 ops.append({'command':'add','parent':'/slide[1]','type':'shape','props':dict(name=name,text=text,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',font='Arial',size=size,color=color,bold=bold,fill='none',line='none',margin='0cm',autoFit='none')})
txt('eyebrow','SALES REVIEW / DEMONSTRATION DATA',2.4,1.3,29,1,12,'63758F')
txt('title','A steadier climb through the first half',2.4,2.6,29,1.7,30,'17324E',True)
txt('unit','Revenue (CNY million)',2.5,5,24,1,16,'3E5872')
txt('footnote','Illustrative data. Each point is an observed monthly value; no interpolation is implied.',2.5,17.6,29,1,12,'63758F')
ops.append({'command':'add','parent':'/slide[1]','type':'chart','props':{'name':'revenue-line-area','chartType':'combo','combotypes':'area,line','categories':'Jan,Feb,Mar,Apr,May,Jun','data':'Revenue fill:24,29,27,35,42,46;Revenue:24,29,27,35,42,46','x':'2.5cm','y':'6cm','width':'28.8cm','height':'10.8cm','title':'none','legend':'bottom','legendFont':'14:3E5872:Arial','axisfont':'14:3E5872:Arial','axismin':'0','axismax':'60','majorunit':'10','chartareafill':'none','chartborder':'none','plotareafill':'none','plotborder':'none','gridlines':'DCE4EE:0.5','colors':'C8DAEF,2678BD'}})
ops.append({'command':'add','parent':'/slide[1]','type':'notes','props':{'text':'Mechanism from BV1VTDGYsEmd P16 1623-1998s. Two identical native series: area below line on the same axes. Original content and synthetic data. Do not smooth discrete observations or imply that the decorative area is a second measure.'}})
f=a.output.with_suffix('.batch.json');f.write_text(json.dumps(ops),encoding='utf-8');call('batch',a.output,'--input',f,'--json')
call('set',a.output,'/slide[1]/chart[1]/series[2]','--prop','lineWidth=3','--prop','marker=circle:9:FFFFFF','--prop','smooth=false')
call('close',a.output);r=call('validate',a.output,'--json');a.output.with_suffix('.validation.json').write_text(r,encoding='utf-8');print(r)
