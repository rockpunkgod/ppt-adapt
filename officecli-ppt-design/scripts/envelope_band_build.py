import argparse,json,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--outdir',type=Path,required=True);p.add_argument('--ellipse-cap',action='store_true');a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);deck=a.outdir/'envelope-before.pptx'
def cli(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8')
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
cli('create',deck);cli('open',deck);cli('set',deck,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm');ops=[]
for i,title in enumerate(['BOLD | original editable text backup','BOLD | real native glyph contours','BOLD | custom asymmetric ribbon envelope','BOLD | second custom monotone envelope'],1):
 if a.ellipse_cap and i==3:title='BOLD | pointed ellipse-cap envelope'
 if a.ellipse_cap and i==4:title='BOLD | positive-height band comparison'
 ops.append(dict(command='add',parent='/',type='slide',props=dict(layout='blank',background='F4F0E8')))
 ops.append(dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name='title',text=title,x='1.5cm',y='0.7cm',width='30.87cm',height='1.5cm',font='Arial',size=24,color='172A35',fill='none',line='none')))
 if i<=2:ops.append(dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name='editable-backup' if i==1 else 'glyph-seed',text='BOLD',x='2cm',y='6cm',width='29.87cm',height='8cm',font='Arial',bold=True,size=170,color='172A35',fill='none',line='none',align='center',valign='middle',margin=0)))
 else:
  for name,fill,line in [('target-boundary','E5DED0','9A8E7B:1'),('warped-glyph','172A35','none')]:ops.append(dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name=name,geometry='rect',x='3cm',y='5cm',width='27.87cm',height='9cm',fill=fill,line=line)))
  ops.append(dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name='scope',text='Custom native geometry | monotone top/bottom boundary | text backup on slide 1',x='2cm',y='16.2cm',width='29.87cm',height='1.5cm',font='Arial',size=16,color='6C655B',fill='none',line='none',align='center')))
f=a.outdir/'batch.json';f.write_text(json.dumps(ops),encoding='utf-8');cli('batch',deck,'--input',f,'--json');cli('close',deck);print(cli('validate',deck,'--json'))
