import argparse,json,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--wav',type=Path,required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);deck=a.outdir/'lifecycle-before.pptx'
def cli(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8',check=True)
 return r.stdout
cli('create',deck);cli('open',deck);cli('set',deck,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
ops=[]
for i,(word,bg) in enumerate([('START','183E58'),('CONTINUE','156C64'),('STOP','8E3448')],1):
 ops.append(dict(command='add',parent='/',type='slide',props=dict(layout='blank',background=bg)))
 for name,text,y,size in [('title',word,6,52),('label','Original 12-second scale | stop after two displayed slides',10,22),('page',f'{i} / 3',15,20)]:
  ops.append(dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name=name,text=text,x='1cm',y=f'{y}cm',width='31.87cm',height='2cm',font='Arial',size=size,color='FFFFFF',fill='none',line='none',align='center')))
 if i==1:ops.append(dict(command='add',parent='/slide[1]',type='audio',props=dict(name='lifecycle-audio',src=str(a.wav.resolve()),x='-2cm',y='-2cm',width='1cm',height='1cm',autoPlay=True,loop=False,volume=80)))
f=a.outdir/'batch.json';f.write_text(json.dumps(ops),encoding='utf-8');cli('batch',deck,'--input',f,'--json');cli('close',deck);print(cli('validate',deck,'--json'))
