import argparse,json,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--video',type=Path,required=True);p.add_argument('--poster',type=Path,required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);deck=a.outdir/'portal-before.pptx'
def cli(*args):return subprocess.run([a.officecli,*map(str,args)],check=True,capture_output=True,text=True,encoding='utf-8').stdout
cli('create',deck);cli('open',deck);cli('set',deck,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
ops=[]
for i in [1,2]:
 ops.append(dict(command='add',parent='/',type='slide',props=dict(layout='blank',background='111827')))
 if i==2:ops.append(dict(command='add',parent='/slide[2]',type='video',props=dict(name='portal-video',src=str(a.video.resolve()),poster=str(a.poster.resolve()),x='0cm',y='0cm',width='33.87cm',height='19.05cm',volume=0,autoPlay=True,loop=False)))
 ops.append(dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name='portal-mask',preset='rect',x='0cm',y='0cm',width='33.87cm',height='19.05cm',fill='753D82',line='none')))
 for name,text,y,h,size,font in [('portal-title','H',4.0,11.05,250,'Impact'),('portal-caption','A NEW HORIZON',1.9,1.5,24,'Arial')]:
  ops.append(dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name=name,text=text,x='3cm',y=f'{y}cm',width='27.87cm',height=f'{h}cm',font=font,size=size,color='FFFFFF',fill='none',line='none',align='center',valign='middle',margin=0)))
f=a.outdir/'batch.json';f.write_text(json.dumps(ops),encoding='utf-8');cli('batch',deck,'--input',f,'--json');cli('close',deck);print(cli('validate',deck,'--json'))
