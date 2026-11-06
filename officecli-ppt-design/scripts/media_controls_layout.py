import argparse,json,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--assets',type=Path,required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);deck=a.outdir/'media-before.pptx'
def cli(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace',check=True)
 if '--json' in args and json.loads(r.stdout).get('success') is False:raise RuntimeError(r.stdout)
 return r.stdout
cli('create',deck);cli('open',deck);cli('set',deck,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
ops=[]
for i,title in enumerate(['Native trim | source 02 to 10 seconds','Native audio fade | 2 second in and out','Native bookmarks | next / previous'],1):
 ops.append(dict(command='add',parent='/',type='slide',props=dict(layout='blank',background='14283F')))
 for name,text,y,size in [('title',title,0.6,27),('note',['Full original 12-second media stays embedded.','The source is constant level; the PowerPoint player must change its level.','Visible source clock must jump to the embedded bookmark times.'][i-1],17.6,15)]:
  ops.append(dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name=name,text=text,x='1.5cm',y=f'{y}cm',width='31cm',height='1cm',font='Arial',size=size,color='FFFFFF',fill='none',line='none',align='center')))
 if i in [1,3]:
  props=dict(name='trim-video' if i==1 else 'bookmark-video',src=str(a.assets/'original-clock-12s.mp4'),poster=str(a.assets/('poster-2s.png' if i==1 else 'poster-0s.png')),x='5.2cm',y='3cm',width='23.5cm',height='13.21875cm',volume=80,autoPlay=True,loop=False)
  if i==1:props.update(trimStart='2s',trimEnd='2s')
  ops.append(dict(command='add',parent=f'/slide[{i}]',type='video',props=props))
 else:
  ops.append(dict(command='add',parent=f'/slide[{i}]',type='audio',props=dict(name='fade-audio',src=str(a.assets/'original-scale-12s.wav'),x='-2cm',y='-2cm',width='1cm',height='1cm',volume=80,autoPlay=True,loop=False,trimStart='2s',trimEnd='2s')))
  ops.append(dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name='fade-label',text='2s fade in  /  4s steady  /  2s fade out',x='3cm',y='8cm',width='28cm',height='2cm',font='Arial',size=32,color='6AAE9D',fill='none',line='none',align='center')))
f=a.outdir/'batch.json';f.write_text(json.dumps(ops),encoding='utf-8');cli('batch',deck,'--input',f,'--json');cli('close',deck);print(cli('validate',deck,'--json'))
