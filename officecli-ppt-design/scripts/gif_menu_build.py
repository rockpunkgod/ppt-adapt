import argparse,json,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--assets',type=Path,required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);deck=a.outdir/'gif-menu.pptx'
def cli(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8')
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
cli('create',deck);cli('open',deck);cli('set',deck,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
ops=[]
def add(parent,type,**props):ops.append(dict(command='add',parent=parent,type=type,props=props))
def shape(s,name,x,y,w,h,**props):
 d=dict(line='none',fill='none');d.update(props);add(f'/slide[{s}]','shape',name=name,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',**d)
add('/','slide',layout='blank',background='101820')
shape(1,'start-title',2,3,29.87,2,text='GIF / PNG state menu',font='Arial',size=36,color='FFFFFF',align='center')
add('/slide[1]','shape',name='start-background',x='10.9cm',y='8cm',width='12cm',height='3cm',geometry='roundRect',fill='24645E',line='none')
shape(1,'start-label',10.9,8,12,3,text='START',font='Arial',size=26,color='FFFFFF',align='center',valign='center')
add('/slide[1]','shape',name='start-hit',x='10.9cm',y='8cm',width='12cm',height='3cm',geometry='roundRect',fill='FFFFFF',opacity=0.03,line='none',link='slide[2]')
for active in range(1,4):
 s=active+1;add('/','slide',layout='blank',background='101820',transition='morph-700',advanceClick=False)
 shape(s,'!!menu-title',2,1,29.87,2,text='One active icon. Two quiet states.',font='Arial',size=30,color='FFFFFF',align='center')
 for icon,cx in enumerate([6.9,16.935,26.97],1):
  src=a.assets/(f'icon-{icon}-active.gif' if icon==active else f'icon-{icon}-inactive.png')
  add(f'/slide[{s}]','picture',name=f'!!icon-{icon}',src=str(src.resolve()),x=f'{cx-1.9}cm',y='5.9cm',width='3.8cm',height='3.8cm')
  shape(s,f'!!label-{icon}',cx-3,11.1,6,1,text=['HOME','WORLD','CREATE'][icon-1],font='Arial',size=20,color='FFFFFF' if icon==active else '6B747C',align='center')
 # The linked circles are deliberately last/topmost, above their GIF/PNG pictures.
 for icon,cx in enumerate([6.9,16.935,26.97],1):
  add(f'/slide[{s}]','shape',name=f'!!circle-{icon}',geometry='ellipse',x=f'{cx-3}cm',y='4.8cm',width='6cm',height='6cm',fill='FFFFFF',opacity=0.03,line=('4CE0D2:2' if icon==active else '737D87:1'),link=f'slide[{icon+1}]')
 shape(s,'!!active-marker',[6.9,16.935,26.97][active-1]-0.5,12.8,1,0.7,geometry='triangle',fill='4CE0D2')
 shape(s,'!!active-caption',2,15.4,29.87,1.4,text=f'ACTIVE {active} | '+['HOME','WORLD','CREATE'][active-1],font='Arial',size=26,color='4CE0D2',align='center')
ops=[x for x in ops if x['type']=='slide']+[x for x in ops if x['type']!='slide']
f=a.outdir/'batch.json';f.write_text(json.dumps(ops),encoding='utf-8');cli('batch',deck,'--input',f,'--json');cli('close',deck);print(cli('validate',deck,'--json'))
