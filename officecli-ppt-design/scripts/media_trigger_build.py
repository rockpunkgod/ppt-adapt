"""Cake-triggered real GIF and title-triggered native video on the same slide."""
from pathlib import Path
import argparse,sys
sys.path.insert(0,str(Path(__file__).resolve().parent))
from build_components import Deck
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--assets',type=Path,required=True);p.add_argument('--out',type=Path,required=True);a=p.parse_args();d=Deck(a.out,a.officecli)
for i in [1,2]:
 s=d.slide('16251D');d.shape(s,'background-disc',18,3,13,13,fill='4D6E66',geometry='ellipse')
 d.run('add',f'/slide[{s}]','--type','video',props=dict(name='confetti-video',src=str((a.assets/'confetti.mp4').resolve()),poster=str((a.assets/'poster.png').resolve()),x='0cm',y='0cm',width='33.87cm',height='19.05cm',autoPlay=False,loop=False,volume=0))
 d.shape(s,'kicker',1.5,1.1,30,1,'TRACE / 指定媒体触发' if i==1 else 'TRACE / 不透明对照',font='Noto Sans SC',size=15,color='DAEF78')
 d.shape(s,'happy-birthday',1.5,5.3,16,4.5,'HAPPY\nBIRTHDAY',font='Arial',size=40,color='F2F3EC')
 d.shape(s,'instruction',1.5,11.2,14,3,'点击蛋糕，点亮烟花。\n点击标题，让庆祝开始。',font='Noto Sans SC',size=19,color='BCC7BB')
 names=[]
 for name,x,y,w,h,color in [('cake-plate',20.5,15,10,.3,'F2F3EC'),('cake-bottom',21,12.6,9,2.3,'F4CE6A'),('cake-filling',21,13.5,9,.4,'F2F3EC'),('cake-icing',21,12,9,1.1,'F2F3EC')]:names.append(name);d.shape(s,name,x,y,w,h,fill=color)
 for j,c in enumerate(['EF3C85','58D9CE','DAEF78']):
  x=22.4+j*2.7;name='candle-'+str(j);names.append(name);d.shape(s,name,x,10.6,.45,1.4,fill=c)
  name='flame-'+str(j);names.append(name);d.shape(s,name,x-.02,10.05,.48,.65,fill='F4CE6A',geometry='teardrop')
 d.run('add',f'/slide[{s}]','--type','group',props={'name':'cake','shapes':','.join(f'/slide[{s}]/shape[@name={n}]' for n in names)})
 d.run('add',f'/slide[{s}]','--type','picture',props=dict(name='fireworks-gif',src=str((a.assets/'fireworks.gif').resolve()),x='20cm',y='3.4cm',width='10cm',height='7.5cm'))
 d.shape(s,'footer',1.5,17.3,31,.9,'蛋糕与标题各自触发，空白处不会开始。' if i==1 else '同一段素材，不加透明强调动画。',font='Noto Sans SC',size=14,color='BCC7BB')
 d.notes(s,'Original birthday scene. Two independent shape triggers coexist: Cake→Appear GIF; title→Fly FromTop + Play + Transparency70%. Second page omits only Transparency as an opaque-video control. No source assets.')
d.run('close');print(a.out)
