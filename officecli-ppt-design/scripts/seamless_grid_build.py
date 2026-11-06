"""Editable two-period native grid; OfficeCLI authoring only."""
import argparse,sys,re
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parent))
from build_components import Deck,W,H
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);a=p.parse_args()
d=Deck(a.out,a.officecli,static=True);s=d.slide('183443')
d.shape(s,'period-bounds',0,0,2*W,H,fill='none',line='none')
original=d.run('get',f'/slide[{s}]/shape[@name=period-bounds]')['data']['results'][0]['format']['id']
ops=[]
def rect(name,x,y,w,h):
 ops.append({'command':'add','parent':f'/slide[{s}]','type':'shape','props':dict(name=name,x=f'{x:.6f}cm',y=f'{y:.6f}cm',width=f'{w:.6f}cm',height=f'{h:.6f}cm',fill='FFFFFF',opacity='0.2',line='none')})
stroke=0.5*2.54/72
for j in range(30):rect(f'column-edge-{j}',j*W/15-stroke/2,0,stroke,H)
for j in range(9):rect(f'row-edge-{j}',0,j*H/8-stroke/2,2*W,stroke)
d.batch(ops)
target=f'/slide[{s}]/shape[@name=period-bounds]'
d.run('add',target,'--type','animation',props={'effect':'fade','class':'entrance','trigger':'withPrevious','duration':1000})
d.run('add',target,'--type','animation',props={'class':'motion','path':'custom','d':'M 0 0 L -1 0 E','trigger':'withPrevious','duration':17000})
d.run('add',f'/slide[{s}]','--type','group',props={'name':'seamless-grid','shapes':','.join(str(i) for i in range(1,41))})
group=d.run('get',f'/slide[{s}]/group[@name=seamless-grid]')['data']['results'][0]['path']
group_id=re.search(r'@id=(\d+)',group).group(1)
d.run('raw-set',f'/slide[{s}]','--xpath',f"//p:spTgt[@spid='{original}']",'--action','setattr','--xml',f'spid={group_id}')
d.run('raw-set',f'/slide[{s}]','--xpath',"//p:cTn[@presetClass='path']",'--action','setattr','--xml','repeatCount=indefinite')
d.shape(s,'heading',1.5,1.2,30,1.8,'A CONTINUOUS BACKGROUND',font='Arial',size=34,bold=True,color='FFFFFF')
d.shape(s,'subtitle',1.5,3.1,30,1,'A steady rhythm behind a clear message.',font='Arial',size=18,color='D0DDE2')
d.finish();print(d.file)
