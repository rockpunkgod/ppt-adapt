"""True donut spotlight, OfficeCLI only; native playback still required."""
import argparse
from build_components import Deck
p=argparse.ArgumentParser();p.add_argument('--out',required=True);p.add_argument('--officecli',required=True);a=p.parse_args()
d=Deck(a.out,a.officecli);s=d.slide('E8EDF2')
centers=[6.65,16.94,27.22]
for i,(x,title,color) in enumerate(zip(centers,['事实','判断','行动'],['356D62','35669B','9B6943'])):
 d.shape(s,f'person-{i}',x-4.6,6,9.2,9.2,fill=color,geometry='ellipse')
 d.shape(s,f'label-{i}',x-3.3,9.5,6.6,1.8,title,size=36,bold=True,color='FFFFFF',align='center')
# For a square donut: inner radius = outer radius - side*adj/100000.
# 100cm side and adj45200 gives a 9.6cm diameter true hole.
d.shape(s,'!!circle-spotlight',centers[0]-50,10.6-50,100,100,fill='101820',opacity=.82,geometry='donut',adj='adj:val 45200')
d.shape(s,'heading',1.5,1.2,31,1.8,'让观众跟随当前重点',size=36,bold=True,color='FFFFFF')
d.shape(s,'caption',1.5,16.5,31,1,'原生空心圆超出画布，孔内保留底层内容。',size=18,color='FFFFFF')
for x in centers[1:]:
 t=d.clone(d.count);d.set(t,'!!circle-spotlight',x=f'{x-50}cm');d.morph(t)
for t in range(1,4):d.notes(t,'Source BV12D4y1J7YU 00:14–00:40: oversized native donut, adjustable hole, opacity, slide duplication and Morph. Own text and native shapes replace source portraits; no 3D lamp claimed. Outside-slide donut bounds are intentional.')
d.finish()
