"""Editable vertical Push continuation, with identical cross-page rail ends."""
import argparse,sys
from pathlib import Path
from build_components import Deck,W,H
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);a=p.parse_args()
d=Deck(a.out,a.officecli,static=True)
routes=[[(4,5),(10,5),(10,10),(24,10),(24,H)],[(24,0),(24,3),(10,3),(10,13),(24,13),(24,H)],[(24,0),(24,3),(10,3),(10,11),(28,11)]]
for i,points in enumerate(routes,1):
 s=d.slide('F6F8FC')
 d.shape(s,'heading',1.7,.8,29,1.3,f'CHAPTER {i}',font='Arial',size=30,bold=True,color='17324E')
 d.shape(s,'subheading',1.7,2.2,25,.6,['From research to a first working model','From validation to a reliable release','From release to continued improvement'][i-1],font='Arial',size=14,color='527086')
 d.shape(s,'route',0,0,W,H,fill='none',line='C04457:3')
 coords=[(round(x/W*100000),round(y/H*100000)) for x,y in points]
 path=f'<a:moveTo><a:pt x="{coords[0][0]}" y="{coords[0][1]}"/></a:moveTo>'+''.join(f'<a:lnTo><a:pt x="{x}" y="{y}"/></a:lnTo>' for x,y in coords[1:])
 geom=f'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="w" b="h"/><a:pathLst><a:path w="100000" h="100000" fill="none">{path}</a:path></a:pathLst></a:custGeom>'
 d.run('raw-set',f'/slide[{s}]','--xpath','//p:sp[p:nvSpPr/p:cNvPr[@name="route"]]/p:spPr/a:prstGeom','--action','replace','--xml',geom)
 for j,(x,y) in enumerate(points[1:-1],1):
  d.shape(s,f'node-{j}',x-.18,y-.18,.36,.36,fill='C04457',geometry='ellipse')
  label_y=y+.3 if y<=3 else y-.95
  d.shape(s,f'label-{j}',x+.4,label_y,7,.7,f'{i}.{j} / '+['Observe','Test','Refine','Decide'][j-1],font='Arial',size=14,color='17324E')
 if i>1:d.run('set',f'/slide[{s}]',props={'transition':'push-up-1400','advanceClick':True})
d.finish()
