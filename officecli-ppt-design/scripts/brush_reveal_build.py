"""Compose editable brush silhouettes with the shared-background recipe."""
import argparse,math,random,xml.etree.ElementTree as E
from pathlib import Path
from build_components import Deck
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--image',type=Path,required=True);p.add_argument('--out',type=Path,required=True);a=p.parse_args()
if not a.image.is_file():raise FileNotFoundError(a.image)
d=Deck(a.out,a.officecli);ns='http://schemas.openxmlformats.org/drawingml/2006/main';q=lambda t:'{'+ns+'}'+t
for page in [1,2,3]:
 s=d.slide('image:'+str(a.image.resolve()))
 if page>1:d.run('add',f'/slide[{s}]','--type','picture',props=dict(name='blurred-cover',src=str(a.image.resolve()),x='0cm',y='0cm',width='33.87cm',height='19.05cm'))
 if page==3:
  for j,(x,y,w) in enumerate([(2.5,4.3,15),(2.5,7,17),(4,9.7,17),(15.5,4.9,14),(16,7.6,15),(14,10.3,16)]):
   name=f'brush-{j}';d.shape(s,name,x,y,w,3.2,fill='FFFFFF')
   rng=random.Random(700+j);top=[(k/240,.08+rng.random()*.04+.12*math.exp(-min(k,240-k)/3)) for k in range(241)];bottom=[(k/240,.91+rng.random()*.04-.12*math.exp(-min(k,240-k)/3)) for k in range(240,-1,-1)]
   points=top+bottom
   root=E.Element(q('custGeom'))
   for n in ['avLst','gdLst','ahLst','cxnLst']:E.SubElement(root,q(n))
   E.SubElement(root,q('rect'),l='0',t='0',r='w',b='h');lst=E.SubElement(root,q('pathLst'));path=E.SubElement(lst,q('path'),w='100000',h='100000')
   for k,(px,py) in enumerate(points):E.SubElement(E.SubElement(path,q('moveTo' if k==0 else 'lnTo')),q('pt'),x=str(round(px*100000)),y=str(round(py*100000)))
   E.SubElement(path,q('close'))
   d.run('raw-set',f'/slide[{s}]','--xpath',f'//p:sp[p:nvSpPr/p:cNvPr[@name="{name}"]]/p:spPr/a:prstGeom','--action','replace','--xml',E.tostring(root,encoding='unicode'))
   d.run('set',f'/slide[{s}]/shape[@name={name}]',props=dict(rotation='-13'))
  d.shape(s,'reveal-title',3.2,13.6,28,2.2,'看见信号，才能向前。',size=32,color='F3F0E5')
 else:d.shape(s,'reference-label',1.7,1.4,30,1.4,'清晰原图' if page==1 else '虚化与降亮度对照',size=25,color='163F36' if page==1 else 'F3F0E5')
 d.notes(s,'Original shared-coordinate image. Page1 sharp reference; page2 native artistic Blur30 and brightness-54%; page3 six editable brush silhouettes use slide-background fill, alternating entrance Wipe500ms, then title Fade. One click starts sequence. BV1jJ4m1g71D 44-101s; no author assets, no alpha-PNG substitution.')
d.run('close');d.run('validate');print(a.out)
