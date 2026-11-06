"""Original native candle groups; animated state plus static control."""
import argparse
from pathlib import Path
from build_components import Deck
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',type=Path,required=True);a=p.parse_args();d=Deck(a.out,a.officecli)
for page in [1,2]:
 s=d.slide('142521');d.shape(s,'heading',1.8,1.4,30,2,'微小的光，也有自己的节奏。',size=31,color='F2EFE5')
 d.shape(s,'subtitle',1.8,3.3,30,1,'三支烛焰 / 纵向呼吸' if page==1 else '相同画面 / 静态对照',size=17,color='B4C5BC')
 for j,x in enumerate([9.8,16.94,24.08]):
  d.shape(s,f'candle-{j}',x-.5,10.6,1,3.8,fill=['416C60','568477','6C9989'][j])
  d.shape(s,f'wick-{j}',x-.04,10.15,.08,.6,fill='16231D')
  d.shape(s,f'flame-outer-{j}',x-.65,8.65,1.3,2.15,fill='F5B34D')
  d.shape(s,f'flame-inner-{j}',x-.28,9.5,.56,1.23,fill='FFF2B7')
  for role in ['outer','inner']:
   name=f'flame-{role}-{j}';ns='http://schemas.openxmlformats.org/drawingml/2006/main'
   geom=f'<a:custGeom xmlns:a="{ns}"><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="w" b="h"/><a:pathLst><a:path w="1000" h="1000"><a:moveTo><a:pt x="650" y="0"/></a:moveTo><a:cubicBezTo><a:pt x="1000" y="450"/><a:pt x="1000" y="1000"/><a:pt x="500" y="1000"/></a:cubicBezTo><a:cubicBezTo><a:pt x="0" y="1000"/><a:pt x="0" y="600"/><a:pt x="650" y="0"/></a:cubicBezTo><a:close/></a:path></a:pathLst></a:custGeom>'
   d.run('raw-set',f'/slide[{s}]','--xpath',f'//p:sp[p:nvSpPr/p:cNvPr[@name="{name}"]]/p:spPr/a:prstGeom','--action','replace','--xml',geom)
  d.run('add',f'/slide[{s}]','--type','group',props={'name':f'flame-{j}','shapes':','.join(f'/slide[{s}]/shape[@name=flame-{role}-{j}]' for role in ['outer','inner'])})
 d.shape(s,'footer',1.8,17.1,30,1,'把运动留给烛焰，让文字安静。',size=16,color='B4C5BC')
 d.notes(s,'Original candles. Native vertical Grow/Shrink70%, 0.2s forward, 0.1s smooth start/end, auto reverse and staggered start. Source 9Ewb0Rd3P3E 4:22:22-4:22:36. Page2 is a static control, not a separate source effect.')
d.run('close');d.run('validate');print(a.out)
