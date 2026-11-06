"""OfficeCLI authors real rounded offsets and outlined filled stroke geometry."""
import argparse,json,math,sys
from pathlib import Path
from build_components import Deck,W
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--geometry',required=True);p.add_argument('--out',required=True);a=p.parse_args()
data=json.loads(Path(a.geometry).read_text(encoding='utf-8-sig'));u=data['unitsToCm'];b=data['bounds'];margin=.8/u
if data.get('text')!='6' or data.get('font')!='Arial Black' or abs(u*b['height']-11)>1e-6:
 raise ValueError('Validated fixture only: Arial Black 6 at 11 cm')
vieww=b['width']+2*margin;viewh=b['height']+2*margin
width,height=vieww*u,viewh*u;x=(W-width)/2;y=4.2
def geom(contours):
 paths=[]
 for contour in contours:
  pts=[(round((px-b['x']+margin)*1000),round((py-b['y']+margin)*1000)) for px,py in contour]
  cmd=f'<moveTo><pt x="{pts[0][0]}" y="{pts[0][1]}"/></moveTo>'
  cmd+=''.join(f'<lnTo><pt x="{px}" y="{py}"/></lnTo>' for px,py in pts[1:])+'<close/>'
  paths.append(cmd)
 return '<custGeom xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><avLst/><gdLst/><ahLst/><cxnLst/><rect l="0" t="0" r="w" b="h"/><pathLst><path w="'+str(round(vieww*1000))+'" h="'+str(round(viewh*1000))+'">'+''.join(paths)+'</path></pathLst></custGeom>'
d=Deck(a.out,a.officecli,static=True)
def pathshape(s,name,contours):
 d.shape(s,name,x,y,width,height,fill='1C6561')
 xml=geom(contours)
 d.batch([{'command':'raw-set','part':f'/slide[{s}]','xpath':f'//p:sp[p:nvSpPr/p:cNvPr[@name="{name}"]]/p:spPr/a:prstGeom','action':'replace','xml':xml}])
def page(title,caption):
 s=d.slide('F6F8FC');d.heading(s,title,caption)
 d.shape(s,'source-text',29,15.3,2,1,'6',font=data['font'],size=20,color='17324E')
 d.shape(s,'footnote',1.6,17.8,30,.7,'Original text retained. Geometry is an independent editable result.',font='Arial',size=12,color='527086')
 return s
s=page('ORIGINAL GLYPH','The original counter and concave contours remain our reference.');pathshape(s,'original-glyph',data['original'])
for stroke in data['offsets']:
 mm=stroke['offsetMm'];s=page(f'OUTWARD OFFSET / {mm:g} mm','Rounded polygon offsets. Closed contours remain editable.')
 pathshape(s,'offset-result',stroke['contours'])
s=page('STROKE CONVERTED TO FILLED PATHS','A 4 mm band follows both sides of the original contour. No line style.');pathshape(s,'outlined-stroke',data['outlinedStroke'])
d.finish()
