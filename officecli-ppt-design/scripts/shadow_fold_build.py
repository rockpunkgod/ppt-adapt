import argparse,sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parent))
from build_components import Deck
p=argparse.ArgumentParser();p.add_argument('--stage',choices=['base','final'],required=True);p.add_argument('--out',required=True);p.add_argument('--officecli',required=True);p.add_argument('--shadow');a=p.parse_args()
if a.stage=='final' and (not a.shadow or not Path(a.shadow).is_file()):p.error('Native ellipse PNG required')
d=Deck(a.out,a.officecli)
if a.stage=='base':
 s=d.slide('FFFFFF');d.shape(s,'soft-ellipse',10,2,4,14,fill='555555',geometry='ellipse',softEdge='18pt')
else:
 for j,bg in enumerate(['F4F5F5','751B23']):
  s=d.slide(bg)
  d.shape(s,'number',6,3.5,10,12,'1',font='Arial',size=250,bold=True,color='27313A' if j==0 else 'E5C685',align='center')
  d.run('add',f'/slide[{s}]','--type','picture',props={'name':'half-shadow','src':str(Path(a.shadow).resolve()),'x':'12cm','y':'3cm','width':'2cm','height':'14cm','crop':'0,0,50,0'})
  if j==1:d.run('set',f'/slide[{s}]/picture[@name=half-shadow]',props={'brightness':-100,'opacity':.5})
  d.shape(s,'front-panel',14,3,15,14,fill=bg)
  paths=f'/slide[{s}]/picture[@name=half-shadow],/slide[{s}]/shape[@name=front-panel]'
  d.run('add',f'/slide[{s}]','--type','group',props={'name':'fold','shapes':paths})
  d.run('set',f'/slide[{s}]/group[@name=fold]',props={'rotation':32})
  text='27313A' if j==0 else 'F7ECD5'
  d.shape(s,'eyebrow',19,9,12,1.5,'SUMMARY',size=26,bold=True,color=text)
  d.shape(s,'title',19,10.7,12,1.5,'阶段总结',size=30,bold=True,color=text)
  d.shape(s,'body',19,12.5,12,1.5,'把一条柔边阴影作为视觉分界',size=16,color=text)
  d.notes(s,'Source BV1qj411Y7Up 25–64s: ellipse soft edge, rasterize, crop half, adjacent rectangle same as slide background, group rotate, number behind. This original sample replaces all artwork/text; static effect only, no animation claim.')
d.finish()
