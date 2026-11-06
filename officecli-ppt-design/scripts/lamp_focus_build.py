"""Compose existing picture-intersection, native model and Morph mechanisms."""
import argparse,math
from pathlib import Path
from layout_story_build import Deck
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--image',type=Path,required=True);p.add_argument('--model',type=Path,required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
if not a.image.is_file() or not a.model.is_file():raise FileNotFoundError('Local image and original lamp model required')
a.outdir.mkdir(parents=True);d=Deck(a.officecli,a.outdir/'lamp-focus-base.pptx');ns='http://schemas.openxmlformats.org/drawingml/2006/main'
for i,(x,word,body) in enumerate([(7.62,'感知','从环境中捕捉有效信号'),(16.94,'判断','让多个信号形成清楚判断'),(26.25,'响应','把判断落实为及时行动')],1):
 s=f'/slide[{i}]';d.batch([d.add('slide','/',layout='blank',background='18262B')])
 # Reorder source scene below the standard header, then add darkening layer.
 d.batch([d.add('picture',s,name='scene',src=str(a.image.resolve()),x='0cm',y='0cm',width='33.87cm',height='19.05cm'),d.shape(s,'dimmer',0,0,33.87,19.05,fill='09181D',opacity='.72'),d.add('picture',s,name='focus-image',src=str(a.image.resolve()),x='0cm',y='0cm',width='33.87cm',height='19.05cm'),d.shape(s,'focus-circle',x-2.35,8.47-2.35,4.7,4.7,fill='FFFFFF',preset='ellipse')])
 # Native triangle points upward; rotate its whole object to the chosen focus.
 apex=(16.94,14.35);dx=x-apex[0];dy=8.47-apex[1];length=math.hypot(dx,dy);angle=math.degrees(math.atan2(dx,-dy))
 d.batch([d.shape(s,'!!light-beam',apex[0]-2.1,apex[1]-length,4.2,length,fill='FFFFFF',preset='triangle')])
 xp='//p:sp[p:nvSpPr/p:cNvPr[@name="!!light-beam"]]/p:spPr'
 geom=f'<a:custGeom xmlns:a="{ns}"><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="w" b="h"/><a:pathLst><a:path w="100000" h="100000"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100000" y="0"/></a:lnTo><a:lnTo><a:pt x="50000" y="100000"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>'
 d.call('raw-set',d.path,s,'--xpath',xp+'/a:prstGeom','--action','replace','--xml',geom)
 grad=f'<a:gradFill xmlns:a="{ns}" rotWithShape="1"><a:gsLst>'+''.join(f'<a:gs pos="{pos}"><a:srgbClr val="FFFFFF"><a:alpha val="{alpha}"/></a:srgbClr></a:gs>' for pos,alpha in [(0,0),(42000,8000),(100000,65000)])+'</a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>'
 d.call('raw-set',d.path,s,'--xpath',xp+'/a:solidFill','--action','replace','--xml',grad)
 # Rotation is completed around the emitting apex by the native finishing step.
 d.batch([d.add('model3d',s,name='!!lamp',src=str(a.model.resolve()),x='14.94cm',y='13.7cm',width='4cm',height='4.4cm',rotation=f'25,0,{-angle}'),d.shape(s,'heading',1.7,1.3,30,1.8,'让每一个重点，都被看见',32,color='F1F0E8'),d.shape(s,'step',1.7,3.1,29,.8,f'0{i} / {word}',18,color='DDB067'),d.shape(s,'message',1.7,17.4,29,.8,body,18,color='C8D7D8')])
 d.note(s,f'Original LUMA scene. Native picture-circle Intersect, editable 3-stop triangular simulated beam, real original GLB above beam; Morph across focus states. Source BV1pa4y117yV 59-131s. Not volumetric lighting. Beam angle={angle}; apex={apex}.')
 d.call('set',d.path,s,'--prop','transition=morph-1500')
d.finish();print(d.path)
