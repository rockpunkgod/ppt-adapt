"""Independent native alpha/source-over and locality checks, not brush algorithm emulation."""
import argparse,json,zipfile,xml.etree.ElementTree as E,hashlib,math
from pathlib import Path
import numpy as np
from PIL import Image
p=argparse.ArgumentParser();p.add_argument('--deck',required=True);p.add_argument('--model',required=True);p.add_argument('--native',required=True);p.add_argument('--out',required=True);a=p.parse_args();out=Path(a.out)
if out.exists():raise FileExistsError('Refusing overwrite')
m=json.loads(Path(a.model).read_text());ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
with zipfile.ZipFile(a.deck) as z:
 assert not any('/media/' in n for n in z.namelist()),'Native field must not be raster image'
 roots=[E.fromstring(z.read('ppt/slides/slide%d.xml'%i)) for i in (1,2,3)];maps=[{s.find('p:nvSpPr/p:cNvPr',ns).get('name'):s for s in root.findall('p:cSld/p:spTree/p:sp',ns)} for root in roots]
 for c in m['stamps']:
  s=maps[c['slide']-1][c['name']];assert s.find('p:spPr/a:prstGeom',ns).get('prst')=='ellipse';g=s.find('p:spPr/a:gradFill',ns);assert g.find('a:path',ns).get('path')=='circle';stops=g.findall('a:gsLst/a:gs',ns);assert len(stops)==6
  for st,(pos,t) in zip(stops,m['stops']):
   assert int(st.get('pos'))==pos;cl=st.find('a:srgbClr',ns);assert cl.get('val')==c['color'];assert int(cl.find('a:alpha',ns).get('val'))==round(t*c['alpha']*100000)
  xf=s.find('p:spPr/a:xfrm',ns);wh=xf.find('a:ext',ns);assert int(wh.get('cx'))==int(wh.get('cy'))
 native=json.loads((Path(a.native)/'native-objects.json').read_text(encoding='utf-8-sig'))
 assert len(native['states'])==2
 for state in native['states']:
  assert len(state['shapes'])==20
  assert all(s['type']==1 and s['fillType']==3 and s['stopCount']==6 for s in state['shapes'])
  for row in state['shapes']:
   c=next(c for c in m['stamps'] if c['slide']==row['slide'] and c['name']==row['name']);rgb=sum(int(c['color'][k:k+2],16)*256**(k//2) for k in [0,2,4])
   for stop,(pos,t) in zip(row['stops'],m['stops']):
    assert abs(stop['position']-pos/100000)<.00002 and abs(stop['transparency']-(1-t*c['alpha']))<.00002 and stop['rgb']==rgb
imgs=[np.asarray(Image.open(Path(a.native)/('slide-%d.png'%i)).convert('RGB'),float) for i in (1,2,3)];H,W=imgs[0].shape[:2];sx=W/m['widthCm'];sy=H/m['heightCm'];white=np.array([255.,255.,255.]);ca=np.array([240.,107.,99.]);cb=np.array([19.,180.,192.])
def pixel(i,x,y):return imgs[i][round(y*sy),round(x*sx)]
def infer_alpha(measured,color):d=white-color;return float(np.dot(d,white-measured)/np.dot(d,d))
fall=[]
for r,v in [(0,1),(.25,.72),(.5,.35),(.75,.10),(.98,.008),(1.03,0)]:
 observed=infer_alpha(pixel(1,4.45+3.8*r,10.2),ca);expected=.7*v;assert abs(observed-expected)<.015,(r,observed,expected);fall.append({'radiusFraction':r,'alphaObserved':observed,'alphaExpected':expected})
composite=[];orderDiff=[]
for dx in [-.9,0,.9]:
 for dy in [-.7,0,.7]:
  aa=infer_alpha(pixel(1,5+dx,10+dy),ca);ab=infer_alpha(pixel(1,13+dx,10+dy),cb)
  expAB=cb*ab+(ca*aa+white*(1-aa))*(1-ab);expBA=ca*aa+(cb*ab+white*(1-ab))*(1-aa)
  actAB=pixel(1,21+dx,10+dy);actBA=pixel(1,29+dx,10+dy);errAB=float(np.max(abs(actAB-expAB)));errBA=float(np.max(abs(actBA-expBA)));assert max(errAB,errBA)<=3,(dx,dy,errAB,errBA)
  orderDiff.append(float(np.max(abs(actAB-actBA))));composite.append({'local':[dx,dy],'alphaA':aa,'alphaB':ab,'normalABMaxError':errAB,'normalBAMaxError':errBA})
assert max(orderDiff)>30,'Opposite layer order must have measurable effect'
# Exact same geometry/all other colors; one local fill color changes.
for name in ['cold-cyan','warm-gold','mid-lime','mid-blue','mid-violet','mid-magenta','warm-light']:
 c1=maps[0][name];c3=maps[2][name];assert E.tostring(c1.find('p:spPr/a:xfrm',ns))==E.tostring(c3.find('p:spPr/a:xfrm',ns))
 g1=c1.find('p:spPr/a:gradFill',ns);g3=c3.find('p:spPr/a:gradFill',ns)
 if name!='mid-lime':assert E.tostring(g1)==E.tostring(g3)
changed=np.max(abs(imgs[0]-imgs[2]),axis=2)>2;target=next(c for c in m['stamps'] if c['slide']==1 and c['name']=='mid-lime');xx,yy=np.meshgrid(np.arange(W)/sx,np.arange(H)/sy);support=(xx-target['cx'])**2+(yy-target['cy'])**2 <= (target['w']/2+.04)**2
annotation=(xx>2.1)&(xx<13.6)&(yy>6.5)&(yy<7.5);outside=int(np.sum(changed&~support&~annotation));inside=int(np.sum(changed&support&~annotation));assert outside==0 and inside>50000,(outside,inside)
result={'passed':True,'nativeGradientObjects':20,'stopsPerObject':6,'nativeReopenPasses':2,'nativeGradientStopValueChecks':240,'nativeRadialFalloffChecks':fall,'normalAlphaSamples':composite,'maximumLayerOrderDifferenceRGB':max(orderDiff),'singleLocalColorEdit':{'changedWithinSupport':inside,'changedOutsideSupport':outside},'pictures':0,'scope':'Circular editable same-color alpha stamps and normal source-over; does not claim Photoshop brush exact falloff or arbitrary elliptical scaling'}
out.write_text(json.dumps(result,indent=2));print(json.dumps({'passed':True,'maxNormalError':max(max(x['normalABMaxError'],x['normalBAMaxError']) for x in composite),'orderDiff':max(orderDiff),'localChangedPixels':inside}))
