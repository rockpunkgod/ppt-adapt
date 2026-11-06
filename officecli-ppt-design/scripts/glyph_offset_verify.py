"""Check actual PPT vector distances, raster dilation and outlined-stroke fill."""
import argparse,json,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import numpy as np
from PIL import Image
from scipy import ndimage
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def world_contours(shape):
 x=shape.find('p:spPr/a:xfrm',N);off=x.find('a:off',N);ext=x.find('a:ext',N)
 assert x.get('rot','0')=='0'
 ox,oy,cx,cy=[float(v)/360000 for v in [off.get('x'),off.get('y'),ext.get('cx'),ext.get('cy')]]
 result=[]
 for path in shape.findall('p:spPr/a:custGeom/a:pathLst/a:path',N):
  pw,ph=float(path.get('w')),float(path.get('h'));current=[]
  for node in path:
   name=node.tag.split('}')[-1]
   if name in ['moveTo','lnTo']:
    pt=node.find('a:pt',N);current.append([ox+float(pt.get('x'))/pw*cx,oy+float(pt.get('y'))/ph*cy])
   elif name=='close':result.append(np.array(current));current=[]
   else:raise AssertionError('Expected actual closed polygonal geometry')
  assert not current
 assert result
 return result
def distances(points,contours):
 aa=np.concatenate(contours);bb=np.concatenate([np.roll(c,-1,axis=0) for c in contours]);delta=bb-aa;den=np.sum(delta*delta,axis=1);den[den==0]=1
 out=[]
 for batch in np.array_split(points,max(1,len(points)//128)):
  rel=batch[:,None,:]-aa[None,:,:];t=np.clip(np.sum(rel*delta[None,:,:],axis=2)/den[None,:],0,1)
  out.extend(np.sqrt(np.min(np.sum((rel-t[:,:,None]*delta[None,:,:])**2,axis=2),axis=1)))
 return np.array(out)
def vector(deck):
 shapes=[];reports=[]
 with zipfile.ZipFile(deck) as z:
  for page,name in [(1,'original-glyph'),(2,'offset-result'),(3,'offset-result'),(4,'offset-result'),(5,'outlined-stroke')]:
   root=E.fromstring(z.read(f'ppt/slides/slide{page}.xml'));assert not root.findall('.//p:pic',N)
   candidates=[s for s in root.findall('p:cSld/p:spTree/p:sp',N) if s.find('p:nvSpPr/p:cNvPr',N).get('name')==name];assert len(candidates)==1
   shape=candidates[0];assert shape.find('p:spPr/a:ln/a:noFill',N) is not None
   backup=[s for s in root.findall('p:cSld/p:spTree/p:sp',N) if s.find('p:nvSpPr/p:cNvPr',N).get('name')=='source-text'];assert len(backup)==1
   assert ''.join(backup[0].find('p:txBody',N).itertext())=='6'
   shapes.append(world_contours(shape))
 for page,expected in [(2,.2),(3,.4),(4,.6),(5,.2)]:
  contours=shapes[page-1];points=np.concatenate([np.concatenate([c,(c+np.roll(c,-1,axis=0))/2]) for c in contours])
  delta=distances(points,shapes[0])-expected;error=float(np.max(np.abs(delta)))
  assert error<.001,'Every vertex and edge midpoint must remain equidistant'
  reports.append({'page':page,'distanceCm':expected,'checkedBoundarySamples':len(points),'maxDistanceErrorCm':error,'closedContours':len(contours)})
 return reports
def mask(path):
 a=np.array(Image.open(path).convert('RGB')).astype(int);assert a.shape==(900,1600,3)
 return np.max(np.abs(a-[28,101,97]),axis=2)<6
def raster(folder):
 masks=[mask(folder/f'page-{i}.png') for i in range(1,6)];original=masks[0];outside=ndimage.distance_transform_edt(~original);inside=ndimage.distance_transform_edt(original);reports=[]
 for mm,expanded in zip([2,4,6],masks[1:4]):
  expected=mm/10*1600/33.87;edge=expanded&~ndimage.binary_erosion(expanded)
  missing=int(np.count_nonzero(original&~expanded));error=float(np.max(np.abs(outside[edge]-expected)))
  assert missing==0 and error<=1.5
  reports.append({'offsetMm':mm,'expectedDistancePx':expected,'missingOriginalPixels':missing,'allBoundaryMaxErrorPx':error})
 half_width=.2*1600/33.87;absolute=np.where(original,inside,outside);band=masks[4]
 missing=int(np.count_nonzero((absolute<=half_width-1.5)&~band));extra=int(np.count_nonzero((absolute>half_width+1.5)&band))
 assert missing==0 and extra==0
 yy,xx=np.where(original);center=np.array([(yy.min()+yy.max())/2,(xx.min()+xx.max())/2]);factor=(yy.max()-yy.min()+2*half_width)/(yy.max()-yy.min())
 fake=ndimage.affine_transform(original.astype(np.uint8),np.eye(2)/factor,offset=center*(1-1/factor),order=0)>0
 fake_edge=fake&~ndimage.binary_erosion(fake);fake_missing=int(np.count_nonzero(original&~fake));fake_error=float(np.max(np.abs(outside[fake_edge]-half_width)))
 assert fake_missing>0 or fake_error>1.5,'A font-size-scale negative must be rejected'
 return {'passed':True,'offsets':reports,'outlinedStroke':{'missedConfidentBandPixels':missing,'extraFarPixels':extra,'halfWidthPx':half_width},'fontScaleNegative':{'rejected':True,'missingOriginal':fake_missing,'maxErrorPx':fake_error}}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('deck');p.add_argument('--native',type=Path,required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
 if a.report.exists():raise FileExistsError(a.report)
 result={'passed':True,'vectors':vector(a.deck),'raster':raster(a.native),'scope':'Original fixed Arial Black 6, 1600x900, round 2/4/6mm offsets and 4mm stroke; not an arbitrary-font guarantee.'}
 a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(a.report)
