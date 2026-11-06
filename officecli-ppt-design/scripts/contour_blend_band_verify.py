import json
from pathlib import Path
import numpy as np
from PIL import Image,ImageDraw
import argparse
p=argparse.ArgumentParser();p.add_argument('--model',required=True);p.add_argument('--out',required=True);p.add_argument('--geometry-deps',required=True);args=p.parse_args();output=Path(args.out)
if output.exists():raise FileExistsError('Refusing overwrite')
m=json.loads(Path(args.model).read_text());S=200;W=900;H=1050
X,Y=np.meshgrid((np.arange(W)+.5)/S-.05,(np.arange(H)+.5)/S-.05);dist=np.full((H,W),np.inf)
for ring in m['originalCm']:
 for i,(x0,y0) in enumerate(ring):
  x1,y1=ring[(i+1)%len(ring)];dx=x1-x0;dy=y1-y0;t=np.clip(((X-x0)*dx+(Y-y0)*dy)/(dx*dx+dy*dy),0,1);dist=np.minimum(dist,np.hypot(X-x0-t*dx,Y-y0-t*dy))
acc=np.zeros((H,W),np.int8)
for ring in m['strokeCm']:
 signed=sum(ring[i][0]*ring[(i+1)%len(ring)][1]-ring[(i+1)%len(ring)][0]*ring[i][1] for i in range(len(ring)))
 im=Image.new('1',(W,H));ImageDraw.Draw(im).polygon([(round((x+.05)*S),round((y+.05)*S)) for x,y in ring],fill=1);acc+=np.asarray(im,np.int8)*(1 if signed>0 else -1)
import sys
sys.path.insert(0,args.geometry_deps)
from shapely.geometry import Polygon
from shapely.ops import unary_union
from shapely import contains_xy
rings=m['strokeCm'];areas=[sum(r[i][0]*r[(i+1)%len(r)][1]-r[(i+1)%len(r)][0]*r[i][1] for i in range(len(r))) for r in rings]
positive=[r for r,area in zip(rings,areas) if area>0];negative=[r for r,area in zip(rings,areas) if area<0]
polys=[Polygon(r,[h for h in negative if Polygon(r).contains(Polygon(h).representative_point())]) for r in positive]
actual=contains_xy(unary_union(polys),X,Y);expected=dist<=.03
falseInterior=int(np.sum(actual&(dist>.038)));missingInterior=int(np.sum((dist<.022)&~actual));iou=float(np.sum(actual&expected)/np.sum(actual|expected))
r={'scale':S,'halfWidthCm':.03,'toleranceCm':.008,'unexpectedOutsidePixels':falseInterior,'missingInnerBandPixels':missingInterior,'rasterIoU':iou,'passed':falseInterior==0 and missingInterior==0,'method':'Independent nearest-point distance to original segments vs saved-band polygon point containment; no Pillow edge-rounding'}
output.write_text(json.dumps(r,indent=2));print(r)
assert r['passed']
