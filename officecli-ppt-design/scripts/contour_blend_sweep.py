"""Independent raster proof: union of full moving glyphs versus boundary sweep mesh."""
import argparse,json,math
from pathlib import Path
import numpy as np
from PIL import Image,ImageDraw
p=argparse.ArgumentParser();p.add_argument('--model',required=True);p.add_argument('--out',required=True);a=p.parse_args();o=Path(a.out)
if o.exists():raise FileExistsError('Refusing overwrite')
o.mkdir(parents=True);m=json.loads(Path(a.model).read_text());W=1200;H=1200;scale=80

def mask(rings):
 acc=np.zeros((H,W),dtype=np.int8)
 for r in rings:
  signed=sum(r[i][0]*r[(i+1)%len(r)][1]-r[(i+1)%len(r)][0]*r[i][1] for i in range(len(r)))
  im=Image.new('1',(W,H));ImageDraw.Draw(im).polygon([(round(x*scale),round(y*scale)) for x,y in r],fill=1)
  acc+=np.asarray(im,dtype=np.int8)*(1 if signed>0 else -1)
 return acc!=0
P=m['originalCm'];ox,oy=m['shadow']['origin'];dx,dy=m['shadow']['delta']
front=[[[x+ox,y+oy] for x,y in r] for r in P]
shift=lambda rings,t:[[[x+dx*t,y+dy*t] for x,y in r] for r in rings]
oracle=np.zeros((H,W),bool)
for j in range(202):oracle|=mask(shift(front,j/201))
mesh=mask(front)|mask(shift(front,1));count=2
for r in front:
 for i,p0 in enumerate(r):
  p1=r[(i+1)%len(r)]
  if abs((p1[0]-p0[0])*dy-(p1[1]-p0[1])*dx)<1e-8:continue
  mesh|=mask([[p0,p1,[p1[0]+dx,p1[1]+dy],[p0[0]+dx,p0[1]+dy]]]);count+=1
x,y,w,h=m['shadow']['clipCm'];clip=np.zeros((H,W),bool);clip[round(y*scale):round((y+h)*scale)+1,round(x*scale):round((x+w)*scale)+1]=1
oracle&=clip;mesh&=clip
intersection=np.sum(mesh&oracle);union=np.sum(mesh|oracle);iou=float(intersection/union)
# Difference limited to boundary discretization; evaluate erosion independently.
eroded=oracle.copy()
for dy0,dx0 in [(1,0),(-1,0),(0,1),(0,-1)]:eroded &= np.roll(oracle,(dy0,dx0),(0,1))
missInterior=int(np.sum(eroded&~mesh));assert iou>.995 and missInterior==0,(iou,missInterior)
for name,b in [('oracle-202',oracle),('mesh-106',mesh),('xor',mesh^oracle)]:Image.fromarray(b.astype('uint8')*255).save(o/(name+'.png'))
result={'passed':True,'scalePixelsPerCm':scale,'oracleFullGlyphPositions':202,'boundaryMeshPrimitives':count,'clippedIoU':iou,'missedOracleInteriorPixels':missInterior,'differencePixels':int(np.sum(mesh^oracle)),'scope':'Discrete independently sampled full glyph union; native vector and visual checks still required'}
(o/'oracle.json').write_text(json.dumps(result,indent=2));print(json.dumps(result))
