import argparse,math,struct,zlib,json
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);W,H=640,360;paper=(248,246,239);ink=(24,30,26)
def write(name,data):
 def chunk(t,b):return struct.pack('>I',len(b))+t+b+struct.pack('>I',zlib.crc32(t+b)&0xffffffff)
 raw=b''.join(b'\0'+bytes(sum(data[y*W:(y+1)*W],())) for y in range(H));(a.outdir/name).write_bytes(b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',W,H,8,6,0,0,0))+chunk(b'IDAT',zlib.compress(raw,9))+chunk(b'IEND',b''))
# Original hand-selected variable-width polyline, plus known faint/gray edge diagnostic strokes.
points=[(80+390*t,170+85*math.sin(2.4*math.pi*t),2+12*(1-t)**0.8) for t in [i/400 for i in range(401)]]
alpha=[0]*(W*H)
for cx,cy,r in points:
 for y in range(max(0,int(cy-r-2)),min(H,int(cy+r+3))):
  for x in range(max(0,int(cx-r-2)),min(W,int(cx+r+3))):
   cov=max(0,min(1,(r+1-math.hypot(x+0.5-cx,y+0.5-cy))/2));idx=y*W+x;alpha[idx]=max(alpha[idx],round(255*cov))
for y in range(H):
 for x in range(W):
  idx=y*W+x
  if 75<=x<550 and abs(y-300)<1.5:alpha[idx]=max(alpha[idx],12)
  if 520<=x<523 and 60<=y<240:alpha[idx]=max(alpha[idx],64)
  if 570<=x<573 and 60<=y<240:alpha[idx]=max(alpha[idx],192)
scan=[];neutral=[];truth=[]
for i,v in enumerate(alpha):
 rgb=tuple(round(paper[c]*(1-v/255)+ink[c]*v/255) for c in range(3));scan.append(rgb+(255,));neutral.append((255-v,)*3+(255,));truth.append((255,255,255,v))
write('simulated-scan.png',scan);write('neutral-probe.png',neutral);write('truth-alpha.png',truth)
(a.outdir/'model.json').write_text(json.dumps({'kind':'Original procedural simulated scan, NOT physical scan, font raster, InkML or author asset','paper':paper,'ink':ink,'width':W,'height':H,'faintAlpha':12,'mediumAlpha':64,'darkAlpha':192,'noise':'none; calibrated uniform paper only','scope':'Synthetic clean uniform paper and known single ink color; no arbitrary real-paper shadow/texture/photo guarantee'},indent=2),encoding='utf-8')
