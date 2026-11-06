import argparse,json,math,struct,zlib,hashlib
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);W,H=640,360
def png(name,data):
 def chunk(t,b):return struct.pack('>I',len(b))+t+b+struct.pack('>I',zlib.crc32(t+b)&0xffffffff)
 raw=b''.join(b'\0'+bytes(sum(data[y*W:(y+1)*W],())) for y in range(H))
 out=b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',W,H,8,6,0,0,0))+chunk(b'IDAT',zlib.compress(raw,9))+chunk(b'IEND',b'');(a.outdir/name).write_bytes(out);return hashlib.sha256(out).hexdigest()
src=[]
for y in range(H):
 for x in range(W):
  if not (24<=x<616 and 24<=y<336):v=(0,0,0,0)
  else:
   v=(round(20+180*x/640),round(30+170*y/360),round(210-130*x/640),255)
   if (x-190)**2+(y-175)**2<95**2:v=(240,180,38,255)
   if 285<x<450 and 75<y<275:v=(25,190,160,255)
   if abs(y-(315-0.42*x))<5:v=(245,245,245,255)
   if x%64<3 or y%48<3:v=(45,52,70,255)
   if 485<=x<590 and 65<=y<135:v=(80,160,240,128)
  src.append(v)
def sample(data,x,y):return data[y*W+x] if 0<=x<W and 0<=y<H else (0,0,0,0)
split=[]
for y in range(H):
 for x in range(W):
  r=sample(src,x-10,y);c=sample(src,x+10,y);alpha=max(r[3],c[3]);split.append((round(r[0]*r[3]/alpha),round(c[1]*c[3]/alpha),round(c[2]*c[3]/alpha),alpha) if alpha else (0,0,0,0))
bands=[dict(y0=90,y1=101,dx=28),dict(y0=166,y1=174,dx=-36),dict(y0=247,y1=261,dx=20)]
def overlay(d,s):
 alpha=s[3]/255;da=d[3]/255;oa=alpha+da*(1-alpha)
 return tuple(round((s[c]*alpha+d[c]*da*(1-alpha))/oa) for c in range(3))+(round(oa*255),) if oa else (0,0,0,0)
slices=split.copy()
for band in bands:
 for y in range(band['y0'],band['y1']):
  for x in range(W):slices[y*W+x]=overlay(slices[y*W+x],sample(split,x-band['dx'],y))
# Explicit directional streak approximation, not Photoshop's undocumented Wind/Blast algorithm.
wind=slices.copy()
for y in range(H):
 for x in range(W):
  if 80<=y<280 and y%9<2:
   samples=[sample(slices,x-k,y) for k in range(19)];v=max(samples,key=lambda q:(q[0]+q[1]+q[2])*q[3]);wind[y*W+x]=v
hashes={name:png(name,data) for name,data in [('original.png',src),('channels.png',split),('slices.png',slices),('streak-approx.png',wind)]}
(a.outdir/'model.json').write_text(json.dumps({'width':W,'height':H,'channels':{'RsourceDx':-10,'GBsourceDx':10,'alpha':'max source alpha; unpremultiply each separated contribution by union alpha'},'bands':bands,'bandOperation':'copy selected bands then translate and alpha-over base, preserving original underneath','streak':'Original bounded rightward 19px max-brightness extension on rows y%9<2, y80..279; NOT Photoshop Wind/Blast equivalence','hashes':hashes},indent=2),encoding='utf-8');print(json.dumps(hashes,indent=2))
