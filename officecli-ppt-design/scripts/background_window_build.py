import sys,zlib,struct,binascii
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parent))
from build_components import Deck
import argparse
p=argparse.ArgumentParser()
p.add_argument('--out',required=True)
p.add_argument('--officecli',required=True)
a=p.parse_args()
folder=Path(a.out).resolve();folder.mkdir(parents=True,exist_ok=False)
def chunk(k,v):return struct.pack('>I',len(v))+k+v+struct.pack('>I',binascii.crc32(k+v)&0xffffffff)
w,h=960,540
raw=bytearray()
for y in range(h):
 raw.append(0)
 for x in range(w):
  raw.extend((int(40+180*x/(w-1)),int(55+150*y/(h-1)),190 if (x//80+y//90)%2 else 110))
png=b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',w,h,8,2,0,0,0))+chunk(b'IDAT',zlib.compress(raw))+chunk(b'IEND',b'')
asset=folder/'coordinate-background.png';asset.write_bytes(png)
d=Deck(folder/'base.pptx',a.officecli);s=d.slide()
d.run('set','/slide[1]',props={'background':'image:'+str(asset)})
d.shape(s,'shade',0,0,33.87,19.05,fill='000000')
d.shape(s,'window-seed',3,6,28,6,'WINDOW',font='Arial',size=100,bold=True,color='FFFFFF')
d.clone(1);d.morph(2);d.finish()
