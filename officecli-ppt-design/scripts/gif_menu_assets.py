import argparse,binascii,json,math,struct,subprocess,zlib
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--ffmpeg',required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);W=160
def png(path,pix):
 def chunk(t,d):return struct.pack('>I',len(d))+t+d+struct.pack('>I',binascii.crc32(t+d)&0xffffffff)
 scan=b''.join(b'\0'+bytes(pix[y*W*4:(y+1)*W*4]) for y in range(W))
 path.write_bytes(b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',W,W,8,6,0,0,0))+chunk(b'IDAT',zlib.compress(scan))+chunk(b'IEND',b''))
for icon in range(1,4):
 folder=a.outdir/f'icon-{icon}';folder.mkdir()
 for phase in range(12):
  pix=bytearray(W*W*4)
  def dot(x,y,r,color):
   for yy in range(max(0,int(y-r)),min(W,int(y+r+1))):
    for xx in range(max(0,int(x-r)),min(W,int(x+r+1))):
     if (xx-x)**2+(yy-y)**2<=r*r:pix[(yy*W+xx)*4:(yy*W+xx)*4+4]=bytes(color)
  def line(x1,y1,x2,y2,width=2):
   n=max(1,int(math.hypot(x2-x1,y2-y1)*2))
   for j in range(n+1):dot(x1+(x2-x1)*j/n,y1+(y2-y1)*j/n,width,(255,255,255,255))
  if icon==1:
   for q in [(44,80,80,47),(80,47,116,80),(53,73,53,111),(53,111,107,111),(107,111,107,73),(73,111,73,88),(73,88,88,88),(88,88,88,111)]:line(*q)
  elif icon==2:
   for j in range(96):
    t=j*2*math.pi/96;dot(80+32*math.cos(t),80+32*math.sin(t),2,(255,255,255,255));dot(80+14*math.cos(t),80+32*math.sin(t),2,(255,255,255,255))
   line(48,80,112,80);line(55,66,105,66);line(55,94,105,94)
  else:
   for q in [(55,106,98,52),(98,52,110,62),(110,62,67,116),(67,116,55,106),(55,106,51,123),(51,123,67,116),(94,58,105,68)]:line(*q)
  angle=-math.pi/2+phase*math.pi/6;dot(80+60*math.cos(angle),80+60*math.sin(angle),8,(76,224,210,255))
  png(folder/f'{phase:02}.png',pix)
  if phase==0:
   for k in range(3,len(pix),4):pix[k]=round(pix[k]*0.25)
   png(a.outdir/f'icon-{icon}-inactive.png',pix)
 cmd=[a.ffmpeg,'-hide_banner','-loglevel','error','-framerate','5','-i',str(folder/'%02d.png'),'-filter_complex','[0:v]split[s0][s1];[s0]palettegen=reserve_transparent=1[p];[s1][p]paletteuse=alpha_threshold=128','-loop','0',str(a.outdir/f'icon-{icon}-active.gif')]
 subprocess.run(cmd,check=True)
(a.outdir/'provenance.json').write_text(json.dumps({'creator':'original procedural house / globe / pencil line icons','frames':12,'frameDelaySeconds':0.2,'loopSeconds':2.4,'pngSourcePhase':0,'inactiveAlpha':0.25,'gifIsActualMultiframe':True,'noAuthorIcons':True},indent=2),encoding='utf-8')
