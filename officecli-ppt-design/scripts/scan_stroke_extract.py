import argparse,json,subprocess,struct,zlib,hashlib,shutil
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--input',type=Path,required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--paper',default='248,246,239');p.add_argument('--ink',default='24,30,26');p.add_argument('--color',default='FFFFFF');p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
raw=a.input.read_bytes()
if raw[:8]!=b'\x89PNG\r\n\x1a\n':raise ValueError('This bounded workflow accepts explicitly supplied PNG only')
if raw[24]!=8:raise ValueError('Only 8-bit encoded PNG is validated; high-bit-depth input unsupported')
W,H=struct.unpack('>II',raw[16:24])
if not(1<=W<=4096 and 1<=H<=4096):raise ValueError('Image bounds unsupported')
paper=tuple(map(int,a.paper.split(',')));ink=tuple(map(int,a.ink.split(',')));color=tuple(bytes.fromhex(a.color));axis=[p-i for p,i in zip(paper,ink)];den=sum(v*v for v in axis)
if len(paper)!=3 or len(ink)!=3 or len(color)!=3 or any(not 0<=v<=255 for v in paper+ink) or den<1000:raise ValueError('Three-channel well-separated calibrated paper/ink colors required')
pix=subprocess.run([a.ffmpeg,'-v','error','-i',str(a.input),'-frames:v','1','-f','rawvideo','-pix_fmt','rgba','-'],capture_output=True,check=True).stdout
if len(pix)!=W*H*4 or any(pix[j]!=255 for j in range(3,len(pix),4)):raise ValueError('Input must be opaque single clean scan fixture; existing alpha unsupported')
a.outdir.mkdir(parents=True);shutil.copyfile(a.input,a.outdir/'original.png');out=bytearray();alpha=bytearray()
for j in range(0,len(pix),4):
 v=max(0,min(255,round(255*sum((paper[c]-pix[j+c])*axis[c] for c in range(3))/den)));alpha.append(v);out.extend(color+(v,))
def chunk(t,b):return struct.pack('>I',len(b))+t+b+struct.pack('>I',zlib.crc32(t+b)&0xffffffff)
rows=b''.join(b'\0'+out[y*W*4:(y+1)*W*4] for y in range(H));png=b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',W,H,8,6,0,0,0))+chunk(b'IDAT',zlib.compress(rows,9))+chunk(b'IEND',b'');(a.outdir/'extracted.png').write_bytes(png)
(a.outdir/'parameters.json').write_text(json.dumps({'inputSHA256':hashlib.sha256(raw).hexdigest(),'outputSHA256':hashlib.sha256(png).hexdigest(),'width':W,'height':H,'paper':paper,'ink':ink,'outputRGB':color,'formula':'alpha=clamp(round(255*dot(paper-pixel,paper-ink)/dot(paper-ink,paper-ink)))','assumptions':'Known uniform paper, single neutral/dark ink; encoded RGB 8bit arithmetic. No unknown shadows, paper texture, colored photos, auto calibration or cleanup threshold.'},indent=2),encoding='utf-8')
