import argparse,binascii,json,struct,subprocess,zlib
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);W,H=960,540;pixels=bytearray()
for y in range(H):
 for x in range(W):
  rgb=(round(30+190*x/(W-1)),round(35+180*y/(H-1)),220 if (x//8+y//8)%2 else 40)
  if 55<=y<120 and 100<=x<860:rgb=(0,0,0) if x<480 else (255,255,255)
  pixels.extend(rgb)
def chunk(t,d):return struct.pack('>I',len(d))+t+d+struct.pack('>I',binascii.crc32(t+d)&0xffffffff)
raw=b''.join(b'\0'+pixels[y*W*3:(y+1)*W*3] for y in range(H));image=a.outdir/'original-coordinate-checker.png';image.write_bytes(b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',W,H,8,2,0,0,0))+chunk(b'IDAT',zlib.compress(raw))+chunk(b'IEND',b''))
deck=a.outdir/'local-blur.pptx'
def cli(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8')
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
cli('create',deck);cli('open',deck);cli('set',deck,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm');ops=[]
for s in [1,2,3,4]:
 ops.append(dict(command='add',parent='/',type='slide',props=dict(layout='blank',background='image:'+str(image.resolve()) if s>=3 else 'FFFFFF')))
 ops.append(dict(command='add',parent=f'/slide[{s}]',type='picture',props=dict(name='clear-foreground' if s>=3 else 'reference-image',src=str(image.resolve()),x='0cm',y='0cm',width='33.87cm',height='19.05cm')))
 if s>=3:
  for i,(x,y) in enumerate(([(2,5),(19.87,5)] if s==3 else [(7,2),(17,7)]),1):
   ops.append(dict(command='add',parent=f'/slide[{s}]',type='shape',props=dict(name=f'glass-{i}',geometry='roundRect',x=f'{x}cm',y=f'{y}cm',width='12cm',height='9cm',fill='none',line='FFFFFF:1.5')))
f=a.outdir/'batch.json';f.write_text(json.dumps(ops),encoding='utf-8');cli('batch',deck,'--input',f,'--json')
for s in [3,4]:
 for i in [1,2]:
  path=f'//p:sp[p:nvSpPr/p:cNvPr[@name="glass-{i}"]]';cli('raw-set',deck,f'/slide[{s}]','--xpath',path,'--action','setattr','--xml','useBgFill=1','--json');cli('raw-set',deck,f'/slide[{s}]','--xpath',path+'/p:spPr/a:noFill','--action','remove','--json')
cli('close',deck);print(cli('validate',deck,'--json'))
