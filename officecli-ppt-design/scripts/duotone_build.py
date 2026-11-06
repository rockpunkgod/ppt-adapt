import argparse,binascii,json,math,struct,subprocess,zlib
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);W,H=512,256;pixels=bytearray()
for y in range(H):
 for x in range(W):
  if y<64:v=round((x//32)*255/15)
  elif y<128:v=round(x*255/(W-1))
  else:
   v=round(30+180*x/(W-1));hill=190+24*math.sin(x/55)
   if y>hill:v=round(25+90*x/(W-1))
   if (x-370)**2+(y-168)**2<30**2:v=235
  pixels.extend([v,v,v])
def chunk(t,d):return struct.pack('>I',len(d))+t+d+struct.pack('>I',binascii.crc32(t+d)&0xffffffff)
raw=b''.join(b'\0'+pixels[y*W*3:(y+1)*W*3] for y in range(H));image=a.outdir/'original-gray.png';image.write_bytes(b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',W,H,8,2,0,0,0))+chunk(b'IDAT',zlib.compress(raw))+chunk(b'IEND',b''))
deck=a.outdir/'duotone.pptx'
def cli(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8')
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
cli('create',deck);cli('open',deck);cli('set',deck,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm');ops=[]
for i,title in enumerate(['Original grayscale | endpoints, ramp and original scene','Native duotone | 201044 to F2D787','Native duotone | F2D787 to 201044'],1):
 ops.append(dict(command='add',parent='/',type='slide',props=dict(layout='blank',background='F5F2EA')))
 ops.append(dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name='title',text=title,x='1.5cm',y='0.7cm',width='30.87cm',height='1.5cm',font='Arial',size=24,color='201044',fill='none',line='none')))
 ops.append(dict(command='add',parent=f'/slide[{i}]',type='picture',props=dict(name='tone-original',src=str(image.resolve()),x='1.5cm',y='3.1cm',width='30.87cm',height='15.435cm')))
f=a.outdir/'batch.json';f.write_text(json.dumps(ops),encoding='utf-8');cli('batch',deck,'--input',f,'--json')
for s,colors in [(2,['201044','F2D787']),(3,['F2D787','201044'])]:
 xml='<a:duotone xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'+''.join(f'<a:srgbClr val="{c}"/>' for c in colors)+'</a:duotone>'
 cli('raw-set',deck,f'/slide[{s}]','--xpath','//p:pic[p:nvPicPr/p:cNvPr[@name="tone-original"]]/p:blipFill/a:blip','--action','append','--xml',xml,'--json')
cli('close',deck);print(cli('validate',deck,'--json'))
