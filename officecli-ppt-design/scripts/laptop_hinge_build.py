"""Original editable laptop planes, equal-height hinge helpers, optional real media."""
import argparse,json,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--poster',type=Path,required=True);p.add_argument('--video',type=Path,required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():p.error('Fresh output directory required')
a.outdir.mkdir(parents=True);deck=a.outdir/'laptop-base.pptx';log=[];counter=100000
def cli(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace');log.append({'args':list(map(str,args)),'code':r.returncode,'stdout':r.stdout})
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
def batch(ops,stem):
 f=a.outdir/(stem+'.json');f.write_text(json.dumps(ops,ensure_ascii=False),encoding='utf-8');cli('batch',deck,'--input',f,'--json')
def shape(ops,s,name,x,y,w,h,text='',**extra):
 global counter
 ident=counter;counter+=1
 props={'id':str(ident),'name':name,'x':f'{x}cm','y':f'{y}cm','width':f'{w}cm','height':f'{h}cm','text':text,'fill':'none','line':'none','font':'Arial','size':'18','color':'F3F1EB','margin':'0cm','valign':'middle'};props.update({k:str(v).lower() if isinstance(v,bool) else str(v) for k,v in extra.items()});ops.append({'command':'add','parent':f'/slide[{s}]','type':'shape','props':props});return f'/slide[{s}]/shape[@id={ident}]'
def group(s,name,paths):
 cli('add',deck,f'/slide[{s}]','--type','group','--prop','name='+name,'--prop','shapes='+','.join(paths))
 return json.loads(cli('get',deck,f'/slide[{s}]/group[@name={name}]','--json'))['data']['results'][0]['path']
cli('create',deck,'--locale','en-US');cli('set',deck,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
for s in [1,2]:
 cli('add',deck,'/','--type','slide','--prop','layout=blank','--prop','background=101A2A')
 ops=[];shape(ops,s,'background',0,0,33.87,19.05,gradient='102438-245B65-90')
 shape(ops,s,'headline',1.3,0.2,31.2,0.8,'TWO PLANES. ONE HINGE.',size=21,bold=True,align='center')
 shape(ops,s,'footer',1.3,17.4,31.2,1.0,'ORIGINAL EDITABLE LAPTOP / LIVE SCREEN',size=17,align='center',color='A9C8CA')
 keys=[];keys.append(shape(ops,s,'keyboard-shell',7.835,11.9,18.2,10.6,fill='BDC7CF',geometry='roundRect',depth='3pt',bevelBottom='circle-5-5',material='matte'))
 for row in range(4):
  for col in range(12):keys.append(shape(ops,s,f'key-{row}-{col}',8.55+col*1.36,12.65+row*1.2,1.18,0.9,fill='182A38',geometry='roundRect'))
 keys.append(shape(ops,s,'space-bar',11.25,17.5,10.2,0.8,fill='182A38',geometry='roundRect'))
 keys.append(shape(ops,s,'trackpad',13.1,18.8,7.1,2.6,fill='AAB8C5',geometry='roundRect',line='8094A3:0.5'))
 key_helper=shape(ops,s,'keyboard-axis-helper',7.835,1.3,18.2,10.6)
 screen=[];screen.append(shape(ops,s,'screen-bezel',7.835,1.3,18.2,10.6,fill='10151D',geometry='roundRect',depth='3pt',bevelBottom='circle-5-5',material='matte'))
 screen.append(shape(ops,s,'screen-canvas',8.435,1.8,17,9.5625,image=str(a.poster.resolve())))
 screen.append(shape(ops,s,'camera-dot',16.85,1.48,0.17,0.17,fill='647C8A',geometry='ellipse'))
 screen_helper=shape(ops,s,'screen-axis-helper',7.835,11.9,18.2,10.6)
 batch(ops,'shapes-'+str(s))
 keyface=group(s,'keyboard-face',keys);group(s,'!!keyboard-hinge',[keyface,key_helper])
 screenface=group(s,'screen-face',screen);group(s,'!!screen-hinge',[screenface,screen_helper])
 if s==2 and a.video:
  cli('add',deck,'/slide[2]','--type','video','--prop','name=screen-video','--prop','src='+str(a.video.resolve()),'--prop','poster='+str(a.poster.resolve()),'--prop','x=8.435cm','--prop','y=1.8cm','--prop','width=17cm','--prop','height=9.5625cm','--prop','autoPlay=true','--prop','loop=true','--prop','volume=0')
 cli('set',deck,f'/slide[{s}]','--prop','transition=morph')
cli('close',deck);(a.outdir/'build-log.json').write_text(json.dumps(log,indent=2),encoding='utf-8');print(deck)
