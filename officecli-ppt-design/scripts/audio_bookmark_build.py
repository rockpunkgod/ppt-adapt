import argparse,json,subprocess,wave
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);deck=a.outdir/'bookmark-before.pptx';wav=a.outdir/'silence-2500ms.wav'
with wave.open(str(wav),'wb') as w:w.setnchannels(1);w.setsampwidth(2);w.setframerate(16000);w.writeframes(b'\0\0'*40000)
def call(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace')
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 if '--json' in args:
  result=json.loads(r.stdout)
  if result.get('success') is False:raise RuntimeError(r.stdout)
 return r.stdout
call('create',deck);call('open',deck);call('set',deck,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
ops=[dict(command='add',parent='/',type='slide',props=dict(layout='blank',background='14283F'))]
def shape(name,x,y,w,h,fill,text='',size=22,typ='rect'):
 ops.append(dict(command='add',parent='/slide[1]',type='shape',props=dict(name=name,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',geometry=typ,fill=fill,line='none',text=text,font='Arial',size=size,color='FFFFFF',margin='0cm',valign='center',align='center',autoFit='none')))
shape('title',2,1.1,30,1.4,'none','One clock. One repeatable sequence.',30)
shape('subtitle',2,3.1,30,1,'none','Silent audio bookmark restarts every 2.5 seconds',16)
shape('photo',14,11.3,5.9,4.5,'FFFFFF','',22)
shape('photo-scene',14.3,11.65,5.3,3.1,'6AAE9D','FIELD NOTES',19)
shape('camera-body',11,7.5,12,5.2,'364F6B')
shape('camera-header',11,6.4,12,1.4,'A3B7CD')
shape('lens',15.3,8.2,3.5,3.5,'14283F','',22,'ellipse')
shape('flash',12,6.65,2.2,.8,'FFFFFF')
shape('note',2,17.1,30,1,'none','2.0-second action + 0.5-second hold. Original vector artwork.',14)
ops.append(dict(command='add',parent='/slide[1]',type='audio',props=dict(src=str(wav),x='-2cm',y='-2cm',width='1cm',height='1cm',autoPlay=True,loop=True,volume=0)))
f=a.outdir/'base.batch.json';f.write_text(json.dumps(ops),encoding='utf-8');call('batch',deck,'--input',f,'--json');call('close',deck);result=call('validate',deck,'--json');(a.outdir/'validation.json').write_text(result,encoding='utf-8');print(result)
