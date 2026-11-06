"""OfficeCLI page + original InkML + our native-UI-generated Replay timing."""
import argparse,json,subprocess,sys,wave
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);deck=a.outdir/'cli-base.pptx';here=Path(__file__).resolve().parent
def call(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace',check=True)
 if '--json' in args and json.loads(r.stdout).get('success') is False:raise RuntimeError(r.stdout)
 return r.stdout
call('create',deck);call('open',deck);call('set',deck,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
ops=[{'command':'add','parent':'/','type':'slide','props':{'layout':'blank','background':'14283F'}}]
for name,text,y,size in [('title','Draw once. Replay, then hold.',1,30),('note','Native ink replay: 5 seconds + 2 seconds hold',17,16)]:
 ops.append({'command':'add','parent':'/slide[1]','type':'shape','props':{'name':name,'text':text,'x':'2cm','y':str(y)+'cm','width':'30cm','height':'1.5cm','fill':'none','line':'none','font':'Arial','size':size,'color':'FFFFFF','align':'center'}})
for i,(x,y) in enumerate([(6,12),(13,7),(20,12),(27,7)]):
 ops.append({'command':'add','parent':'/slide[1]','type':'shape','props':{'name':'node-'+str(i),'text':str(i+1),'x':str(x)+'cm','y':str(y)+'cm','width':'1.5cm','height':'1.5cm','geometry':'ellipse','fill':'6AAE9D','line':'none','font':'Arial','size':24,'color':'FFFFFF','align':'center','valign':'center','margin':'0cm'}})
wav=a.outdir/'silence-7000ms.wav'
with wave.open(str(wav),'wb') as w:w.setnchannels(1);w.setsampwidth(2);w.setframerate(16000);w.writeframes(b'\0\0'*112000)
ops.append({'command':'add','parent':'/slide[1]','type':'audio','props':{'src':str(wav),'x':'-2cm','y':'-2cm','width':'1cm','height':'1cm','autoPlay':True,'loop':True,'volume':0}})
batch=a.outdir/'batch.json';batch.write_text(json.dumps(ops),encoding='utf-8');call('batch',deck,'--input',batch,'--json');call('close',deck)
ink=a.outdir/'ink-only.pptx';final=a.outdir/'ink-before.pptx'
subprocess.run([sys.executable,str(here/'ink_pause_contentpart.py'),'--base',str(deck),'--output',str(ink)],check=True)
subprocess.run([sys.executable,str(here/'ink_pause_replay_seed.py'),'--base',str(ink),'--timing',str(here.parent/'assets'/'ink-replay-timing.xml'),'--output',str(final)],check=True)
print(call('validate',final,'--json'))
