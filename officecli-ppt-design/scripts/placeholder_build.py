"""Build original content for genuine title/body placeholder migration tests."""
import argparse,json,subprocess,zipfile
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);a=p.parse_args();out=Path(a.out)
if any(q.exists() for q in [out,*[out.with_suffix(t) for t in ('.batch.json','.validation.json','.build-log.json')]]):raise FileExistsError('Refusing overwrite')
out.parent.mkdir(parents=True,exist_ok=True)
def run(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8')
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
run('create',out);run('set',out,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
content=[('Research signals','Observe customer needs\nSynthesize three interviews'),('Local emphasis','Keep the local statement\nRetain supporting evidence'),('Prototype learning','Test the critical assumption\nRecord one clear outcome'),('Delivery readiness','Prepare a release checklist\nName one accountable owner')]
ops=[]
for i,(title,body) in enumerate(content,1):
 ops.append(dict(command='add',parent='/',type='slide',props={'layout':'blank','background':'F6F8FC'}))
 for name,text,x,y,w,h,size in [('source-title',title,1.7,2,29,2,30),('source-body',body,2,5.7,28,7,22),('local-caption',f'NOVA / STUDY NOTES   {i:02d}',1.7,17.5,29,0.8,12)]:
  ops.append(dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name=name,text=text,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',font='Arial',size=str(size),color='17324E',fill='none',line='none',margin='0cm')))
 ops.append(dict(command='add',parent=f'/slide[{i}]',type='notes',props={'text':'Original content. BV1VTDGYsEmd P23 228-435s / 747-1182s: real title/body placeholder inheritance, plain-text migration, local overrides, native reset and layout change boundaries. No tutorial logo or client copy reused.'}))
f=out.with_suffix('.batch.json');f.write_text(json.dumps(ops),encoding='utf-8');r=run('batch',out,'--input',f,'--json');out.with_suffix('.build-log.json').write_text(r,encoding='utf-8');assert json.loads(r)['success'];run('close',out)
with zipfile.ZipFile(out) as z:assert len([n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')])==4
out.with_suffix('.validation.json').write_text(run('validate',out,'--json'),encoding='utf-8');print(out)
