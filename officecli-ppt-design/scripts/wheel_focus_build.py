"""OfficeCLI foundation for Wheel / focus animation foundation. Refuses overwrite."""
import sys,argparse,json
from pathlib import Path
import subprocess
W,H=33.87,19.05
FONT='Microsoft YaHei'
class Deck:
    def __init__(self, exe: str, path: Path):
        self.exe,self.path=exe,path
        self.log=[]
        if path.exists():raise FileExistsError(f'Refusing to overwrite existing presentation: {path}. Choose a new --outdir.')
        self.call('create',str(path),'--locale','zh-CN')
        self.call('set',str(path),'/', '--prop',f'slideWidth={W}cm','--prop',f'slideHeight={H}cm')
        self.n=0
    def call(self,*args):
        proc=subprocess.run([self.exe,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace')
        self.log.append({'args':list(map(str,args)),'code':proc.returncode,'stdout':proc.stdout,'stderr':proc.stderr})
        if proc.returncode or 'UNSUPPORTED' in proc.stdout+proc.stderr:
            raise RuntimeError(str(args)+'\n'+proc.stdout+'\n'+proc.stderr)
        return proc.stdout
    def batch(self,ops):
        for j in range(0,len(ops),12):
            inp=self.path.with_suffix('.batch.json')
            inp.write_text(json.dumps(ops[j:j+12],ensure_ascii=False),encoding='utf-8')
            out=self.call('batch',str(self.path),'--input',str(inp),'--json')
            parsed=json.loads(out)
            if parsed.get('success') is False or 'atomicRolledBack": true' in out:
                raise RuntimeError(out)
    def add(self,typ,parent,**props):
        return {'command':'add','parent':parent,'type':typ,'props':{k:str(v).lower() if isinstance(v,bool) else str(v) for k,v in props.items()}}
    def slide(self,bg='151027',**extra):
        self.n+=1;self.batch([self.add('slide','/',layout='blank',background=bg,**extra)])
        return f'/slide[{self.n}]'
    def shape(self,s,name,x,y,w,h,text='',size=20,color='FFFFFF',fill='none',preset='rect',**extra):
        props=dict(name=name,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',text=text,font=FONT,size=size,color=color,fill=fill,preset=preset,line='none',margin='0cm')
        props.update(extra)
        return self.add('shape',s,**props)
    def title(self,s,t,sub=''):
        ops=[self.shape(s,'title',1.5,1.1,30,2,t,size=36,bold=True)]
        if sub:ops.append(self.shape(s,'subtitle',1.5,3.1,29.5,1.2,sub,size=18,color='C7C0D5'))
        return ops
    def note(self,s,t):return self.add('notes',s,text=t)
    def motion_flags(self,s,reverse=False,repeat=False):
        # 1.0.159/160 help advertises these but Add warns unsupported and Set
        # reports success without persisting them. Explicit OOXML is verified.
        if reverse:self.call('raw-set',str(self.path),s,'--xpath',"//p:cTn[@presetClass='path']",'--action','setattr','--xml','autoRev=1')
        if repeat:self.call('raw-set',str(self.path),s,'--xpath',"//p:cTn[@presetClass='path']",'--action','setattr','--xml','repeatCount=indefinite')
    def finish(self):
        self.call('close',str(self.path))
        validation=self.call('validate',str(self.path))
        issues=self.call('view',str(self.path),'issues','--json')
        self.path.with_suffix('.validation.txt').write_text(validation+'\n'+issues,encoding='utf-8')
        self.call('view',str(self.path),'html','-o',str(self.path.with_suffix('.html')))
        self.path.with_suffix('.build-log.json').write_text(json.dumps(self.log,ensure_ascii=False,indent=2),encoding='utf-8')


import math
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--outdir',required=True);a=p.parse_args();o=Path(a.outdir)
if o.exists():p.error('Refusing existing output directory')
o.mkdir(parents=True);d=Deck(a.officecli,o/'wheel-focus-base.pptx')
s=d.slide('F2F0E9');titles=d.title(s,'FOCUS / PRIORITY','The chosen idea grows; the surrounding context recedes.')
for t in titles:t['props']['color']='193C40'
d.batch(titles)
ops=[]
for i,(x,text) in enumerate([(3.5,'CONTEXT'),(14.1,'PRIORITY'),(24.7,'DETAIL')],1):
 ops.append(d.shape(s,f'focus-{i}',x,7,5.6,5.6,text,size=22,fill='21645E',preset='ellipse',align='center',valign='middle'))
for op in ops:op['props']['color']='FFFFFF'
d.batch(ops)
d.call('set',d.path,s,'--prop','advanceTime=5000','--prop','advanceClick=false')
s=d.slide('092B3E');d.batch(d.title(s,'WHEEL / REVEAL','Reveal each point as the ring reaches its position.'))
ops=[d.shape(s,'wheel-ring',11.2,5,11.4,11.4,fill='none',preset='ellipse',line='66B9AE:3'),d.shape(s,'wheel-core',14.8,8.6,4.2,4.2,'IDEA',size=23,fill='21645E',preset='ellipse',align='center',valign='middle')]
for i in range(6):
 angle=math.radians(-90+i*60);cx=16.9+5.7*math.cos(angle);cy=10.7+5.7*math.sin(angle)
 ops.append(d.shape(s,f'wheel-node-{i+1}',cx-1.1,cy-1.1,2.2,2.2,str(i+1).zfill(2),size=22,fill='66B9AE',preset='ellipse',align='center',valign='middle',color='092B3E'))
d.batch(ops);d.call('set',d.path,s,'--prop','advanceTime=9500','--prop','advanceClick=false');d.finish();print(d.path)
