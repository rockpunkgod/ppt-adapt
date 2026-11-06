"""OfficeCLI foundation for named Pin open/close trigger. Refuses overwrite."""
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

p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--outdir',required=True);a=p.parse_args();o=Path(a.outdir)
if o.exists():p.error('Refusing existing output directory')
o.mkdir(parents=True);d=Deck(a.officecli,o/'pin-base.pptx');s=d.slide('E9E7DF',advanceClick=False)
ops=d.title(s,'PIN / REVEAL','Click the same pin to open, then close the city label.')
for op in ops:op['props']['color']='193C40'
ops += [d.shape(s,'trigger-pin',4.5,7,6,6,'PIN',size=40,fill='1C6561',preset='ellipse',align='center',valign='middle'),d.shape(s,'trigger-label',13,7,16,6,'BERLIN\n52.52° N / 13.40° E',size=32,fill='B9D3CD',color='193C40',align='center',valign='middle'),d.shape(s,'footer',2,16.6,29,1,'One slide · one named trigger · native interactive sequence',size=18,color='193C40')]
d.batch(ops);d.finish();print(d.path)
