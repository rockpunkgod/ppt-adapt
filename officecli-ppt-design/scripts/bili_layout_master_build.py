"""OfficeCLI foundations for a three-slide real master/layout inheritance specimen."""
import argparse,json,subprocess,zipfile,xml.etree.ElementTree as E
from pathlib import Path

def build(cli,out):
    out=Path(out).resolve()
    if out.exists():raise FileExistsError('Refusing overwrite; choose a new PPT')
    out.parent.mkdir(parents=True,exist_ok=True)
    def run(*args):
        r=subprocess.run([cli,*map(str,args)],capture_output=True,text=True,encoding='utf-8')
        if r.returncode:raise RuntimeError(r.stdout+r.stderr)
        return r.stdout
    run('create',out);run('set',out,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
    commands=[]
    def add(parent,kind,**props):commands.append(dict(command='add',parent=parent,type=kind,props={k:str(v) for k,v in props.items()}))
    content=[('研究问题','把问题先讲清楚。\n观察真实使用场景，记录用户遇到的阻碍。','LOCAL-A1'),('验证结果','把证据连成结论。\n用原型测试验证假设，保留反例和反馈。','LOCAL-A2'),('实施计划','把结论转为行动。\n明确责任、顺序和检查点，再推进下一阶段。','LOCAL-B1')]
    for i,(title,body,key) in enumerate(content,1):
        add('/','slide',layout='blank',background='F5F3EE')
        for name,text,y,h,size,color in [('local-title',title,5.8,2.2,40,'173743'),('local-body',body,9,4,25,'38565F'),('local-key',key,17.2,.9,13,'667F86')]:
            add(f'/slide[{i}]','shape',name=name,text=text,x='1.6cm',y=f'{y}cm',width='30.2cm',height=f'{h}cm',font='Microsoft YaHei',size=size,color=color,fill='none',line='none',margin='0cm')
        add(f'/slide[{i}]','notes',text='Source mechanism: BV1VTDGYsEmd P23 453-618s. Local content stays on this slide; shared mark belongs to master and branch strip to CustomLayout.')
    path=out.with_suffix('.batch.json');path.write_text(json.dumps(commands,ensure_ascii=False,indent=2),encoding='utf-8')
    result=run('batch',out,'--input',path,'--json');out.with_suffix('.build-log.json').write_text(result,encoding='utf-8')
    if json.loads(result).get('success') is False:raise RuntimeError(result)
    run('close',out)
    with zipfile.ZipFile(out) as z:
        root=E.fromstring(z.read('ppt/presentation.xml'));assert len(root.findall('{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst/*'))==3
    out.with_suffix('.validation.json').write_text(run('validate',out,'--json'),encoding='utf-8');print(out)
if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);a=p.parse_args();build(a.officecli,a.out)
