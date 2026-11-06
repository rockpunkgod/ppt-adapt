"""Create original SmartArt study slide foundations using OfficeCLI only."""
import argparse,json,subprocess,zipfile,xml.etree.ElementTree as E
from pathlib import Path

def build(cli,out):
    out=Path(out).resolve()
    if out.exists():raise FileExistsError('Choose a new PPT; refusing overwrite')
    out.parent.mkdir(parents=True,exist_ok=True)
    def run(*args):
        r=subprocess.run([str(cli),*map(str,args)],capture_output=True,text=True,encoding='utf-8')
        if r.returncode:raise RuntimeError(r.stdout+r.stderr)
        return r.stdout
    run('create',out);run('set',out,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
    cmds=[]
    def add(parent,kind,**props):cmds.append(dict(command='add',parent=parent,type=kind,props={k:str(v) for k,v in props.items()}))
    for i,title in enumerate(['同一份内容，保留真正的父子结构。','更换布局后，内容关系保持一致。','图片与标题，逐项确认配对。'],1):
        add('/','slide',layout='blank',background='F5F3EE')
        def text(name,content,y,size,color='173743'):
            add(f'/slide[{i}]','shape',name=name,text=content,x='1.4cm',y=f'{y}cm',width='31cm',height='1.8cm',size=size,color=color,font='Microsoft YaHei',fill='none',line='none',margin='0cm')
        text('study-title',title,1,34)
        subtitle='产品研究分为需求洞察与原型测试；反馈归纳属于原型测试。' if i<3 else '每个节点同时关联图片与标题，避免排版时错配。'
        text('study-subtitle',subtitle,3,19,'557078')
        text('study-footnote','原生 SmartArt · 内容与关系可以继续编辑',17.2,14,'557078')
        add(f'/slide[{i}]','notes',text='Source mechanism: BV1VTDGYsEmd P13. Original example, not course artwork. OfficeCLI foundation; COM adds native SmartArt and tests node edits/layout.')
    batch=out.with_suffix('.batch.json');batch.write_text(json.dumps(cmds,ensure_ascii=False,indent=2),encoding='utf-8')
    result=run('batch',out,'--input',str(batch),'--json');out.with_suffix('.build-log.json').write_text(result,encoding='utf-8')
    parsed=json.loads(result)
    if parsed.get('success') is False or parsed.get('ok') is False:raise RuntimeError(result)
    run('close',out)
    with zipfile.ZipFile(out) as z:
        root=E.fromstring(z.read('ppt/presentation.xml'))
        assert len(root.findall('{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst/*'))==3,'Disk presentation must contain three slides before COM opens it'
    out.with_suffix('.validation.json').write_text(run('validate',out,'--json'),encoding='utf-8')
    print(out)
if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);a=p.parse_args();build(a.officecli,a.out)
