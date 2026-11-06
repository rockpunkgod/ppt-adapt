"""OfficeCLI-native data table with real grouped/merged headers and selective cell borders."""
import argparse,json,subprocess,zipfile,xml.etree.ElementTree as E
from pathlib import Path
DATA=[('线上',[100,120,130]),('现场',[70,80,90]),('伙伴',[30,40,40])]
def build(cli,out):
    out=Path(out).resolve()
    if out.exists():raise FileExistsError('Refusing overwrite; choose a new presentation')
    out.parent.mkdir(parents=True,exist_ok=True)
    def run(*args):
        r=subprocess.run([cli,*map(str,args)],capture_output=True,text=True,encoding='utf-8')
        if r.returncode:raise RuntimeError(r.stdout+r.stderr)
        return r.stdout
    run('create',out);run('set',out,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
    ops=[]
    def add(parent,kind,**props):ops.append(dict(command='add',parent=parent,type=kind,props={k:str(v) for k,v in props.items()}))
    def set(path,**props):ops.append(dict(command='set',path=path,props={k:str(v) for k,v in props.items()}))
    add('/','slide',layout='blank',background='F5F3EE')
    for name,text,y,h,size,color in [('table-title','一季度完成 700 笔服务。',1.3,2.1,40,'173743'),('table-subtitle','月份、渠道与合计各有位置，分组标题只合并必要单元格。',3.5,1,19,'557078'),('table-source','单位：笔。数据为原创演示；统计口径为已完成服务，不含取消订单。',16.2,1.5,17,'557078')]:
        add('/slide[1]','shape',name=name,text=text,x='1.6cm',y=f'{y}cm',width='30.2cm',height=f'{h}cm',size=size,font='Microsoft YaHei',color=color,fill='none',line='none',margin='0cm')
    rows=[['服务渠道','服务笔数（笔）','','',''],['','1月','2月','3月','季度合计']]
    rows += [[name,*map(str,values),str(sum(values))] for name,values in DATA]
    monthly=[sum(values[i] for name,values in DATA) for i in range(3)]
    rows += [['合计',*map(str,monthly),str(sum(monthly))]]
    add('/slide[1]','table',name='quarter-service-table',data=';'.join(','.join(row) for row in rows),x='1.6cm',y='5.2cm',width='30.2cm',height='9.6cm',colWidths='7cm,5.8cm,5.8cm,5.8cm,5.8cm',style='none',firstRow='false',lastRow='false',bandedRows='false',**{'border.all':'none'})
    table='/slide[1]/table[1]'
    for r in range(1,7):
        for c in range(1,6):
            props={'font':'Microsoft YaHei','size':22,'valign':'center','align':'left' if c==1 else 'right','padding':'0.18cm','fill':'FFFFFF','color':'173743','border.all':'none','bold':'false'}
            if r==1:props.update(fill='176F73',color='FFFFFF',bold='true',align='center')
            if r==2:props.update(fill='E3ECE8',bold='true',align='center')
            if c==5 and r>=3:props.update(color='176F73',bold='true')
            if r==6:props.update(fill='E3ECE8',color='176F73',bold='true')
            if r==2:props['border.bottom']='1.5pt solid 176F73'
            if r in (3,4):props['border.bottom']='0.5pt solid D8E1DC'
            if r==5:props['border.bottom']='1.5pt solid 176F73'
            if r==6:props.update({'border.top':'1.5pt solid 176F73','border.bottom':'1.5pt solid 176F73'})
            set(f'{table}/tr[{r}]/tc[{c}]',**props)
    set(table+'/tr[1]/tc[1]',rowspan=2,align='left',**{'border.bottom':'1.5pt solid 176F73'})
    set(table+'/tr[1]/tc[2]',colspan=4)
    add('/slide[1]','notes',text='Source mechanism: BV1VTDGYsEmd P14 867-1044s merge/split, 1074-1308s numeric consistency, 2211-2754s selective borders. Original fictional data. A real table, not textbox/rectangle simulation.')
    batch=out.with_suffix('.batch.json');batch.write_text(json.dumps(ops,ensure_ascii=False,indent=2),encoding='utf-8')
    result=run('batch',out,'--input',batch,'--json');out.with_suffix('.build-log.json').write_text(result,encoding='utf-8')
    if json.loads(result).get('success') is False:raise RuntimeError(result)
    run('close',out)
    with zipfile.ZipFile(out) as z:
        root=E.fromstring(z.read('ppt/slides/slide1.xml'));assert len(root.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}tbl'))==1
    out.with_suffix('.validation.json').write_text(run('validate',out,'--json'),encoding='utf-8');print(out)
if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);a=p.parse_args();build(a.officecli,a.out)
