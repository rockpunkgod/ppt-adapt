"""Prepare and independently verify workbook materialization for literal category charts."""
import argparse,copy,hashlib,io,json,math,posixpath,re,zipfile,xml.etree.ElementTree as E
from pathlib import Path
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','c':'http://schemas.openxmlformats.org/drawingml/2006/chart','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships','s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
def rels(z,part):
    path=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
    if path not in z.namelist():return {}
    return {r.get('Id'):posixpath.normpath(posixpath.join(posixpath.dirname(part),r.get('Target'))).lstrip('/') for r in E.fromstring(z.read(path)) if r.get('TargetMode')!='External'}
def points(node):
    cache=next((node.find('c:'+name,N) for name in ('strLit','numLit','strRef/c:strCache','numRef/c:numCache') if node.find('c:'+name,N) is not None),None)
    assert cache is not None,'Chart cache not supported'
    pts=cache.findall('c:pt',N);assert [int(p.get('idx')) for p in pts]==list(range(len(pts))),'Sparse chart caches require a separate policy'
    return [p.find('c:v',N).text or '' for p in pts]
def charts(file):
    result=[]
    with zipfile.ZipFile(file) as z:
        slide_names=sorted((n for n in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml',n)),key=lambda x:int(re.search(r'(\d+)\.xml',x).group(1)))
        for slide in slide_names:
            root=E.fromstring(z.read(slide));sr=rels(z,slide)
            for f in root.findall('.//p:graphicFrame',N):
                chart=f.find('.//c:chart',N)
                if chart is None:continue
                name=f.find('p:nvGraphicFramePr/p:cNvPr',N).get('name');part=sr[chart.get('{'+N['r']+'}id')];cr=E.fromstring(z.read(part));series=[]
                for ser in cr.findall('.//c:ser',N):
                    tx=ser.find('c:tx',N);direct=tx.find('c:v',N)
                    label=direct.text if direct is not None else points(tx)[0]
                    cats=points(ser.find('c:cat',N));vals=[float(v) for v in points(ser.find('c:val',N))]
                    assert len(cats)==len(vals) and len(vals)>0 and all(math.isfinite(v) for v in vals)
                    series.append({'name':label,'categories':cats,'values':vals})
                assert series and all(s['categories']==series[0]['categories'] for s in series),'Only shared category charts are supported'
                plot=cr.find('c:chart/c:plotArea',N);kinds=[c.tag.rsplit('}',1)[-1] for c in plot if c.tag.rsplit('}',1)[-1].endswith('Chart')]
                result.append({'slide':int(re.search(r'(\d+)\.xml',slide).group(1)),'shape':name,'part':part,'kinds':kinds,'series':series,'hasWorkbook':cr.find('c:externalData',N) is not None})
    return result
def col(n):
    result=''
    while n:n,k=divmod(n-1,26);result=chr(65+k)+result
    return result
def workbook_cells(data):
    with zipfile.ZipFile(io.BytesIO(data)) as w:
        strings=[]
        if 'xl/sharedStrings.xml' in w.namelist():strings=[''.join(t.text or '' for t in si.findall('.//s:t',N)) for si in E.fromstring(w.read('xl/sharedStrings.xml'))]
        wb=E.fromstring(w.read('xl/workbook.xml'));sheet=wb.find('s:sheets/s:sheet',N);wp=rels(w,'xl/workbook.xml')[sheet.get('{'+N['r']+'}id')]
        cells={}
        for cell in E.fromstring(w.read(wp)).findall('.//s:c',N):
            v=cell.find('s:v',N);typ=cell.get('t')
            value=''.join(t.text or '' for t in cell.findall('.//s:t',N)) if typ=='inlineStr' else (v.text if v is not None else '')
            if typ=='s':value=strings[int(value)]
            cells[cell.get('r')]=value
        return sheet.get('name'),cells
def expected(source,edits):
    model=copy.deepcopy(charts(source))
    for edit in edits:
        chart=next(c for c in model if c['slide']==edit['slide'] and c['shape']==edit['shape'])
        chart['series'][edit['series']-1]['values'][edit['point']-1]=float(edit['value'])
    return model
def local_texts(file):
    result={}
    with zipfile.ZipFile(file) as z:
        for name in z.namelist():
            if not re.fullmatch(r'ppt/slides/slide\d+\.xml',name):continue
            root=E.fromstring(z.read(name));result[name]={}
            for shape in root.findall('.//p:sp',N):
                nv=shape.find('p:nvSpPr/p:cNvPr',N)
                if nv is not None:result[name][nv.get('name')]=''.join(t.text or '' for t in shape.findall('.//a:t',N))
    return result
def verify(source,output,edits):
    original=expected(source,edits);actual=charts(output);assert len(original)==len(actual)>0
    assert local_texts(source)==local_texts(output),'Non-chart slide text or annotations changed'
    reports=[]
    with zipfile.ZipFile(output) as z:
        for orig in original:
            cur=next(c for c in actual if c['slide']==orig['slide'] and c['shape']==orig['shape'])
            assert cur['series']==orig['series'],('Chart cache changed incorrectly',orig['shape'])
            assert cur['kinds']==orig['kinds'],'Chart type changed'
            part=cur['part'];root=E.fromstring(z.read(part));external=root.find('c:externalData',N);assert external is not None
            embedded=rels(z,part)[external.get('{'+N['r']+'}id')];sheet,cells=workbook_cells(z.read(embedded))
            for j,series in enumerate(cur['series'],2):
                assert cells[col(j)+'1']==series['name']
                for row,(cat,value) in enumerate(zip(series['categories'],series['values']),2):
                    assert cells['A'+str(row)]==cat
                    assert float(cells[col(j)+str(row)])==value
                ser=root.findall('.//c:ser',N)[j-2]
                vformula=ser.find('c:val/c:numRef/c:f',N);xformula=ser.find('c:cat/c:strRef/c:f',N)
                assert vformula is not None and xformula is not None,'Series is not bound to embedded cells'
                assert vformula.text.endswith(f'${col(j)}$2:${col(j)}${len(series["values"])+1}')
                assert xformula.text.endswith(f'$A$2:$A${len(series["values"])+1}')
            reports.append({'slide':cur['slide'],'shape':cur['shape'],'embeddedWorkbook':embedded,'sheet':sheet,'series':cur['series']})
    return {'passed':True,'source':str(source),'output':str(output),'edits':edits,'localTextPreserved':True,'charts':reports,'limits':'Embedded values, formulas and caches checked; separately compare native PNG and COM format audit.'}
if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);sub=p.add_subparsers(dest='action',required=True)
    prep=sub.add_parser('plan');prep.add_argument('input');prep.add_argument('--out',required=True)
    check=sub.add_parser('verify');check.add_argument('source');check.add_argument('output');check.add_argument('--report',required=True);check.add_argument('--edits')
    a=p.parse_args();dest=Path(a.out if a.action=='plan' else a.report)
    if dest.exists():raise FileExistsError('Refusing overwrite')
    if a.action=='plan':
        found=charts(a.input)
        assert len(found)==1,'Workbook materialization is validated for one chart per document only; multi-chart lifecycle needs separate validation'
        assert not any(c['hasWorkbook'] for c in found),'Only literal charts without existing workbooks are accepted'
        result={'source':str(Path(a.input).resolve()),'sourceSha256':hashlib.sha256(Path(a.input).read_bytes()).hexdigest(),'charts':found}
    else:result=verify(a.source,a.output,json.loads(Path(a.edits).read_text(encoding='utf-8-sig')) if a.edits else [])
    dest.parent.mkdir(parents=True,exist_ok=True);dest.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(dest)
