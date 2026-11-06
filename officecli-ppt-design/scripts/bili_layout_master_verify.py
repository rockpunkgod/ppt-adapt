"""Independently verify true slide->layout->master relationships and isolated updates."""
import argparse,json,posixpath,zipfile,xml.etree.ElementTree as E
from pathlib import Path
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def relation(z,part,kind):
    path=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
    matches=[r for r in E.fromstring(z.read(path)) if r.get('Type').endswith('/'+kind)]
    assert len(matches)==1,(part,kind,len(matches))
    return posixpath.normpath(posixpath.join(posixpath.dirname(part),matches[0].get('Target')))
def shape_map(root):
    return {s.find('p:nvSpPr/p:cNvPr',N).get('name'):s for s in root.findall('p:cSld/p:spTree/p:sp',N)}
def text(shape):return '\n'.join(''.join(t.text or '' for t in p.findall('.//a:t',N)) for p in shape.findall('p:txBody/a:p',N))
def fill(shape):return shape.find('p:spPr/a:solidFill/a:srgbClr',N).get('val')
def inspect(path,phase):
    with zipfile.ZipFile(path) as z:
        layouts=[];masters=[];local=[]
        for i in (1,2,3):
            slide=f'ppt/slides/slide{i}.xml';root=E.fromstring(z.read(slide));shapes=shape_map(root)
            assert set(shapes)=={'local-title','local-body','local-key'},(phase,i,'Inherited graphics copied onto slide')
            assert root.get('showMasterSp','1')!='0'
            local.append({name:text(shape) for name,shape in shapes.items()})
            layout=relation(z,slide,'slideLayout');layouts.append(layout)
            lr=E.fromstring(z.read(layout));assert lr.get('showMasterSp','1')!='0'
            masters.append(relation(z,layout,'slideMaster'))
        assert layouts[0]==layouts[1] and layouts[0]!=layouts[2],'A/A/B must be actual separate layout bindings'
        assert len(set(masters))==1,'All layouts must share the same master'
        master=E.fromstring(z.read(masters[0]));ms=shape_map(master)
        assert 'master-shared-brand' in ms
        masterText='NOVA / STUDIO' if phase=='before' else 'NOVA / SYSTEM 2026'
        assert text(ms['master-shared-brand'])==masterText
        assert fill(ms['master-shared-brand'])==('3867A8' if phase=='before' else '1B7775')
        a=shape_map(E.fromstring(z.read(layouts[0])));b=shape_map(E.fromstring(z.read(layouts[2])))
        assert set(a)=={'layout-A-strip'} and set(b)=={'layout-B-strip'}
        assert text(a['layout-A-strip'])==('A / RESEARCH UPDATED' if phase=='after' else 'A / RESEARCH')
        assert fill(a['layout-A-strip'])==('B7633C' if phase=='after' else '3867A8')
        assert text(b['layout-B-strip'])=='B / DELIVERY' and fill(b['layout-B-strip'])=='467665'
        # Parent master must also list both actual layouts, not just layouts pointing to a foreign master.
        mr=E.fromstring(z.read(posixpath.join(posixpath.dirname(masters[0]),'_rels',posixpath.basename(masters[0])+'.rels')))
        listed={posixpath.normpath(posixpath.join(posixpath.dirname(masters[0]),r.get('Target'))) for r in mr if r.get('Type').endswith('/slideLayout')}
        assert set(layouts).issubset(listed)
        return {'phase':phase,'file':str(path),'layouts':layouts,'master':masters[0],'localText':local,'masterText':masterText,'A':text(a['layout-A-strip']),'B':text(b['layout-B-strip'])}
def verify(file):
    p=Path(file);stem=p.stem
    states=[inspect(p.with_name(stem+'.before.pptx'),'before'),inspect(p.with_name(stem+'.master-update.pptx'),'master'),inspect(p,'after')]
    assert states[0]['localText']==states[1]['localText']==states[2]['localText'],'Local slide text changed across inheritance updates'
    assert [r['local-key'] for r in states[2]['localText']]==['LOCAL-A1','LOCAL-A2','LOCAL-B1']
    return {'passed':True,'states':states,'scope':'Real master shared by A/A/B layouts, isolated A update, unchanged B and unchanged local slide text. Native PNG must be inspected separately.'}
if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('file');p.add_argument('--report',required=True);a=p.parse_args();out=Path(a.report)
    if out.exists():raise FileExistsError('Refusing overwrite of report')
    result=verify(a.file);out.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(result,ensure_ascii=True,indent=2))
