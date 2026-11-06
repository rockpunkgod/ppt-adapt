"""Check actual diagram parts and hierarchy instead of trusting COM shape appearance."""
import argparse,json,zipfile,xml.etree.ElementTree as E,posixpath,hashlib,io,struct
from pathlib import Path
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','dgm':'http://schemas.openxmlformats.org/drawingml/2006/diagram','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

def jpeg_size(data):
    f=io.BytesIO(data);assert f.read(2)==b'\xff\xd8'
    while True:
        b=f.read(1)
        while b!=b'\xff':b=f.read(1)
        while b==b'\xff':b=f.read(1)
        marker=b[0];length=struct.unpack('>H',f.read(2))[0]
        if marker in (0xC0,0xC1,0xC2):_,h,w=struct.unpack('>BHH',f.read(5));return w,h
        f.seek(length-2,1)

def verify(file,prepared_dir):
    rows=[]
    with zipfile.ZipFile(file) as z:
        for i in (1,2,3):
            slide=f'ppt/slides/slide{i}.xml';root=E.fromstring(z.read(slide))
            frames=root.findall('.//p:graphicFrame',N)
            diagrams=[f for f in frames if f.find('.//dgm:relIds',N) is not None]
            assert len(diagrams)==1,(i,'expected one real diagram graphicFrame')
            rid=diagrams[0].find('.//dgm:relIds',N).get('{'+N['r']+'}dm')
            relroot=E.fromstring(z.read(f'ppt/slides/_rels/slide{i}.xml.rels'))
            rel=next(r for r in relroot if r.get('Id')==rid)
            path=posixpath.normpath(posixpath.join('ppt/slides',rel.get('Target')))
            data=E.fromstring(z.read(path));pts=data.findall('dgm:ptLst/dgm:pt',N)
            nodes={p.get('modelId'):''.join(t.text or '' for t in p.findall('.//a:t',N)) for p in pts if p.get('type','node')=='node'}
            # OOXML ST_CxnType defaults to parOf when the native file omits @type.
            parent={c.get('destId'):c.get('srcId') for c in data.findall('dgm:cxnLst/dgm:cxn',N) if c.get('type','parOf')=='parOf'}
            def depth(key):
                seen=set();level=0
                while key in parent:
                    assert key not in seen;seen.add(key);key=parent[key];level+=1
                return level
            actual={t:depth(k) for k,t in nodes.items()}
            if i<3:assert actual=={'产品研究':1,'需求洞察':2,'原型测试':2,'反馈归纳':3},(i,actual)
            else:assert set(actual)=={'日落','步道'},actual
            expected_layout=['hierarchy1','hierarchy2','pList1'][i-1]
            doc=next(p for p in pts if p.get('type')=='doc')
            assert doc.find('dgm:prSet',N).get('loTypeId').endswith('/'+expected_layout)
            for point in pts:
                if point.get('modelId') in nodes:
                    runs=point.findall('.//a:rPr',N);assert runs
                    assert all(r.find('a:ea',N) is not None and r.find('a:ea',N).get('typeface')=='Microsoft YaHei' for r in runs)
                    assert all(r.get('sz')=='2200' for r in runs),'Expected 22pt node captions'
            row={'slide':i,'diagramData':path,'nodes':actual,'layout':expected_layout}
            if i==3:
                rels=E.fromstring(z.read('ppt/diagrams/_rels/'+posixpath.basename(path)+'.rels'))
                image_paths={r.get('Id'):posixpath.normpath(posixpath.join('ppt/diagrams',r.get('Target'))) for r in rels}
                drawing=E.fromstring(z.read('ppt/diagrams/drawing3.xml'));dn={'dsp':'http://schemas.microsoft.com/office/drawing/2008/diagram',**N}
                drawn={s.get('modelId'):s for s in drawing.findall('.//dsp:sp',dn)}
                pairs=[]
                for point in pts:
                    blip=point.find('dgm:spPr/a:blipFill/a:blip',N)
                    if blip is None:continue
                    assoc=point.find('dgm:prSet',N).get('presAssocID');label=nodes[assoc]
                    embedded=image_paths[blip.get('{'+N['r']+'}embed')];raw=z.read(embedded)
                    expected=Path(prepared_dir)/('picture-1.jpg' if label=='日落' else 'picture-2.jpg')
                    assert hashlib.sha256(raw).digest()==hashlib.sha256(expected.read_bytes()).digest(),(label,'wrong embedded photo')
                    extent=drawn[point.get('modelId')].find('dsp:spPr/a:xfrm/a:ext',dn)
                    w,h=jpeg_size(raw);ratio=int(extent.get('cx'))/int(extent.get('cy'))
                    assert abs((w/h)/ratio-1)<.002,(label,'image stretch ratio')
                    pairs.append({'caption':label,'embeddedImage':embedded,'sha256':hashlib.sha256(raw).hexdigest(),'aspectError':abs((w/h)/ratio-1)})
                assert len(pairs)==2;row['verifiedPicturePairs']=pairs
            rows.append(row)
    return {'passed':True,'file':str(file),'nativeDiagrams':rows,'limits':'Real diagram node relationships, layout and prepared embedded-picture mapping checked; separately inspect native PNG for visual quality. Source-photo authorization remains caller responsibility.'}
if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('file');p.add_argument('--report',required=True);p.add_argument('--prepared-dir',required=True);a=p.parse_args();out=Path(a.report)
    if out.exists():raise FileExistsError('Refusing overwrite of report')
    result=verify(a.file,a.prepared_dir);out.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(result,ensure_ascii=True,indent=2))
