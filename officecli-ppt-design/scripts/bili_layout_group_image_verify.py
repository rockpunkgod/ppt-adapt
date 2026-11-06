"""Check actual image parts, group transforms and crop-to-common-coordinate mapping."""
import argparse,hashlib,json,posixpath,zipfile,xml.etree.ElementTree as E
from pathlib import Path
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def box(x):
    off=x.find('a:off',N);ext=x.find('a:ext',N)
    assert x.get('rot','0')=='0','Rotated pieces need a separate mapping test'
    return [int(off.get('x')),int(off.get('y')),int(ext.get('cx')),int(ext.get('cy'))]
def rels(z,slide):
    part=f'ppt/slides/slide{slide}.xml'
    return {r.get('Id'):posixpath.normpath(posixpath.join('ppt/slides',r.get('Target'))).lstrip('/') for r in E.fromstring(z.read(f'ppt/slides/_rels/slide{slide}.xml.rels'))}
def sha(z,slide,fill):
    rid=fill.find('a:blip',N).get('{'+N['r']+'}embed');part=rels(z,slide)[rid]
    return hashlib.sha256(z.read(part)).hexdigest(),part
def verify(file,image):
    original=hashlib.sha256(Path(image).read_bytes()).hexdigest();result=[]
    with zipfile.ZipFile(file) as z:
        r=E.fromstring(z.read('ppt/slides/slide1.xml'));ref=next(s for s in r.findall('.//p:sp',N) if s.find('p:nvSpPr/p:cNvPr',N).get('name')=='reference')
        rb=box(ref.find('p:spPr/a:xfrm',N));assert sha(z,1,ref.find('p:spPr/a:blipFill',N))[0]==original
        for slide in (2,3):
            root=E.fromstring(z.read(f'ppt/slides/slide{slide}.xml'));groups=root.findall('p:cSld/p:spTree/p:grpSp',N);assert len(groups)==1
            group=groups[0];gx=group.find('p:grpSpPr/a:xfrm',N);gb=box(gx);co=gx.find('a:chOff',N);ce=gx.find('a:chExt',N)
            assert all(abs(a-b)<4 for a,b in zip(rb,gb)),'Group bounding box must match reference image'
            sx=gb[2]/int(ce.get('cx'));sy=gb[3]/int(ce.get('cy'));pieces=group.findall('p:sp',N);assert len(pieces)==4
            actual=[]
            for piece in pieces:
                pr=piece.find('p:spPr',N);cb=box(pr.find('a:xfrm',N));world=[gb[0]+(cb[0]-int(co.get('x')))*sx,gb[1]+(cb[1]-int(co.get('y')))*sy,cb[2]*sx,cb[3]*sy]
                u=(world[0]-rb[0])/rb[2];v=(world[1]-rb[1])/rb[3];w=world[2]/rb[2];h=world[3]/rb[3]
                assert pr.find('a:grpFill',N) is not None and pr.find('a:blipFill',N) is None,'Child must inherit actual parent image'
                fill=group.find('p:grpSpPr/a:blipFill',N);assert fill is not None
                digest,part=sha(z,slide,fill);assert digest==original
                assert fill.find('a:stretch/a:fillRect',N) is not None
                actual.append({'name':piece.find('p:nvSpPr/p:cNvPr',N).get('name'),'normalizedBox':[u,v,w,h],'inheritedFill':True,'imagePart':part,'worldBoxEmu':world})
            pattern=[(0,0,0.5,0.5),(0.5,0,0.5,0.5),(0,0.5,0.5,0.5),(0.5,0.5,0.5,0.5)] if slide==2 else [(0,0,0.4,0.4),(0.6,0,0.4,0.4),(0,0.6,0.4,0.4),(0.6,0.6,0.4,0.4)]
            for p,expected in zip(actual,pattern):assert all(abs(a-b)<0.00001 for a,b in zip(p['normalizedBox'],expected))
            result.append({'slide':slide,'group':group.find('p:nvGrpSpPr/p:cNvPr',N).get('name'),'pieces':actual})
        inherited=None
        if 'ppt/slides/slide5.xml' in z.namelist():
            root=E.fromstring(z.read('ppt/slides/slide5.xml'));g=root.find('p:cSld/p:spTree/p:grpSp',N);gf=g.find('p:grpSpPr/a:blipFill',N)
            assert gf is not None and sha(z,5,gf)[0]==original
            children=g.findall('p:sp',N);assert len(children)==4 and all(s.find('p:spPr/a:grpFill',N) is not None and s.find('p:spPr/a:blipFill',N) is None for s in children)
            inherited={'slide':5,'parentImageFill':True,'childInheritedFills':4,'limits':'Structural inheritance only; native pixels must confirm global mapping rather than per-child repetition.'}
    return {'passed':True,'file':str(file),'imageSha256':original,'referenceBoxEmu':rb,'mappedGroups':result,'inheritedProbe':inherited,'limits':'Real parent fill and inherited child fills; native pixels required. Tested axis-aligned rectangles only; arbitrary rotated geometry not yet verified.'}
if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('file');p.add_argument('--image',required=True);p.add_argument('--report',required=True);a=p.parse_args();out=Path(a.report)
    if out.exists():raise FileExistsError('Refusing overwrite')
    result=verify(a.file,a.image);out.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(out)
