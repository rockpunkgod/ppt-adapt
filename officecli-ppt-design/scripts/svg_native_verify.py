"""Check caller SVG identity, true custom geometry, contours and persisted native edit."""
import argparse,hashlib,json,posixpath,zipfile,xml.etree.ElementTree as E
from pathlib import Path
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','svg':'http://schemas.microsoft.com/office/drawing/2016/SVG/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def name(node):
 for path in ('p:nvSpPr/p:cNvPr','p:nvPicPr/p:cNvPr','p:nvGrpSpPr/p:cNvPr'):
  x=node.find(path,N)
  if x is not None:return x.get('name')
 return None
def points(x):
 o=x.find('a:off',N);e=x.find('a:ext',N);return [int(o.get('x')),int(o.get('y')),int(e.get('cx')),int(e.get('cy'))]
def verify(base,svg,evidence,edited_color="3867A8"):
 source=Path(svg).read_bytes();sr=E.fromstring(source);palette=[p.get('fill').lstrip('#').upper() for p in sr];assert len(palette)==2 and len(set(palette))==2
 with zipfile.ZipFile(base) as z:
  r=E.fromstring(z.read('ppt/slides/slide1.xml'));tree=r.find('p:cSld/p:spTree',N);pic=next(q for q in tree if name(q)=='svg-target');assert pic.tag.endswith('}pic')
  blip=pic.find('.//svg:svgBlip',N);assert blip is not None
  rels={q.get('Id'):q.get('Target') for q in E.fromstring(z.read('ppt/slides/_rels/slide1.xml.rels'))}
  part=posixpath.normpath(posixpath.join('ppt/slides',rels[blip.get('{'+N['r']+'}embed')])).lstrip('/')
  assert z.read(part)==source,'The embedded SVG must be the actual caller source'
  box=points(pic.find('p:spPr/a:xfrm',N));ps=E.fromstring(z.read('ppt/presentation.xml')).find('p:sldSz',N);canvas=[int(ps.get('cx')),int(ps.get('cy'))]
  decorations={name(q):''.join(q.itertext()) for q in tree if name(q) not in (None,'svg-target')}
 command=json.loads((evidence/'command.json').read_text(encoding='utf-8-sig'));assert command['command']=='SVGEdit' and command['enabled']
 results=[];geometry=[]
 for filename,expected in [('converted.pptx',palette),('edited.pptx',[edited_color,palette[1]])]:
  with zipfile.ZipFile(evidence/filename) as z:
   root=E.fromstring(z.read('ppt/slides/slide1.xml'));tree=root.find('p:cSld/p:spTree',N)
   assert len(root.findall('.//p:pic',N))==0 and len(root.findall('.//a:blipFill',N))==0,'Native target cannot remain SVG/raster pictures'
   for q in tree:
    if name(q) in decorations:assert ''.join(q.itertext())==decorations[name(q)],'Non-target text changed'
   leaves=[q for q in root.findall('.//p:sp',N) if name(q) not in decorations]
   assert len(leaves)==2,'Two source paths must become two actual editable geometry leaves'
   entries=[]
   for q in leaves:
    pr=q.find('p:spPr',N);geom=pr.find('a:custGeom',N);assert geom is not None and geom.find('a:pathLst/a:path',N) is not None
    assert pr.find('a:prstGeom',N) is None
    color=pr.find('a:solidFill/a:srgbClr',N).get('val');moves=len(geom.findall('.//a:moveTo',N));closes=len(geom.findall('.//a:close',N));assert moves==closes and moves in (1,2)
    entries.append({'id':q.find('p:nvSpPr/p:cNvPr',N).get('id'),'color':color,'contours':moves,'geometrySha256':hashlib.sha256(E.tostring(geom)).hexdigest(),'boxEmu':points(pr.find('a:xfrm',N))})
   assert sorted(q['color'] for q in entries)==sorted(expected)
   assert sorted(q['contours'] for q in entries)==[1,2],'Compound native geometry must retain the separate hole contour'
   results.append({'file':filename,'nativeLeaves':entries,'pictureCount':0});geometry.append(sorted((q['id'],q['geometrySha256'],q['boxEmu']) for q in entries))
 assert geometry[0]==geometry[1],'The actual leaf-fill edit must retain native geometry and position'
 native=json.loads((evidence/'native-objects.json').read_text(encoding='utf-8-sig'));assert [q['phase'] for q in native]==['converted','edited','reopen']
 for stage in native:assert len(stage['leaves'])==2 and all(q['type']==5 and q['nodes']>0 and q['fillType']==1 for q in stage['leaves'])
 assert native[1]['leaves']==native[2]['leaves'],'Saved/reopened native objects must preserve the actual fill edit'
 return {'passed':True,'sourceSha256':hashlib.sha256(source).hexdigest(),'viewBox':[float(q) for q in sr.get('viewBox').split()],'imageBoxEmu':box,'canvasEmu':canvas,'sourcePalette':palette,'editedPalette':[edited_color,palette[1]],'results':results,'nativeObjects':native,'scope':'Simple original two-solid-path SVG, compound hole, no text/filter/external resources; not arbitrary SVG feature support.'}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--base',required=True);p.add_argument('--svg',required=True);p.add_argument('--evidence',required=True);p.add_argument('--report',required=True);p.add_argument('--edited-color',default='3867A8');a=p.parse_args();out=Path(a.report)
 if out.exists():raise FileExistsError('Refusing overwrite')
 result=verify(a.base,a.svg,Path(a.evidence),a.edited_color.upper());out.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(out)
