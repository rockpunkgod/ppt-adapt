"""Check native editable geometry; visual/motion review remains required."""
import argparse,json,zipfile
from pathlib import Path
import xml.etree.ElementTree as E
p=argparse.ArgumentParser();p.add_argument('deck',type=Path);p.add_argument('--mode',choices=['cards','background'],required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
rows=[]
with zipfile.ZipFile(a.deck) as z:
 for i in (1,2):
  s=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'))
  if a.mode=='cards':
   found=[]
   for g in s.findall('.//p:spTree/p:grpSp',ns):
    name=g.find('p:nvGrpSpPr/p:cNvPr',ns).get('name')
    if not name.startswith('!!true-card-'):continue
    shapes=g.findall('p:sp',ns)
    native=[x for x in shapes if x.find('p:spPr/a:custGeom/a:pathLst/a:path',ns) is not None]
    assert len(native)==1 and len(shapes)==3,(name,'expected union glyph/card plus two text shapes')
    assert not any((t.text or '').strip() for t in native[0].findall('.//a:t',ns)),(name,'union still contains editable digits')
    assert len(s.findall('.//p:pic',ns))==0,'Unexpected flattened picture'
    xf=g.find('p:grpSpPr/a:xfrm',ns)
    found.append({'name':name,'x':int(xf.find('a:off',ns).get('x')),'y':int(xf.find('a:off',ns).get('y')),'width':int(xf.find('a:ext',ns).get('cx')),'height':int(xf.find('a:ext',ns).get('cy'))})
   assert len(found)==4,found
   rows.append(found)
  else:
   glyph=[x for x in s.findall('.//p:sp',ns) if x.find('p:nvSpPr/p:cNvPr',ns).get('name')=='!!background-window']
   assert len(glyph)==1
   g=glyph[0]
   assert g.get('useBgFill')=='1' and g.find('p:spPr/a:custGeom',ns) is not None
   assert g.find('.//a:blipFill',ns) is None,'Object-local picture fill is not slide background sampling'
   off=g.find('p:spPr/a:xfrm/a:off',ns)
   rows.append({'x':int(off.get('x')),'y':int(off.get('y')),'backgroundFill':True})
if a.mode=='cards':
 for first,second in zip(sorted(rows[0],key=lambda r:r['name']),sorted(rows[1],key=lambda r:r['name'])):
  assert first['name']==second['name'] and first['width']==second['width'] and first['height']==second['height'],'Morph scales card instead of translation'
  assert first['x']!=second['x'],'No card displacement'
else:assert rows[0]['x']!=rows[1]['x'] and rows[0]['y']!=rows[1]['y']
a.report.parent.mkdir(parents=True,exist_ok=True)
a.report.write_text(json.dumps({'passed':True,'mode':a.mode,'slides':rows,'limit':'Structural checks do not establish rendered appearance or intermediate motion.'},indent=2),encoding='utf-8')
print('Native structure verified:',a.mode)
