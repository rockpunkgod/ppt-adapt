"""Independent route-edge and real Push transition checks."""
import argparse,json,zipfile,xml.etree.ElementTree as E
from pathlib import Path
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def verify(deck):
 rows=[]
 with zipfile.ZipFile(deck) as z:
  size=E.fromstring(z.read('ppt/presentation.xml')).find('p:sldSz',N);w,h=int(size.get('cx')),int(size.get('cy'))
  for i in [1,2,3]:
   s=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'));shapes=[sp for sp in s.findall('p:cSld/p:spTree/p:sp',N) if sp.find('p:nvSpPr/p:cNvPr',N).get('name')=='route'];assert len(shapes)==1
   sp=shapes[0];xfrm=sp.find('p:spPr/a:xfrm',N);assert xfrm.find('a:off',N).attrib=={'x':'0','y':'0'}
   assert xfrm.find('a:ext',N).attrib=={'cx':str(w),'cy':str(h)}
   line=sp.find('p:spPr/a:ln',N);assert line.get('w')=='38100' and line.find('a:solidFill/a:srgbClr',N).get('val')=='C04457'
   paths=sp.findall('p:spPr/a:custGeom/a:pathLst/a:path',N);assert len(paths)==1
   path=paths[0];assert path.get('fill')=='none' and path.find('a:close',N) is None
   pts=[(int(q.get('x')),int(q.get('y'))) for q in path.findall('.//a:pt',N)];assert len(pts)==[5,6,5][i-1]
   if i>1:
    tr=s.find('p:transition',N);assert tr.get('{http://schemas.microsoft.com/office/powerpoint/2010/main}dur')=='1400'
    assert tr.find('p:push',N).get('dir')=='u'
   assert s.find('p:timing',N) is None
   rows.append({'slide':i,'first':pts[0],'last':pts[-1],'points':pts})
 for prev,nxt in zip(rows,rows[1:]):
  assert prev['last'][1]==100000 and nxt['first'][1]==0 and prev['last'][0]==nxt['first'][0]
 return {'passed':True,'routes':rows,'edgeCoordinatesEqual':True,'transition':'Upward Push 1400ms','scope':'Fixed three-page open-path fixture, not reverse navigation.'}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('deck');p.add_argument('--report',type=Path,required=True);a=p.parse_args()
 if a.report.exists():raise FileExistsError(a.report)
 a.report.write_text(json.dumps(verify(a.deck),indent=2),encoding='utf-8');print(a.report)
