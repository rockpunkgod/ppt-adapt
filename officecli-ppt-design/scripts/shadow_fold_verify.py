"""Independent native shadow geometry/crop/alpha checks, not visual acceptance."""
import argparse,json,zipfile,posixpath,struct
from pathlib import Path
import xml.etree.ElementTree as E
p=argparse.ArgumentParser();p.add_argument('deck',type=Path);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
rows=[]
with zipfile.ZipFile(a.deck) as z:
 for i in (1,2):
  s=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'))
  g=s.find('.//p:grpSp',ns);assert g is not None
  rot=int(g.find('p:grpSpPr/a:xfrm',ns).get('rot'));assert rot==32*60000
  pic=g.find('p:pic',ns);assert pic is not None
  src=pic.find('p:blipFill/a:srcRect',ns);assert src.get('r')=='50000'
  blip=pic.find('p:blipFill/a:blip',ns);rid=blip.get('{'+ns['r']+'}embed')
  rels=E.fromstring(z.read(f'ppt/slides/_rels/slide{i}.xml.rels'))
  rel=next(r for r in rels if r.get('Id')==rid)
  target=rel.get('Target')
  image=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('ppt/slides',target))
  png=z.read(image);assert png.startswith(b'\x89PNG\r\n\x1a\n')
  w,h=struct.unpack('>II',png[16:24]);assert png[25]==6,'Native transparent RGBA required'
  assert abs(w/h/(4/14)-1)<.005,'Ellipse raster was exported with distorted slide-relative scaling'
  extent=pic.find('p:spPr/a:xfrm/a:ext',ns)
  assert abs((int(extent.get('cx'))/int(extent.get('cy')))/(.5*w/h)-1)<.005,'Cropped shadow stretched'
  if i==2:
   assert blip.find('a:lum',ns).get('bright')=='-100000'
   assert blip.find('a:alphaModFix',ns).get('amt')=='50000'
  rows.append({'slide':i,'png':[w,h],'rightCrop':.5,'rotation':32,'picture_alpha':.5 if i==2 else 1})
a.report.parent.mkdir(parents=True,exist_ok=True)
a.report.write_text(json.dumps({'passed':True,'slides':rows,'limit':'Native appearance and blending must be reviewed separately.'},indent=2),encoding='utf-8')
print('Shadow geometry, native alpha and crop verified')
