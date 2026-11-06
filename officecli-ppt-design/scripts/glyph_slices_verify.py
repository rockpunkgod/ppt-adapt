import argparse,json,zipfile
from pathlib import Path
import xml.etree.ElementTree as E
p=argparse.ArgumentParser();p.add_argument('deck',type=Path);p.add_argument('--report',type=Path,required=True);p.add_argument('--native-evidence',type=Path);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
with zipfile.ZipFile(a.deck) as z:
 slides=[E.fromstring(z.read(f'ppt/slides/slide{i}.xml')) for i in (1,2)]
 byname=[{s.find('p:nvSpPr/p:cNvPr',ns).get('name'):s for s in slide.findall('.//p:sp',ns)} for slide in slides]
 assert ''.join(t.text or '' for t in byname[0]['number-seed'].findall('.//a:t',ns))=='01'
 part=byname[0]['partial-gold-glyph'];cut=byname[1]['corner-subtracted-glyph']
 for q in (part,cut):
  assert q.find('p:spPr/a:custGeom/a:pathLst/a:path',ns) is not None
  assert not ''.join(t.text or '' for t in q.findall('.//a:t',ns))
  assert q.find('p:spPr/a:solidFill/a:srgbClr',ns).get('val')=='A57119'
 assert abs(int(part.find('p:spPr/a:xfrm/a:off',ns).get('y'))-4356000)<1270
 assert not any(n in byname[1] for n in ('number-seed','corner-cut','outline-cover','full-glyph'))
 assert slides[1].find('.//p:cxnSp/p:nvCxnSpPr/p:cNvPr',ns).get('name')=='independent-slash'
 assert all(not s.findall('.//p:pic',ns) for s in slides)
result={'passed':True,'lowerBoundaryCm':12.1,'partialGlyphIsTrueGeometry':True,'cornerIsTrueGeometry':True,'noTriangleCover':True,'noRasterImages':True,'limit':'Native before/after pixel tests and visual checks remain required.'}
if a.native_evidence:
 from PIL import Image,ImageChops
 imgs={name:Image.open(a.native_evidence/(name+'.png')).convert('RGB') for name in ['before-1','slide-1','before-2','slide-2']}
 assert all(im.size==(1600,900) for im in imgs.values()),'Expected native1600x900 exports'
 def color_count(im,color,box):
  return sum(max(abs(c-t) for c,t in zip(px,color))<5 for px in im.crop(box).getdata())
 gold=(165,113,25);lower=(180,573,700,700);above=(180,300,700,570);all_glyph=(180,300,700,700)
 original_lower_white=color_count(imgs['before-1'],(255,255,255),lower)
 lower_gold=color_count(imgs['slide-1'],gold,lower)
 assert ImageChops.difference(imgs['before-1'],imgs['slide-1']).crop((180,300,650,570)).getbbox() is None,'Upper original glyph moved/changed'
 assert color_count(imgs['slide-1'],gold,above)==0,'Gold leaked above the boundary'
 assert lower_gold>5000 and original_lower_white>5000 and 0.95<lower_gold/original_lower_white<1.05,'Lower color no longer matches original glyph footprint'
 before_gold=color_count(imgs['before-2'],gold,all_glyph);after_gold=color_count(imgs['slide-2'],gold,all_glyph)
 assert before_gold-after_gold>1000 and after_gold>30000,'Corner subtraction missing or whole glyph lost'
 result['nativePixels']={'passed':True,'upperGlyphUnchanged':True,'goldAboveBoundary':0,'lowerGoldPixels':lower_gold,'originalLowerWhitePixels':original_lower_white,'cornerGoldPixelsRemoved':before_gold-after_gold,'remainingGoldPixels':after_gold}
 result['limit']='Native PNG pixel checks passed for this fixed1600x900 sample. Review exported pages and recheck when changing font/text/layout.'

a.report.parent.mkdir(parents=True,exist_ok=True);a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print('True glyph intersection/subtraction verified')
