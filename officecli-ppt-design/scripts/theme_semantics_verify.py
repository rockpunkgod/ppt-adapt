"""Independently verify theme semantics, PowerPoint round-trip and native pixels."""
import argparse,json,zipfile
from pathlib import Path
import xml.etree.ElementTree as E
from PIL import Image,ImageChops
p=argparse.ArgumentParser();p.add_argument('--input-directory',type=Path,required=True);p.add_argument('--evidence-directory',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args()
if a.output.exists():p.error('Refusing existing report')
n={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'};checks={};metrics={}
def need(name,ok):checks[name]=bool(ok)
source_slide_bytes={}
for stage in ['source','roundtrip']:
 for state in ['base','color','font']:
  file=a.input_directory/f'theme-{state}.pptx' if stage=='source' else a.evidence_directory/f'roundtrip-{state}.pptx'
  tag=f'{stage}-{state}'
  with zipfile.ZipFile(file) as z:
   th=E.fromstring(z.read('ppt/theme/theme1.xml'));slides=[E.fromstring(z.read(f'ppt/slides/slide{i}.xml')) for i in [1,2]]
   if stage=='source':source_slide_bytes[state]=[E.tostring(E.fromstring(z.read(f'ppt/slides/slide{i}.xml'))) for i in [1,2]]
  expected_color='1C6561' if state=='base' else 'AE4E32';expected_ea='SimSun' if state=='font' else 'Microsoft YaHei'
  need(tag+'-accent1',th.find('.//a:clrScheme/a:accent1/a:srgbClr',n).get('val')==expected_color)
  need(tag+'-majorEA',th.find('.//a:fontScheme/a:majorFont/a:ea',n).get('typeface')==expected_ea)
  need(tag+'-minorEA',th.find('.//a:fontScheme/a:minorFont/a:ea',n).get('typeface')=='Microsoft YaHei')
  need(tag+'-majorLatin',th.find('.//a:fontScheme/a:majorFont/a:latin',n).get('typeface')=='Arial')
  named=[{q.find('p:nvSpPr/p:cNvPr',n).get('name'):q for q in s.findall('.//p:sp',n)} for s in slides]
  need(tag+'-schemeFill',named[0]['theme-card'].find('p:spPr/a:solidFill/a:schemeClr',n).get('val')=='accent1')
  need(tag+'-directFill',named[0]['direct-card'].find('p:spPr/a:solidFill/a:srgbClr',n).get('val')=='1C6561')
  need(tag+'-secondSlideScheme',named[1]['theme-rule'].find('p:spPr/a:solidFill/a:schemeClr',n).get('val')=='accent1')
  need(tag+'-textSchemeColor',named[0]['theme-footer'].find('.//a:rPr/a:solidFill/a:schemeClr',n).get('val')=='accent1')
  for name,ea,latin in [('major-reference','+mj-ea','+mj-lt'),('minor-reference','+mn-ea','+mn-lt'),('latin-reference','+mj-ea','+mj-lt'),('major-direct','Microsoft YaHei','Arial'),('minor-direct','Microsoft YaHei','Arial'),('latin-direct','Microsoft YaHei','Arial')]:
   q=named[1][name];need(tag+'-'+name+'-refs',q.find('.//a:rPr/a:ea',n).get('typeface')==ea and q.find('.//a:rPr/a:latin',n).get('typeface')==latin)
  need(tag+'-noRaster',all(not s.findall('.//p:pic',n) for s in slides))
need('color_and_font_changes_preserved_slide_object_trees',source_slide_bytes['base']==source_slide_bytes['color']==source_slide_bytes['font'])
images={}
for state in ['base','color','font']:
 for page in [1,2]:
  first=Image.open(a.evidence_directory/f'{state}-slide-{page}.png').convert('RGB');reopen=Image.open(a.evidence_directory/f'{state}-reopen-{page}.png').convert('RGB');images[(state,page)]=first
  need(f'{state}-{page}-native-size',first.size==(1600,900))
  need(f'{state}-{page}-saveReopenPixelIdentity',ImageChops.difference(first,reopen).getbbox() is None)
 def close_color(px,want):return sum(abs(c-t) for c,t in zip(px,want))<=3
 theme_rgb=(28,101,97) if state=='base' else (174,78,50)
 need(state+'-nativeThemeCardColor',close_color(images[(state,1)].getpixel((150,600)),theme_rgb))
 need(state+'-nativeDirectCardColor',close_color(images[(state,1)].getpixel((950,600)),(28,101,97)))
 need(state+'-nativeSecondPageThemeRule',close_color(images[(state,2)].getpixel((120,225)),theme_rgb))
 need(state+'-nativeSecondPageDirectRule',close_color(images[(state,2)].getpixel((950,225)),(28,101,97)))
 need(state+'-directCardPixelsUnchanged',ImageChops.difference(images[('base',1)].crop((830,280,1520,650)),images[(state,1)].crop((830,280,1520,650))).getbbox() is None)
before,after=images[('color',2)],images[('font',2)]
for name,box in [('major',(80,330,760,420)),('minor',(80,525,760,615)),('latin',(80,715,760,810)),('direct',(830,270,1520,820))]:
 diff=ImageChops.difference(before.crop(box),after.crop(box));count=sum(any(v for v in px) for px in diff.getdata());metrics[name+'ChangedPixels']=count
 need(name+'-fontPixelScope',count>500 if name=='major' else count==0)
readback=json.loads((a.evidence_directory/'native-readback.json').read_text(encoding='utf-8-sig'))
for row in readback:
 state=row['state'];stage=row['stage'];tag=state+'-'+stage
 by={(q['slide'],q['name']):q for q in row['shapes']}
 need(tag+'-nativeThemeColorIdentity',by[(1,'theme-card')]['fillTheme']==5 and by[(1,'direct-card')]['fillTheme']==0)
 need(tag+'-nativeResolvedMajorEA',by[(2,'major-reference')]['fontEA']==('SimSun' if state=='font' else 'Microsoft YaHei'))
 need(tag+'-nativeResolvedMinorAndDirectEA',by[(2,'minor-reference')]['fontEA']=='Microsoft YaHei' and by[(2,'major-direct')]['fontEA']=='Microsoft YaHei')
 need(tag+'-nativeResolvedLatin',by[(2,'latin-reference')]['fontLatin']=='Arial')
result={'passed':all(checks.values()),'checks':checks,'metrics':metrics,'backend':'officecli-native','boundary':'Fixed two-page theme-reference experiment. Native COM only renders/saves/reads; it does not author the theme. No placeholder/master animation claim.'}
a.output.parent.mkdir(parents=True,exist_ok=True);a.output.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps({'passed':result['passed'],'checks':len(checks),'failed':[k for k,v in checks.items() if not v],'metrics':metrics},indent=2));raise SystemExit(0 if result['passed'] else 1)
