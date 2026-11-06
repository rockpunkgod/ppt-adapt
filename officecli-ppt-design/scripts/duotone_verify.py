import argparse,hashlib,json,zipfile,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--deck',type=Path,required=True);p.add_argument('--source',type=Path,required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'};checks={};src=a.source.read_bytes()
with zipfile.ZipFile(a.deck) as z:
 pictures=[z.read(n) for n in z.namelist() if n.startswith('ppt/media/') and n.endswith('.png')];checks['allEmbeddedPNGsOriginalBytes']=len(pictures)>=1 and all(x==src for x in pictures)
 for i in [1,2,3]:
  s=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'));duo=s.findall('.//a:duotone',ns);pics=s.findall('.//p:pic',ns);checks[f's{i}onePictureOnly']=len(pics)==1
  checks[f's{i}noCoverShapes']=len(s.findall('p:cSld/p:spTree/p:sp',ns))==1
  if i==1:checks['s1NoColorEffect']=not duo
  else:
   expect=['201044','F2D787'] if i==2 else ['F2D787','201044'];checks[f's{i}nativeDuotoneColors']=len(duo)==1 and [c.get('val') for c in duo[0]]==expect
   blip=pics[0].find('p:blipFill/a:blip',ns);checks[f's{i}duotoneDirectBlipChild']=list(blip)==duo
 result={'passed':all(checks.values()),'checks':checks,'originalSHA256':hashlib.sha256(src).hexdigest(),'embeddedPNGCount':len(pictures),'note':'Only the picture blip carries a native duotone effect; no grayscale/recolor bitmap overwrite or colored-cover shape.'}
a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2))
