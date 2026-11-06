"""Independent read-only checks for the paired-section sample; no authoring package."""
import argparse,json,zipfile,hashlib,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--deck',required=True);p.add_argument('--click-report',required=True);p.add_argument('--output',required=True);a=p.parse_args();o=Path(a.output)
if o.exists():p.error('Refusing existing report')
n={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','p14':'http://schemas.microsoft.com/office/powerpoint/2010/main','sz':'http://schemas.microsoft.com/office/powerpoint/2016/sectionzoom','slz':'http://schemas.microsoft.com/office/powerpoint/2016/slidezoom','z':'http://schemas.microsoft.com/office/powerpoint/2016/6/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','mc':'http://schemas.openxmlformats.org/markup-compatibility/2006'}
checks=[]
def check(name,value):checks.append({'name':name,'passed':bool(value)})
with zipfile.ZipFile(a.deck) as z:
 pres=E.fromstring(z.read('ppt/presentation.xml'));hub=E.fromstring(z.read('ppt/slides/slide1.xml'));sections={s.get('id'):[t.get('id') for t in s.findall('.//p14:sldId',n)] for s in pres.findall('.//p14:section',n)};cards=hub.findall('.//sz:sectionZmObj',n)
 check('five-slide-foundation',len(pres.findall('p:sldIdLst/p:sldId',n))==5)
 check('two-real-section-zoom-objects',len(cards)==2 and not hub.findall('.//slz:sldZmObj',n))
 check('two-distinct-two-slide-sections',len({s.get('sectionId') for s in cards})==2 and all(len(sections.get(s.get('sectionId'),[]))==2 for s in cards))
 check('transparent-background-native-return',all(s.find('sz:zmPr',n).get('returnToParent') in ['1','true'] and s.find('sz:zmPr',n).get('showBg') in ['0','false'] for s in cards))
 check('choice-and-fallback-no-frame-outline',len(hub.findall('.//sz:zmPr/z:spPr/a:ln/a:noFill',n))==2 and len(hub.findall('.//mc:Fallback/p:pic/p:spPr/a:ln/a:noFill',n))==2)
 for i in [2,4]:
  s=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'));tr=s.find('p:transition',n);check(f'entry-{i}-zero-auto-advance',tr is not None and tr.get('advTm')=='0' and tr.get('advClick')=='0')
 report=json.loads(Path(a.click_report).read_text(encoding='utf-8-sig'));expected=[('section-A',3),('section-A-range-end-return',1),('section-B',5),('section-B-range-end-return',1)];observed=report.get('checks',[])
 check('four-native-interaction-results',report.get('passed') is True and len(observed)==4 and all(x.get('name')==name and x.get('expected')==target and x.get('actual')==target and x.get('passed') is True for x,(name,target) in zip(observed,expected)))
result={'passed':all(x['passed'] for x in checks),'deck_sha256':hashlib.sha256(Path(a.deck).read_bytes()).hexdigest(),'checks':checks,'limits':'This checks structure plus native interaction log; inspect desktop recording middle frames separately.'};o.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2));raise SystemExit(0 if result['passed'] else 1)
