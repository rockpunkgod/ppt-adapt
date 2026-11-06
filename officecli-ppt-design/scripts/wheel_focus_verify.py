"""Read-only assertions for the two-page Wheel/focus recipe; no visual claims."""
import argparse, json, zipfile
from pathlib import Path
import xml.etree.ElementTree as E

p=argparse.ArgumentParser()
p.add_argument('--deck',required=True)
p.add_argument('--output',required=True)
a=p.parse_args(); out=Path(a.output)
if out.exists(): p.error('Refusing existing report')
n={'p':'http://schemas.openxmlformats.org/presentationml/2006/main'}
with zipfile.ZipFile(a.deck) as z:
    one=E.fromstring(z.read('ppt/slides/slide1.xml'))
    two=E.fromstring(z.read('ppt/slides/slide2.xml'))
    slides=E.fromstring(z.read('ppt/presentation.xml')).findall('p:sldIdLst/p:sldId',n)
checks={}
checks['two_pages']=len(slides)==2
scales=one.findall('.//p:animScale/p:to',n)
checks['scale_85_120_85']=[x.get('x') for x in scales]==['85000','120000','85000']
opacity=one.findall('.//p:set',n)
checks['opacity_50']=len(opacity)==2 and all(x.find('p:to/p:strVal',n).get('val')=='0.5' for x in opacity)
emph=one.findall('.//p:cTn[@presetID="9"]',n)
checks['opacity_children_hold']=len(emph)==2 and all(t.get('fill')=='hold' for x in emph for t in x.findall('.//p:cBhvr/p:cTn',n))
effects=one.findall('.//p:cTn[@presetClass="emph"]',n)
checks['synchronous_delay_1000']=len(effects)==5 and all(x.find('p:stCondLst/p:cond',n).get('delay')=='1000' for x in effects)
wheel=two.findall('.//p:animEffect[@filter="wheel(1)"]',n)
checks['one_spoke_7000']=len(wheel)==1 and wheel[0].find('p:cBhvr/p:cTn',n).get('dur')=='7000'
fades=two.findall('.//p:cTn[@presetID="10"]',n)
delays=[int(x.find('p:stCondLst/p:cond',n).get('delay','0')) for x in fades][1:]
checks['core_and_node_fade_durations']=[x.find('p:cBhvr/p:cTn',n).get('dur') for x in two.findall('.//p:animEffect[@filter="fade"]',n)]==['500']+['400']*6
checks['six_angular_delays']=len(delays)==6 and all(abs(x-round(i*7000/6))<=1 for i,x in enumerate(delays))
report={'deck':str(Path(a.deck).resolve()),'checks':checks,'passed':all(checks.values()),'boundary':'Structure only. Inspect native movie separately.'}
out.parent.mkdir(parents=True,exist_ok=True);out.write_text(json.dumps(report,indent=2),encoding='utf-8')
print(json.dumps(report,indent=2));raise SystemExit(0 if report['passed'] else 1)
