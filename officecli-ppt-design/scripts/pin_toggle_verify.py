"""Read-only OOXML + physical-click report validation for same-slide Pin toggle."""
import argparse,json,zipfile
from pathlib import Path
import xml.etree.ElementTree as E
p=argparse.ArgumentParser();p.add_argument('--deck',required=True);p.add_argument('--click-report',required=True);p.add_argument('--output',required=True);a=p.parse_args()
o=Path(a.output)
if o.exists():p.error('Refusing existing report')
n={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
with zipfile.ZipFile(a.deck) as z:
 r=E.fromstring(z.read('ppt/slides/slide1.xml'));pres=E.fromstring(z.read('ppt/presentation.xml'))
names={x.get('name'):x.get('id') for x in r.findall('.//p:cNvPr',n)}
pin=names['trigger-pin'];label=names['trigger-label']
seq=r.findall('.//p:cTn[@nodeType="interactiveSeq"]',n)
effects=r.findall('.//p:cTn[@presetClass="entr"]',n)+r.findall('.//p:cTn[@presetClass="exit"]',n)
checks={
 'one_slide':len(pres.findall('p:sldIdLst/p:sldId',n))==1,
 'one_true_interactive_sequence':len(seq)==1,
 'no_main_sequence':not r.findall('.//p:cTn[@nodeType="mainSeq"]',n),
 'one_entrance_one_exit':len(r.findall('.//p:cTn[@presetClass="entr"]',n))==1 and len(r.findall('.//p:cTn[@presetClass="exit"]',n))==1,
 'trigger_on_named_pin':any(x.get('evt')=='onClick' and x.find('p:tgtEl/p:spTgt',n) is not None and x.find('p:tgtEl/p:spTgt',n).get('spid')==pin for x in r.findall('.//p:cond',n)),
 'both_effects_target_label':len(effects)==2 and all(x.find('.//p:spTgt',n).get('spid')==label for x in effects),
 'no_hyperlinks':not r.findall('.//a:hlinkClick',n),
 'two_fades_500ms':len(r.findall('.//p:animEffect[@filter="fade"]',n))==2 and all(x.find('p:cBhvr/p:cTn',n).get('dur')=='500' for x in r.findall('.//p:animEffect[@filter="fade"]',n)),
}
log=json.loads(Path(a.click_report).read_text(encoding='utf-8-sig'));rows={x['name']:x for x in log['checks']}
for name in ['state-0-closed','state-blank-before','state-1-open','state-blank-after','state-2-closed','state-3-reopen']:
 checks[name]=rows.get(name,{}).get('passed') is True
checks['native_physical_report_passed']=log.get('passed') is True
result={'checks':checks,'passed':all(checks.values()),'boundary':'Pixel samples and native screenshots/video must be interpreted together.'}
o.parent.mkdir(parents=True,exist_ok=True);o.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2));raise SystemExit(0 if result['passed'] else 1)
