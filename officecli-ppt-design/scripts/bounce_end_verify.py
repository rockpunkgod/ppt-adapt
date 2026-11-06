"""Read-only Fly/Bounce End structure plus actual-PTS trajectory assertions."""
import argparse,json,zipfile
from pathlib import Path
import xml.etree.ElementTree as E
p=argparse.ArgumentParser();p.add_argument('--deck',required=True);p.add_argument('--measure-report',required=True);p.add_argument('--output',required=True);a=p.parse_args();out=Path(a.output)
if out.exists():p.error('Refusing existing output')
n={'p':'http://schemas.openxmlformats.org/presentationml/2006/main'};be='{http://schemas.microsoft.com/office/powerpoint/2010/main}bounceEnd'
with zipfile.ZipFile(a.deck) as z:r=E.fromstring(z.read('ppt/slides/slide1.xml'))
# Native PowerPoint executes modern Choice; Fallback deliberately remains linear.
for parent in r.iter():
 for child in list(parent):
  if child.tag=='{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback':parent.remove(child)
ids={x.get('name'):x.get('id') for x in r.findall('.//p:cNvPr',n)}
effects=r.findall('.//p:cTn[@presetID]',n)
checks={'two_fly_presets_not_bounce_preset':len(effects)==2 and all(x.get('presetID')=='2' and x.get('presetClass')=='entr' for x in effects)}
checks['effect_level_bounce_preset_75000']=any(x.get('{http://schemas.microsoft.com/office/powerpoint/2010/main}presetBounceEnd')=='75000' for x in effects)
checks['both_delay_1000']=all(x.find('p:stCondLst/p:cond',n).get('delay')=='1000' for x in effects)
checks['both_native_xy_2000ms']=all(len(x.findall('.//p:anim',n))==2 and all(y.find('p:cBhvr/p:cTn',n).get('dur')=='2000' for y in x.findall('.//p:anim',n)) for x in effects)
for name in ['control','bounce']:
 anims=[x for x in r.findall('.//p:anim',n) if x.find('.//p:spTgt',n).get('spid')==ids[name]]
 checks[name+'_bounce_attributes']=len(anims)==2 and all(x.get(be)==('75000' if name=='bounce' else None) for x in anims)
d=json.loads(Path(a.measure_report).read_text(encoding='utf-8-sig'));rows=d['rows']
paths={name:[x['cy'] for x in rows if x['name']==name and 1<=x['seconds']<3.8 and x['area']>15000 and x['top']>0 and x['bottom']<719] for name in ['control','bounce']}
dif={name:[b-a for a,b in zip(v,v[1:])] for name,v in paths.items()}
checks['actual_pts_samples']=d.get('uniqueRecordedFrames',0)>=180 and d.get('analysisWindowSeconds')==[1.0,3.8]
checks['control_monotonic_with_1px_tolerance']=len(dif['control'])>20 and min(dif['control'])>-1
checks['bounce_nonmonotonic']=len(dif['bounce'])>20 and min(dif['bounce'])<-2 and max(dif['bounce'])>2
checks['bounce_clear_deviation']=d['summary']['bounce']['maxPositiveDeviationFromFinalPixels']>15
checks['control_no_material_deviation']=abs(d['summary']['control']['maxPositiveDeviationFromFinalPixels'])<1
for name in ['control','bounce']:
 final=d['summary'][name]['finalCy'];late=[x['cy'] for x in rows if x['name']==name and 4<x['seconds']<5]
 checks[name+'_stable_final']=len(late)>10 and max(abs(v-final) for v in late)<1
report={'passed':all(checks.values()),'checks':checks,'summary':d['summary'],'boundary':'One fixed 2s Fly/Intensity0.75 vertical test; no universal timing or direction guarantee.'}
out.parent.mkdir(parents=True,exist_ok=True);out.write_text(json.dumps(report,indent=2),encoding='utf-8');print(json.dumps(report,indent=2));raise SystemExit(0 if report['passed'] else 1)
