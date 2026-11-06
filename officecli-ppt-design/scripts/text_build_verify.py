"""Saved native text targets and independent actual-frame evidence checks."""
import argparse,json
from pathlib import Path
from xml.etree import ElementTree as E
from zipfile import ZipFile
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def slides(path):
 with ZipFile(path) as z:return [E.fromstring(z.read(f'ppt/slides/slide{i}.xml')) for i in (1,2,3)]
def textshape(root):return next(s for s in root.findall('.//p:sp',N) if s.find('p:nvSpPr/p:cNvPr',N).get('name')=='text-build-target')
p=argparse.ArgumentParser();p.add_argument('--base',required=True);p.add_argument('--deck',required=True);p.add_argument('--measure',required=True);p.add_argument('--report',required=True);a=p.parse_args();out=Path(a.report)
if out.exists():p.error('New report required')
checks=[]
def check(name,value):checks.append({'name':name,'passed':bool(value)})
for i,(before,root) in enumerate(zip(slides(a.base),slides(a.deck)),1):
 shape=textshape(root);original=textshape(before);target=shape.find('p:nvSpPr/p:cNvPr',N).get('id')
 check(f'{i} original editable text preserved',[t.text for t in shape.findall('.//a:t',N)]==[t.text for t in original.findall('.//a:t',N)])
 check(f'{i} only one target text box',len([s for s in root.findall('.//p:sp',N) if s.find('p:nvSpPr/p:cNvPr',N).get('name')=='text-build-target'])==1)
 check(f'{i} no raster substitution',not root.findall('.//p:pic',N))
 transition=root.find('p:transition',N)
 check(f'{i} physical page click enabled',transition is None or transition.get('advClick','1')=='1')
 effects=[c for c in root.findall('.//p:cTn',N) if c.get('presetClass')=='entr']
 check(f'{i} native entrance count',len(effects)==(3 if i==1 else 1))
 check(f'{i} native effect preset',all(e.get('presetID')==('10' if i==1 else '19') for e in effects))
 check(f'{i} all effects genuinely target one shape',all(s.get('spid')==target for e in effects for s in e.findall('.//p:spTgt',N)))
 if i==1:
  check('three real paragraphs',len(shape.findall('p:txBody/a:p',N))==3)
  check('three separate click groups',all(e.get('nodeType')=='clickEffect' for e in effects))
  check('non-overlapping actual paragraph targets', [{(r.get('st'),r.get('end')) for r in e.findall('.//p:pRg',N)} for e in effects]==[{('0','0')},{('1','1')},{('2','2')}])
 else:
  it=effects[0].find('p:iterate',N)
  check(f'{i} unit type',it is not None and it.get('type')==('lt' if i==2 else 'wd'))
  check(f'{i} ten percent stagger',it is not None and it.find('p:tmPct',N).get('val')=='10000')
  durations={c.get('dur') for c in effects[0].findall('.//p:cTn',N) if c.get('dur') not in (None,'indefinite')}
  check(f'{i} source 200ms duration','200' in durations)
data=json.loads(Path(a.measure).read_text(encoding='utf-8-sig'))
check('independent actual movie measure',data.get('passed') is True)
check('actual unique PTS decoded',data['actualDecodedFrames']==len(data['actualPTS']) and data['actualDecodedFrames']>100)
check('paragraph invisible/first/two/all states',all(r['passed'] for r in data['paragraph']) and len(data['paragraph'])==4)
check('letter and word partial middle frames',len(data['motion'])==2 and all(r['passed'] and r['partialFrameCount']>0 for r in data['motion']))
result={'passed':all(c['passed'] for c in checks),'count':len(checks),'checks':checks}
out.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps({'passed':result['passed'],'checks':len(checks)}));raise SystemExit(0 if result['passed'] else 1)
