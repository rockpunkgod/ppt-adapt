"""Independent saved OOXML and actual-hover state evidence checks (stdlib)."""
import argparse,json,posixpath
from pathlib import Path
from xml.etree import ElementTree as E
from zipfile import ZipFile
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
REL='http://schemas.openxmlformats.org/package/2006/relationships'
def verify(deck,live=None):
 checks=[]
 def check(name,value):
  checks.append({'name':name,'passed':bool(value)})
 with ZipFile(deck) as z:
  check('five states',len(E.fromstring(z.read('ppt/presentation.xml')).findall('p:sldIdLst/p:sldId',N))==5)
  total=0
  for i in range(1,6):
   root=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'))
   rels={r.get('Id'):r for r in E.fromstring(z.read(f'ppt/slides/_rels/slide{i}.xml.rels'))}
   shapes={s.find('p:nvSpPr/p:cNvPr',N).get('name'):s for s in root.findall('.//p:sp',N)}
   for j in range(1,5):
    s=shapes[f'hover-card-{j}'];hover=s.find('.//a:hlinkHover',N);selected=i==j+1
    if selected:check(f'{i}-{j} selected has no self-hover',hover is None)
    else:
     total+=1
     rid=hover.get('{'+N['r']+'}id') if hover is not None else None
     rel=rels.get(rid)
     target=posixpath.normpath(posixpath.join('ppt/slides',rel.get('Target'))) if rel is not None else None
     check(f'{i}-{j} internal hover target',hover is not None and hover.get('action')=='ppaction://hlinksldjump' and target==f'ppt/slides/slide{j+1}.xml' and rel.get('TargetMode')!='External')
    check(f'{i}-{j} no click substitute',s.find('.//a:hlinkClick',N) is None)
   if i>1:
    bio=shapes.get(f'bio-{i-1}');check(f'{i} editable profile',bio is not None and len(''.join(bio.itertext()))>90)
   check(f'{i} no rasterized slide',not root.findall('.//p:pic',N))
  check('sixteen hover links',total==16)
 if live:
  data=json.loads(Path(live).read_text(encoding='utf-8-sig'))
  check('actual hover evidence',data.get('passed') is True and len(data.get('events',[]))==15 and all(e.get('passed') is True for e in data['events']))
  check('all four targets and repeats', [e['actual'] for e in data.get('events',[]) if e['name'].startswith('hover-')]==[2,3,4,5,2,5,3])
 return {'passed':all(x['passed'] for x in checks),'checks':checks,'count':len(checks)}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('deck');p.add_argument('--live-report');p.add_argument('--report',required=True);a=p.parse_args()
 out=Path(a.report)
 if out.exists():p.error('Refusing existing report')
 result=verify(a.deck,a.live_report);out.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps({'passed':result['passed'],'checks':result['count']}));raise SystemExit(0 if result['passed'] else 1)
