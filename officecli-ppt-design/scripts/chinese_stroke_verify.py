import argparse,json,zipfile,xml.etree.ElementTree as E,hashlib
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--deck',type=Path,required=True);p.add_argument('--model',type=Path,required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'};checks={};details=[]
with zipfile.ZipFile(a.deck) as z:
 slides=[E.fromstring(z.read(f'ppt/slides/slide{i}.xml')) for i in range(1,5)]
 checks['editableOriginalBackup']=any(t.text=='山川' for t in slides[0].findall('.//a:t',ns))
 states=[]
 for i,root in enumerate(slides[1:],2):
  d={s.find('p:nvSpPr/p:cNvPr',ns).get('name'):s for s in root.findall('.//p:sp',ns)};d={n:s for n,s in d.items() if n not in ['title','scope']};states.append(d)
  checks[f's{i}StrokeCount']=len(d)==(7 if i==2 else 6)
  checks[f's{i}EditableClosedCurvedGeometry']=all(s.find('p:spPr/a:custGeom/a:pathLst/a:path/a:close',ns) is not None and s.find('p:spPr/a:custGeom/a:pathLst/a:path/a:cubicBezTo',ns) is not None for s in d.values())
  checks[f's{i}NoRasterOrWarp']=not root.findall('.//p:pic',ns) and not root.findall('.//a:prstTxWarp',ns)
  details.append({'slide':i,'names':list(d),'cubicNodes':sum(len(s.findall('.//a:cubicBezTo',ns)) for s in d.values())})
 for i,d in zip([3,4],states[1:]):checks[f's{i}ExactlyOneSharedNoDuplicates']='shared-stroke' in d and 'shan-right' not in d and 'chuan-left' not in d
 checks['middleActuallyReplaced']='chuan-middle' in states[1] and 'chuan-middle' not in states[2] and 'river-symbol' in states[2]
 def geom(s):return E.tostring(s.find('p:spPr/a:custGeom',ns))
 checks['otherFiveGeometriesUnchanged']=all(geom(states[1][n])==geom(states[2][n]) for n in states[1] if n!='chuan-middle')
 checks['newSymbolHasDifferentContour']=geom(states[1]['chuan-middle'])!=geom(states[2]['river-symbol'])
 model=json.loads(a.model.read_text(encoding='utf-8'));checks['sharedRoleOwnershipExplicit']=all(m['roles']==['shan.right','chuan.left'] for m in model['states'] if m['name']=='shared-stroke')
r={'passed':all(checks.values()),'checks':checks,'details':details,'sha256':hashlib.sha256(a.deck.read_bytes()).hexdigest(),'scope':'Geometry/role ownership only; readability requires separate human inspection.'};a.report.write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(r,ensure_ascii=False,indent=2))
