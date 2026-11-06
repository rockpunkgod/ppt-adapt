"""Independent checks of native blend witnesses, backend labels and roundtrip."""
import argparse,json,hashlib,zipfile,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--deck',type=Path,required=True);p.add_argument('--measurements',type=Path,required=True);p.add_argument('--math-audit',type=Path,required=True);p.add_argument('--fields',type=Path,required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():p.error('Fresh report required')
r=json.loads(a.measurements.read_text());math=json.loads(a.math_audit.read_text());checks=[]
def check(name,ok,**meta):checks.append({'name':name,'pass':bool(ok),**meta})
expected={'base':'backdrop','source':'normal-encoded','image-mult':'multiply-encoded','image-over':'overlay-encoded','image-screen':'screen-encoded','normal50':'normal-50-encoded','mult50':'multiply-50-encoded','softlight-ref':'softlight-encoded','local-softlight':'local-softlight-encoded','local-overlay':'local-overlay-encoded'}
for row in r['panels']:
 check('Native reopen pixel identity '+str(row['slide'])+':'+row['panel'],row['roundtripWholePageMax']==0)
 if row['panel'] in expected:
  best=row['bestCandidates'][0];check('RGB semantic '+row['panel'],best['reference']==expected[row['panel']] and best['maxAbsRGB']<=2,best=best)
check('858 independent local samples',len(r['localWitnesses'])>=850 and r['localInsideCount']>200 and r['localOutsideCount']>200,inside=r['localInsideCount'],outside=r['localOutsideCount'])
check('Native local mult maximum 1 RGB level',r['localMaxError']<=1,actual=r['localMaxError'])
check('Independent scalar/local formula audit',math['passed']==math['count']==14)
check('Input unchanged during native SaveAs/reopen',bool(r.get('inputIntegrity')) and r['inputIntegrity']['unchanged'])
manifest=json.loads((a.fields/'manifest.json').read_text())
for name,sha in manifest['files'].items():check('Asset SHA256 '+name,hashlib.sha256((a.fields/name).read_bytes()).hexdigest()==sha)
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
with zipfile.ZipFile(a.deck) as z:
 for i in [1,2]:
  root=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'))
  for pic in root.findall('.//p:pic',ns):
   name=pic.find('p:nvPicPr/p:cNvPr',ns).get('name');effects=pic.findall('.//a:fillOverlay',ns)
   if name in ['image-mult','image-over','image-screen','mult50','local-mult']:check('Real native effect '+name,len(effects)==1 and effects[0].get('blend')==('over' if name=='image-over' else 'screen' if name=='image-screen' else 'mult'))
   if name in ['softlight-ref','local-softlight','local-overlay']:check('Explicit external output has no fake native effect '+name,len(effects)==0)
  check('No failed native alpha Overlay recipe '+str(i),not any(n.get('name')=='over50' for n in root.findall('.//p:cNvPr',ns)))
check('Native effect nodes survive SaveAs',r['roundtripNodes'][0]['blendValues']==r['roundtripNodes'][2]['blendValues'] and r['roundtripNodes'][1]['blendValues']==r['roundtripNodes'][3]['blendValues'])
out={'passed':sum(c['pass'] for c in checks),'count':len(checks),'inputSHA256':hashlib.sha256(a.deck.read_bytes()).hexdigest(),'checks':checks};a.report.write_text(json.dumps(out,indent=2),encoding='utf-8');print(out['passed'],'/',out['count']);assert all(c['pass'] for c in checks)
