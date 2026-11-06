import argparse,array,hashlib,io,json,math,posixpath,statistics,wave,zipfile
from pathlib import Path
from xml.etree import ElementTree as E
p=argparse.ArgumentParser();p.add_argument('deck',type=Path);p.add_argument('--assets',type=Path,required=True);p.add_argument('--native',type=Path,required=True);p.add_argument('--captures',type=Path,required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
n={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','p14':'http://schemas.microsoft.com/office/powerpoint/2010/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'};checks={};shas={};blob_wave=None
with zipfile.ZipFile(a.deck) as z:
 for slide,name,file in [(1,'trim-video','original-clock-12s.mp4'),(2,'fade-audio','original-scale-12s.wav'),(3,'bookmark-video','original-clock-12s.mp4')]:
  root=E.fromstring(z.read(f'ppt/slides/slide{slide}.xml'));rels={x.get('Id'):posixpath.normpath(posixpath.join('ppt/slides',x.get('Target'))) for x in E.fromstring(z.read(f'ppt/slides/_rels/slide{slide}.xml.rels'))};pic=next(x for x in root.findall('.//p:pic',n) if x.find('p:nvPicPr/p:cNvPr',n).get('name')==name);media=pic.find('.//p14:media',n);blob=z.read(rels[media.get('{'+n['r']+'}embed')]);checks[f'slide{slide}FullOriginalMediaRetained']=blob==(a.assets/file).read_bytes();shas[name]=hashlib.sha256(blob).hexdigest()
  volumes=[x.get('vol') for x in root.findall('.//p:cMediaNode',n)];checks[f'slide{slide}Volume80']=bool(volumes) and all(x=='80000' for x in volumes)
  if slide in [1,2]:checks[f'slide{slide}TrimCutAmounts']=media.find('p14:trim',n).attrib=={'st':'2000','end':'2000'}
  if slide==2:checks['nativeFade2000ms']=media.find('p14:fade',n).attrib=={'in':'2000','out':'2000'};blob_wave=blob
  if slide==3:checks['nativeBookmarks6000And9000']=[(x.get('name'),x.get('time')) for x in media.findall('p14:bmkLst/p14:bmk',n)]==[('source-six','6000'),('source-nine','9000')]
with wave.open(io.BytesIO(blob_wave),'rb') as w:
 rate=w.getframerate();samples=array.array('h',w.readframes(w.getnframes()));levels=[math.sqrt(sum(v*v for v in samples[round((i+0.2)*rate):round((i+0.8)*rate)])/round(0.6*rate)) for i in range(12)];median=statistics.median(levels);checks['sourceWaveNotPrefaded']=max(abs(x/median-1) for x in levels)<0.01
before=json.loads((a.native/'before.json').read_text(encoding='utf-8-sig'));checks['cliTrimSurvivesNativeBeforePolish']=all(x['start']==2000 and x['end']==10000 and x['length']==12000 for x in before[:2])
records=[json.loads((a.captures/name/'recording.json').read_text(encoding='utf-8-sig')) for name in ['trim','fade','bookmarks']];checks['samePowerPointProcess']=len({x['targetPid'] for x in records})==1;checks['liveAudioNoReportedGaps']=all(x['audioDiscontinuities']==0 and x['audioTimestampErrors']==0 and x['audioBytes']>0 for x in records);checks['bookmarksActuallyNextNextPrevious']=[x['after'] for x in records[2]['events'][:3]]==[6000,9000,6000]
result={'passed':all(checks.values()),'checks':checks,'embeddedMediaSHA256':shas,'note':'Requires separate live PCM bounds/matched-fade reports and visual review of actual captured video; these structure/event checks alone do not prove audible fade or visible seek.'};a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2))
if not result['passed']:raise SystemExit(1)
