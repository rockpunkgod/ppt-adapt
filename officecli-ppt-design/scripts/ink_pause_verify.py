"""Audit genuine InkML, preserved Replay behavior, bookmark group and original PCM clock."""
import argparse,io,json,posixpath,re,wave,zipfile
from pathlib import Path
from xml.etree import ElementTree as E
p=argparse.ArgumentParser();p.add_argument('deck',type=Path);p.add_argument('--before',type=Path);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
n={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships','p14':'http://schemas.microsoft.com/office/powerpoint/2010/main','mc':'http://schemas.openxmlformats.org/markup-compatibility/2006','ink':'http://www.w3.org/2003/InkML','rel':'http://schemas.openxmlformats.org/package/2006/relationships'}
checks={}
def decode_integer_trace(text):
 # InkML 3.2: ! absolute, ' first difference, " second difference;
 # each channel retains its own encoding mode between comma-separated points.
 modes=['!']*3;previous=[None]*3;velocity=[None]*3;acceleration=[None]*3;points=[]
 token=re.compile(r'''([!'\"]?)\s*([+-]?\d+|\*)''')
 for raw in text.split(','):
  parts=token.findall(raw)
  if len(parts)!=3 or token.sub('',raw).strip():raise ValueError('Expected three regular integer InkML channels')
  point=[]
  for j,(prefix,value) in enumerate(parts):
   if prefix:modes[j]=prefix
   if modes[j]=='!':
    if value=='*' and previous[j] is None:raise ValueError('Initial wildcard has no channel state')
    absolute=previous[j] if value=='*' else int(value);velocity[j]=None;acceleration[j]=None
   elif modes[j]=="'":
    if previous[j] is None:raise ValueError('Difference without explicit point')
    delta=velocity[j] if value=='*' else int(value)
    if delta is None:raise ValueError('Difference wildcard without velocity')
    absolute=previous[j]+delta;velocity[j]=delta;acceleration[j]=None
   else:
    if previous[j] is None or velocity[j] is None:raise ValueError('Second difference without first difference')
    delta2=acceleration[j] if value=='*' else int(value)
    if delta2 is None:raise ValueError('Second-difference wildcard without acceleration')
    velocity[j]+=delta2;absolute=previous[j]+velocity[j];acceleration[j]=delta2
   previous[j]=absolute;point.append(absolute)
  points.append(point)
 return points
with zipfile.ZipFile(a.deck) as z:
 slide=E.fromstring(z.read('ppt/slides/slide1.xml'));rels=E.fromstring(z.read('ppt/slides/_rels/slide1.xml.rels'));targets={x.get('Id'):posixpath.normpath(posixpath.join('ppt/slides',x.get('Target'))) for x in rels}
 part=next(x for x in slide.findall('.//p:contentPart',n) if x.find('p14:nvContentPartPr/p14:cNvPr',n).get('name')=='original-native-ink');ink_id=part.find('p14:nvContentPartPr/p14:cNvPr',n).get('id')
 ink=E.fromstring(z.read(targets[part.get('{'+n['r']+'}id')]))
 traces=ink.findall('ink:trace',n);packets=[point for t in traces for point in decode_integer_trace(t.text)]
 if a.before:
  with zipfile.ZipFile(a.before) as before:
   bs=E.fromstring(before.read('ppt/slides/slide1.xml'));br=E.fromstring(before.read('ppt/slides/_rels/slide1.xml.rels'));bt={x.get('Id'):posixpath.normpath(posixpath.join('ppt/slides',x.get('Target'))) for x in br}
   bp=next(x for x in bs.findall('.//p:contentPart',n) if x.find('p14:nvContentPartPr/p14:cNvPr',n).get('name')=='original-native-ink');bi=E.fromstring(before.read(bt[bp.get('{'+n['r']+'}id')]))
   expected=[point for t in bi.findall('ink:trace',n) for point in decode_integer_trace(t.text)]
  checks['nativeRoundtripPreservesEveryXYT']=packets==expected
 checks['genuineInkMLThreeTraces']=ink.tag=='{'+n['ink']+'}ink' and len(traces)==3 and len(packets)==183 and all(len(x)==3 for x in packets)
 checks['millisecondTimingMonotonic']=all(x[2]<=y[2] for x,y in zip(packets,packets[1:])) and packets[0][2]==0 and packets[-1][2]==5000
 timing=slide.find('mc:AlternateContent/mc:Choice/p:timing',n)
 if timing is None:timing=slide.find('p:timing',n)
 if timing is None:raise ValueError('Native timing missing')
 seqs=[x for x in timing.findall('.//p:cTn',n) if x.get('nodeType')=='interactiveSeq'];checks['oneInteractiveSequence']=len(seqs)==1
 seq=seqs[0];replays=[x for x in seq.findall('.//p:cTn',n) if x.get('presetID')=='63'];checks['oneReplayPreset']=len(replays)==1
 anims=[x for x in replays[0].findall('.//p:anim',n) if any(y.text=='drawProgress' for y in x.findall('.//p:attrName',n))]
 checks['nativeDrawProgressFiveSeconds']=len(anims)==1 and anims[0].find('p:cBhvr/p:cTn',n).get('dur')=='5000' and anims[0].find('p:cBhvr/p:tgtEl/p:spTgt',n).get('spid')==ink_id and [x.get('val') for x in anims[0].findall('p:tavLst/p:tav/p:val/p:fltVal',n)]==['0','1']
 checks['noDuplicateMainInk']=all(x.get('spid')!=ink_id for main in timing.findall('.//p:cTn',n) if main.get('nodeType')=='mainSeq' for x in main.findall('.//p:spTgt',n))
 audio=next(x for x in slide.findall('.//p:pic',n) if x.find('p:nvPicPr/p:cNvPr',n).get('name')=='ink-clock');audio_id=audio.find('p:nvPicPr/p:cNvPr',n).get('id')
 bmks=audio.findall('.//p14:bmk',n);checks['bookmarkAtStart']=any(x.get('name')=='InkCycle' and x.get('time')=='0' for x in bmks)
 triggers=seq.findall('.//p14:bmkTgt',n);checks['bookmarkTargetsClock']=bool(triggers) and all(x.get('spid')==audio_id and x.get('bmkName')=='InkCycle' for x in triggers)
 media=audio.find('.//p14:media',n);wav=z.read(targets[media.get('{'+n['r']+'}embed')])
 with wave.open(io.BytesIO(wav),'rb') as w:
  params=[w.getnchannels(),w.getsampwidth(),w.getframerate(),w.getnframes()];samples=w.readframes(w.getnframes())
 checks['originalSevenSecondPCM']=params==[1,2,16000,112000] and not any(samples)
 clocks=[x for x in timing.findall('.//p:audio',n) if x.find('p:cMediaNode/p:tgtEl/p:spTgt',n).get('spid')==audio_id];checks['clockLoops']=len(clocks)==1 and clocks[0].find('p:cMediaNode/p:cTn',n).get('repeatCount')=='indefinite'
 anchor=next(x for x in slide.findall('.//p:sp',n) if x.find('p:nvSpPr/p:cNvPr',n).get('name')=='ink-bookmark-anchor');off=anchor.find('p:spPr/a:xfrm/a:off',n);checks['anchorOffSlide']=int(off.get('x'))<0 and int(off.get('y'))<0
result={'passed':all(checks.values()),'checks':checks,'note':'Structure audit requires separate native playback evidence of progressive drawing, complete hold and reset. It cannot establish dynamic success by itself.'}
a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2))
if not result['passed']:raise SystemExit(1)
