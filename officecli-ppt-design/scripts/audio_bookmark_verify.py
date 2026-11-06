import argparse,json,zipfile,wave,io
from pathlib import Path
from xml.etree import ElementTree as E
p=argparse.ArgumentParser();p.add_argument('deck',type=Path);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
n={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','p14':'http://schemas.microsoft.com/office/powerpoint/2010/main'}
with zipfile.ZipFile(a.deck) as z:
 r=E.fromstring(z.read('ppt/slides/slide1.xml'))
 seq=r.find('.//p:cTn[@nodeType="interactiveSeq"]',n);assert seq is not None
 cond=seq.find('p:stCondLst/p:cond',n);assert cond.get('evt')=='onMediaBookmark'
 bmk=cond.find('.//p14:bmkTgt',n);assert bmk.get('bmkName')=='CycleStart'
 groups=seq.find('p:childTnLst',n);assert len(groups)==1,('separate bookmark-advancing groups',len(groups))
 effects=[q for q in seq.iter() if q.tag.endswith('}cTn') and q.get('presetClass') in ['entr','exit','path']]
 assert len(effects)==5
 timeline=[]
 for q in effects:
  delay=int(q.find('p:stCondLst/p:cond',n).get('delay'))
  durations=[int(c.get('dur')) for c in q.findall('.//p:cTn',n) if (c.get('dur') or '').isdigit()]
  timeline.append({'class':q.get('presetClass'),'delayMs':delay,'maxBehaviorDurationMs':max(durations)})
 assert [q['delayMs'] for q in timeline]==[0,120,200,600,1700]
 assert max(q['delayMs']+q['maxBehaviorDurationMs'] for q in timeline)==2000,timeline
 audio=r.find('.//p:audio/p:cMediaNode/p:cTn',n);assert audio.get('repeatCount')=='indefinite'
 wavs=[name for name in z.namelist() if name.startswith('ppt/media/') and name.endswith('.wav')];assert wavs
 with wave.open(io.BytesIO(z.read(wavs[0]))) as w:duration=w.getnframes()/w.getframerate();assert duration==2.5
 result={'passed':True,'oneInteractiveGroup':True,'bookmark':'CycleStart','audioSeconds':duration,'sequenceEndMs':2000,'holdMs':500,'effects':timeline,'playbackVerificationRequired':True}
a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2))
