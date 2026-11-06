"""Same embedded WAV, same native playback chain, source-frequency-aligned live PCM comparison."""
import argparse,array,hashlib,json,math,posixpath,statistics,wave,zipfile
from pathlib import Path
from xml.etree import ElementTree as E
p=argparse.ArgumentParser();p.add_argument('--fade-deck',type=Path,required=True);p.add_argument('--control-deck',type=Path,required=True);p.add_argument('--fade-wave',type=Path,required=True);p.add_argument('--control-wave',type=Path,required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
n={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','p14':'http://schemas.microsoft.com/office/powerpoint/2010/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def media_identity(deck):
 with zipfile.ZipFile(deck) as z:
  root=E.fromstring(z.read('ppt/slides/slide2.xml'));rels={x.get('Id'):posixpath.normpath(posixpath.join('ppt/slides',x.get('Target'))) for x in E.fromstring(z.read('ppt/slides/_rels/slide2.xml.rels'))};media=root.find('.//p14:media',n);blob=z.read(rels[media.get('{'+n['r']+'}embed')]);fade=root.find('.//p14:fade',n)
  return {'mediaSHA256':hashlib.sha256(blob).hexdigest(),'trim':root.find('.//p14:trim',n).attrib,'fade':fade.attrib if fade is not None else {'in':'0','out':'0'},'volumes':[x.get('vol') for x in root.findall('.//p:cMediaNode',n)],'timing':E.tostring(root.find('p:timing',n),encoding='unicode')}
fid=media_identity(a.fade_deck);cid=media_identity(a.control_deck)
same_media=fid['mediaSHA256']==cid['mediaSHA256'] and fid['trim']==cid['trim']=={'st':'2000','end':'2000'} and fid['volumes']==cid['volumes']==['80000'] and fid['timing']==cid['timing']
fm=json.loads((a.fade_wave.parent/'recording.json').read_text(encoding='utf-8-sig'));cm=json.loads((a.control_wave.parent/'recording.json').read_text(encoding='utf-8-sig'))
same_chain=fm['targetPid']==cm['targetPid'] and fm['slide']==cm['slide']==2 and fm['mediaId']==cm['mediaId'] and bool(fm.get('recordScriptSHA256')) and fm.get('recordScriptSHA256')==cm.get('recordScriptSHA256') and bool(fm.get('captureHelperSHA256')) and fm.get('captureHelperSHA256')==cm.get('captureHelperSHA256') and all(m['audioDiscontinuities']==0 and m['audioTimestampErrors']==0 for m in [fm,cm])
def load(path):
 with wave.open(str(path),'rb') as w:
  fmt=[w.getnchannels(),w.getsampwidth(),w.getframerate()];assert fmt==[2,2,44100];samples=array.array('h',w.readframes(w.getnframes()))[::2]
 return samples,fmt
freqs={i:261.625565*2**(i/12) for i in range(2,10)}
def alignment(samples):
 fs=44100;window=round(fs*0.12);step=round(fs*0.05);rows=[]
 for start in range(0,len(samples)-window,step):
  seg=samples[start:start+window];energy=math.sqrt(sum(x*x for x in seg)/len(seg));powers={}
  for note,f in freqs.items():
   c=2*math.cos(2*math.pi*f/fs);v1=v2=0.0
   for x in seg:v=x+c*v1-v2;v2=v1;v1=v
   powers[note]=max(0,v1*v1+v2*v2-c*v1*v2)
  rows.append(((start+window/2)/fs,max(powers,key=powers.get),energy))
 peak=max(x[2] for x in rows);first={}
 for i in range(len(rows)-2):
  t,note,energy=rows[i]
  if note not in first and energy>peak*0.02 and all(rows[j][1]==note for j in range(i,i+3)):first[note]=t
 transitions={str(note):first[note]-0.025 for note in range(3,10) if note in first};offsets=[t-(int(note)-2) for note,t in transitions.items()]
 if len(offsets)<6:raise ValueError('Insufficient stable source-note transitions for alignment')
 offset=statistics.median(offsets);return offset,transitions,max(offsets)-min(offsets)
fade,ffmt=load(a.fade_wave);control,cfmt=load(a.control_wave);fo,ft,fj=alignment(fade);co,ct,cj=alignment(control)
def rms(samples,offset,t):
 seg=samples[max(0,round((offset+t-0.10)*44100)):round((offset+t+0.10)*44100)];return math.sqrt(sum(x*x for x in seg)/len(seg))
times=[0.25,0.75,1.25,1.75,2.5,3.5,4.5,5.5,6.25,6.75,7.25,7.75];raw=[rms(fade,fo,t)/rms(control,co,t) for t in times];scale=statistics.median(raw[4:8]);ratios=[x/scale for x in raw]
# These are the original fade criteria; no threshold widening.
fade_ok=ratios[0]<0.65 and ratios[0]<ratios[2]<ratios[3] and ratios[-1]<0.65 and ratios[-1]<ratios[-3]<ratios[-4] and all(0.85<x<1.15 for x in ratios[4:8])
conditions=same_media and same_chain and fid['fade']=={'in':'2000','out':'2000'} and cid['fade']=={'in':'0','out':'0'} and fj<0.15 and cj<0.15
result={'passed':conditions and fade_ok,'sameEmbeddedWaveTrimVolumeAndTiming':same_media,'samePowerPointProcessPlayerAndRecorderChain':same_chain,'embeddedWaveSHA256':fid['mediaSHA256'],'fadeFields':fid['fade'],'controlFadeFields':cid['fade'],'trim':fid['trim'],'volumes':fid['volumes'],'captureFormat':ffmt,'sourceClockOffsetSeconds':{'fade':fo,'control':co},'noteTransitionTimesSeconds':{'fade':ft,'control':ct},'sourceClockOffsetSpreadSeconds':{'fade':fj,'control':cj},'normalizedFadeOverMatchedControl':dict(zip(map(str,times),ratios)),'originalFadeCriteriaPassed':fade_ok,'note':'Comparison uses identical embedded WAV bytes and native trim/volume/timing plus recorded same-PID/recorder/helper provenance. Independent live captures are aligned by stable source-note transitions, not first audible threshold. Original uncalibrated false reports remain preserved. No linear fade-curve or sample-exact timing claim.'}
a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2))
if not result['passed']:raise SystemExit(1)
