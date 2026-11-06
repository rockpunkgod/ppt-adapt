import argparse,array,json,math,wave
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--evidence',type=Path,required=True);p.add_argument('--report',type=Path);a=p.parse_args()
report=a.report or a.evidence/'analysis.json'
if report.exists():raise FileExistsError(report)
info=json.loads((a.evidence/'recording.json').read_text(encoding='utf-8-sig'))
with wave.open(str(a.evidence/'powerpoint-audio.wav'),'rb') as w:
 rate=w.getframerate();channels=w.getnchannels();assert w.getsampwidth()==2
 mono=array.array('h',w.readframes(w.getnframes()))[::channels]
step=round(rate*0.05);rms=[math.sqrt(sum(x*x for x in mono[i:i+step])/len(mono[i:i+step])) for i in range(0,len(mono),step) if mono[i:i+step]]
peak=max(rms,default=0);active=[i for i,x in enumerate(rms) if x>max(2,peak*0.012)]
if not active:raise RuntimeError('No actual PCM signal')
onset=active[0]*0.05;end=(active[-1]+1)*0.05
events=[e for e in info['events'] if e['action']=='native-next'];first,stop=[e['absoluteSeconds']-info['audioReadySeconds'] for e in events]
def note(t):
 segment=mono[round((onset+t)*rate):round((onset+t+0.20)*rate)];powers=[]
 for i in range(12):
  f=261.625565*2**(i/12);c=2*math.cos(2*math.pi*f/rate);v1=v2=0.0
  for x in segment:v=x+c*v1-v2;v2=v1;v1=v
  powers.append(max(0,v1*v1+v2*v2-c*v1*v2))
 return max(range(12),key=powers.__getitem__)
probes=[t/2 for t in range(1,int((end-onset-0.3)*2))];sequence=[{'sourceElapsed':t,'observedNote':note(t),'expectedNote':int(t)} for t in probes if abs(t-round(t))>0.2]
interior_start=first+0.3;interior_end=min(first+1.5,stop-0.5)
between=rms[max(0,round(interior_start/0.05)):max(0,round(interior_end/0.05))]
checks={'nativeSlides':[e['slide'] for e in events]==[2,3],'signalContinuesOnSecondPage':bool(between) and all(x>max(2,peak*0.012) for x in between),'pitchClockDoesNotRestart':bool(sequence) and all(x['observedNote']==x['expectedNote'] for x in sequence),'stopsNearThirdPage':abs(end-stop)<0.45,'stopsBeforeNaturalEnd':end-onset<11.5,'captureNoDiscontinuities':info['audioDiscontinuities']==0 and info['audioTimestampErrors']==0}
result={'passed':all(checks.values()),'checks':checks,'signalOnset':onset,'signalEnd':end,'signalDuration':end-onset,'nativeNextRelativeToPCMReady':[first,stop],'stopDelta':end-stop,'secondPageInteriorWindow':[interior_start,interior_end],'pitchProbes':sequence,'note':'50ms RMS windows and PCM-ready clock approximation; COM Next validates lifecycle only, not physical click interaction. No loop-until-stopped claim. Initial false analysis sampled the stop boundary using a COM-return clock; preserved separately. Continuity now samples a fixed interior 1.2-second page-2 window, preserving the same RMS threshold; stop boundary is independently checked with 0.45s tolerance.'}
report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2))
