"""Live retained-clip duration and source-pitch audit; output level flatness is not a trim criterion."""
import argparse,array,json,math,wave
from pathlib import Path
def analyze(path):
 with wave.open(str(path),'rb') as w:
  channels=w.getnchannels();rate=w.getframerate();assert w.getsampwidth()==2;mono=array.array('h',w.readframes(w.getnframes()))[::channels]
 step=round(rate*0.05);levels=[math.sqrt(sum(x*x for x in mono[i:i+step])/len(mono[i:i+step])) for i in range(0,len(mono),step) if mono[i:i+step]];peak=max(levels,default=0);active=[i for i,x in enumerate(levels) if x>peak*0.012 and x>2]
 if not active:return {'passed':False,'reason':'No non-silent PowerPoint signal captured'}
 onset=active[0]*0.05;end=(active[-1]+1)*0.05;duration=end-onset;freqs=[261.625565*2**(i/12) for i in range(12)]
 def note(t):
  segment=mono[round((onset+t)*rate):round((onset+t+0.30)*rate)];powers=[]
  for f in freqs:
   c=2*math.cos(2*math.pi*f/rate);v1=v2=0.0
   for x in segment:v=x+c*v1-v2;v2=v1;v1=v
   powers.append(max(0,v1*v1+v2*v2-c*v1*v2))
  return max(range(len(powers)),key=powers.__getitem__)
 first=note(0.25);last=note(max(0.3,duration-0.70));return {'passed':abs(duration-8)<0.25 and first==2 and last==9,'captureSeconds':len(mono)/rate,'signalOnsetSeconds':onset,'signalEndSeconds':end,'signalDurationSeconds':duration,'firstSourceNoteIndex':first,'lastSourceNoteIndex':last,'expectedSourceIntervalSeconds':[2,10],'captureFormat':[channels,2,rate],'note':'Actual process-scoped PCM,50ms analysis windows and thresholded onset/end; not sample-exact trimming. RMS across changing frequencies is not assumed flat.'}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('wav',type=Path);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
 if a.report.exists():raise FileExistsError(a.report)
 result=analyze(a.wav);a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2));raise SystemExit(0 if result['passed'] else 1)
