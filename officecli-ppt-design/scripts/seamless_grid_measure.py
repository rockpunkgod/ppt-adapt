"""Actual-PTS Fourier phase verifies continuous spatial scrolling across repeats."""
import argparse,json,subprocess
from pathlib import Path
import numpy as np
p=argparse.ArgumentParser();p.add_argument('--movie',required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--ffprobe',required=True);p.add_argument('--report',required=True);a=p.parse_args();out=Path(a.report)
if out.exists():p.error('Fresh report required')
info=json.loads(subprocess.check_output([a.ffprobe,'-v','error','-select_streams','v:0','-show_frames','-show_entries','stream=width,height,avg_frame_rate:frame=best_effort_timestamp_time','-of','json',a.movie],encoding='utf-8'))
w,h=info['streams'][0]['width'],info['streams'][0]['height'];pts=np.array([float(f['best_effort_timestamp_time']) for f in info['frames']]);assert np.all(np.diff(pts)>0)
height=min(h,w*9/16);offset=(h-height)/2;top=round(offset+height*.5);bottom=round(offset+height*.9)
proc=subprocess.Popen([a.ffmpeg,'-v','error','-i',a.movie,'-fps_mode','passthrough','-pix_fmt','rgb24','-f','rawvideo','pipe:1'],stdout=subprocess.PIPE)
phase=[];amp=[];count=0
try:
 while True:
  raw=proc.stdout.read(w*h*3)
  if not raw:break
  assert len(raw)==w*h*3
  frame=np.frombuffer(raw,dtype=np.uint8).reshape(h,w,3)
  profile=frame[top:bottom].astype(float).mean(axis=(0,2));coefficient=np.fft.rfft(profile-profile.mean())[15]
  phase.append(float(np.angle(coefficient)));amp.append(float(abs(coefficient)/w));count+=1
finally:proc.stdout.close();returncode=proc.wait()
assert returncode==0 and count==len(pts)
phase=np.unwrap(phase);amp=np.array(amp);expected=-w/17
fits=[]
for start,end in [(2,17),(17,34),(34,51),(2,53)]:
 mask=(pts>=start)&(pts<end);x=pts[mask];y=phase[mask]
 slope,intercept=np.polyfit(x,y,1);residual=y-(slope*x+intercept)
 velocity=-slope*w/(2*np.pi*15);error=abs(velocity-expected)/abs(expected)
 r2=1-float(np.square(residual).sum()/np.square(y-y.mean()).sum())
 fits.append({'seconds':[start,end],'samples':len(x),'velocityPixelsPerSecond':float(velocity),'expected':expected,'relativeSpeedError':float(error),'rSquared':r2,'phaseRMS':float(np.sqrt(np.square(residual).mean())),'passed':bool(len(x)>100 and velocity<0 and error<.02 and r2>.998 and np.sqrt(np.square(residual).mean())<.12)})
result={'passed':all(f['passed'] for f in fits) and pts[-1]>53 and float(np.median(amp[pts>2]))>.08,'fits':fits,'actualFrames':count,'actualPTS':pts.tolist(),'phaseRadians':phase.tolist(),'amplitude':amp.tolist(),'stream':info['streams'][0],'limit':'Fixed15-column native grid. Fourier phase is ambiguous by one cell; continuous unwrapped motion and late-window speed verify no stall/reset discontinuity, not arbitrary texture tracking.'}
out.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps({'passed':result['passed'],'frames':count,'fits':fits}))
