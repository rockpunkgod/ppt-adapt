"""Original diagnostic media: visible source time + one constant-level note per second."""
import argparse,json,math,struct,subprocess,wave
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--ffmpeg',required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);rate=48000;duration=12;frequencies=[261.625565*2**(i/12) for i in range(duration)];wav=a.outdir/'original-scale-12s.wav'
with wave.open(str(wav),'wb') as w:
 w.setnchannels(1);w.setsampwidth(2);w.setframerate(rate)
 for second,f in enumerate(frequencies):
  values=[]
  for i in range(rate):
   edge=min(1,i/(rate*0.004),(rate-1-i)/(rate*0.004))
   values.append(round(0.30*32767*max(0,edge)*math.sin(2*math.pi*f*i/rate)))
  w.writeframes(struct.pack('<'+'h'*len(values),*values))
font=r"C\:/Windows/Fonts/arial.ttf"
vf=r"drawtext=fontfile='"+font+r"':text='SOURCE %{pts\:hms}':fontsize=72:fontcolor=white:x=(w-tw)/2:y=260,drawtext=fontfile='"+font+r"':text='Original 12 second diagnostic clip':fontsize=32:fontcolor=0x6AAE9D:x=(w-tw)/2:y=390"
video=a.outdir/'original-clock-12s.mp4'
subprocess.run([a.ffmpeg,'-hide_banner','-loglevel','error','-f','lavfi','-i','color=c=0x14283F:s=1280x720:r=30:d=12','-i',str(wav),'-vf',vf,'-c:v','libx264','-preset','fast','-pix_fmt','yuv420p','-c:a','aac','-b:a','160k','-shortest',str(video)],check=True)
for t in [0,2]:subprocess.run([a.ffmpeg,'-hide_banner','-loglevel','error','-ss',str(t),'-i',str(video),'-frames:v','1',str(a.outdir/f'poster-{t}s.png')],check=True)
(a.outdir/'manifest.json').write_text(json.dumps({'creator':'Original synthetic media generated in this task; no tutorial media reused','durationSeconds':12,'sampleRate':rate,'sourceAmplitude':0.30,'boundaryClickRampSeconds':0.004,'powerpointFadeNotBakedIn':True,'notes':[{'sourceSecond':i,'frequencyHz':f} for i,f in enumerate(frequencies)]},indent=2),encoding='utf-8')
print(video)
