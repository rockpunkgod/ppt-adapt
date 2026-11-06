"""Original procedural motion clip + matching first-frame poster; no stock assets."""
import argparse,subprocess,json
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--ffmpeg',required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():p.error('Fresh output directory required')
a.outdir.mkdir(parents=True);clip=a.outdir/'original-screen.mp4';poster=a.outdir/'first-frame.png';log=[]
for args in [['-f','lavfi','-i','testsrc2=size=960x540:rate=30:duration=6','-an','-c:v','libx264','-preset','fast','-pix_fmt','yuv420p',str(clip)],['-i',str(clip),'-frames:v','1',str(poster)]]:
 r=subprocess.run([a.ffmpeg,'-hide_banner','-loglevel','error',*args],capture_output=True,text=True,encoding='utf-8');log.append({'args':args,'code':r.returncode,'stderr':r.stderr})
 if r.returncode:raise RuntimeError(r.stderr)
(a.outdir/'asset-log.json').write_text(json.dumps(log,indent=2),encoding='utf-8')
