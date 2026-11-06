import argparse,json,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--movie',type=Path,required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
raw=subprocess.run([a.ffmpeg,'-hide_banner','-loglevel','error','-i',str(a.movie),'-vf',"fps=30,crop='min(iw,ih*16/9)':'min(ih,iw*9/16)',scale=320:180",'-f','rawvideo','-pix_fmt','rgb24','-'],capture_output=True,check=True).stdout
n=320*180*3;rows=[]
for i in range(len(raw)//n):
 frame=raw[i*n:(i+1)*n];purple=white=sy=0
 for j in range(0,n,3):
  r,g,b=frame[j:j+3]
  if abs(r-117)<14 and abs(g-61)<14 and abs(b-130)<14:purple+=1
  if r>225 and g>225 and b>225:white+=1;sy+=(j//3)//320
 rows.append({'t':i/30,'purpleFraction':purple/(320*180),'whitePixels':white,'whiteCentroidY':sy/white if white else None})
initial=rows[5];mixed=[r for r in rows if 0.10<r['purpleFraction']<0.94 and r['whitePixels']>100];final=rows[-30:]
checks={'startsCovered':initial['purpleFraction']>0.94,'partialRevealWithForeground':len(mixed)>=2,'foregroundMovesUp':bool(mixed) and min(r['whiteCentroidY'] for r in mixed)<initial['whiteCentroidY']-12,'eventualFullReveal':all(r['purpleFraction']<0.002 for r in final),'visibleMovieChanges':len(set(raw[k*n:(k+1)*n] for k in range(max(0,len(rows)-60),len(rows),10)))>=4}
result={'passed':all(checks.values()),'checks':checks,'mixedFrames':len(mixed),'initial':initial,'mixedSamples':mixed[::max(1,len(mixed)//5)],'last':rows[-1],'frames':rows,'note':'Color/centroid audit is for the original solid-purple diagnostic prototype; central 16:9 slide is cropped from the slideshow desktop. Actual compound geometry and movie identity are independently audited. Changing clock and partial letter-hole reveal were also inspected visually.'}
a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps({k:v for k,v in result.items() if k!='frames'},indent=2))
