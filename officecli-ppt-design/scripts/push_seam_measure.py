"""Register recorded native Push frames, then inspect the moving rail seam."""
import argparse,json,subprocess,hashlib
from pathlib import Path
import numpy as np
from PIL import Image
def seam(frame,y,x):
 roi=frame[y-8:y+9,x-5:x+6].astype(int)
 red=(roi[:,:,0]>130)&(roi[:,:,1]<115)&(roi[:,:,2]<145)&(roi[:,:,0]>roi[:,:,1]*1.3)
 missing=int((~red.any(axis=1)).sum())
 centers=[float(np.where(row)[0].mean()-5) for row in red if row.any()]
 return {'missingRows':missing,'maxCenterErrorPx':max(map(abs,centers),default=999)}
def run(movie,native,out,ff,fp):
 if out.exists():raise FileExistsError(out)
 out.mkdir(parents=True)
 subprocess.run([ff,'-hide_banner','-loglevel','error','-i',str(movie),'-fps_mode','passthrough',str(out/'f-%04d.png')],check=True)
 probe=json.loads(subprocess.check_output([fp,'-v','error','-select_streams','v:0','-show_frames','-show_entries','frame=best_effort_timestamp_time','-of','json',str(movie)],text=True))
 pts=[float(q['best_effort_timestamp_time']) for q in probe['frames']]
 files=sorted(out.glob('f-*.png'));assert len(files)==len(pts)
 first=np.array(Image.open(files[0]).convert('RGB'))
 bg=np.array([246,248,252]);rows=np.where(np.abs(first[:,5].astype(int)-bg).max(axis=1)<12)[0];inner_lo,inner_hi=int(rows.min()),int(rows.max())+1
 w=first.shape[1];h=round(w*9/16)
 assert abs((inner_hi-inner_lo)-h)<=4,'Unexpected canvas/letterbox geometry'
 # Compression can darken the edge row. Preserve the known 16:9 slide height
 # and center it on the independently observed plain-background interior.
 lo=round((inner_lo+inner_hi-h)/2);hi=lo+h;assert lo>=0 and hi<=first.shape[0]
 refs=[np.array(Image.open(native/f'page-{i}.png').convert('RGB').resize((w,h),Image.Resampling.LANCZOS)).astype(np.int16) for i in [1,2,3]]
 records=[];negative=None
 for i in range(0,len(files),3):
  t=pts[i]
  if not(1.5<t<3.8 or 4.5<t<6.8):continue
  pair=0 if t<4 else 1;before,after=refs[pair:pair+2]
  frame=np.array(Image.open(files[i]).convert('RGB'))[lo:hi];sample=frame[::4,::16].astype(np.int16)
  def cost(k):
   merged=np.concatenate((before[k:],after[:k]),axis=0)
   return float(np.abs(sample-merged[::4,::16]).mean())
  coarse=min(range(0,h+1,8),key=cost)
  k=min(range(max(0,coarse-8),min(h,coarse+8)+1),key=cost)
  if not(.1*h<k<.9*h):continue
  error=cost(k);assert error<3,'Frame does not match a translated native page pair'
  y=h-k;x=round(w*24/33.87);check=seam(frame,y,x)
  assert check['missingRows']==0 and check['maxCenterErrorPx']<=2
  records.append({'pts':t,'transition':pair+1,'pushShiftPx':k,'seamY':y,'templateMeanAbsError':error,**check})
  if negative is None:
   blank=frame.copy();blank[y-3:y+4,x-5:x+6]=bg
   negative=seam(blank,y,x);assert negative['missingRows']>=7
 for pair in [1,2]:
  group=[q for q in records if q['transition']==pair];assert len(group)>=6
  assert all(b['pushShiftPx']>=a['pushShiftPx'] for a,b in zip(group,group[1:]))
 result={'passed':True,'encodedFrames':len(files),'sampling':'Every third encoded frame at original ffprobe PTS, no interpolation. Duplicate recorded frames may remain.','crop':[lo,hi],'middleFrames':records,'inMemoryGapNegative':negative,'scope':'Two upward Push transitions on original fixed rail fixture; no physical-click or reverse-navigation claim.'}
 (out/'measure.json').write_text(json.dumps(result,indent=2),encoding='utf-8');print(out/'measure.json')
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('movie',type=Path);p.add_argument('--native',type=Path,required=True);p.add_argument('--out',type=Path,required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--ffprobe',required=True);a=p.parse_args();run(a.movie,a.native,a.out,a.ffmpeg,a.ffprobe)
