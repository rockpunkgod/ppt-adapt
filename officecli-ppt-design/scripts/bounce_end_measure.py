"""Measure flat-color circle centroids frame-by-frame. NumPy is analysis-only."""
import argparse,subprocess,json
from pathlib import Path
import numpy as np
p=argparse.ArgumentParser();p.add_argument('--movie',required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--ffprobe',required=True);p.add_argument('--output',required=True);p.add_argument('--top-pad',type=int);a=p.parse_args();out=Path(a.output)
if out.exists():p.error('Refusing existing report')
metadata=json.loads(subprocess.check_output([a.ffprobe,'-v','error','-select_streams','v:0','-show_entries','stream=width,height,r_frame_rate,avg_frame_rate:frame=best_effort_timestamp_time','-of','json',a.movie],text=True))
stream=metadata['streams'][0]
if stream['width']!=1280 or stream['height']<720:raise ValueError('Requires 1280-wide fixed experiment video')
if a.top_pad is None:a.top_pad=(stream['height']-720)//2
timestamps=[float(v['best_effort_timestamp_time']) for v in metadata['frames']]
rows=[]
for name,x,color in [('control',240,[45,207,218]),('bounce',840,[249,163,74])]:
 proc=subprocess.Popen([a.ffmpeg,'-loglevel','error','-i',a.movie,'-vf',f'crop=220:720:{x}:{a.top_pad}','-fps_mode','passthrough','-f','rawvideo','-pix_fmt','rgb24','-'],stdout=subprocess.PIPE)
 i=0
 while True:
  b=proc.stdout.read(220*720*3)
  if not b:break
  if len(b)!=220*720*3:raise RuntimeError('Incomplete decoded frame')
  rgb=np.frombuffer(b,np.uint8).reshape(720,220,3).astype(np.int16);mask=(abs(rgb-np.array(color))<35).all(axis=2);weights=mask.sum(axis=1);area=int(weights.sum());ys=np.where(weights>0)[0]
  rows.append({'name':name,'frame':i,'seconds':timestamps[i],'area':area,'cy':float((weights*np.arange(720)).sum()/area) if area else None,'top':int(ys.min()) if area else None,'bottom':int(ys.max()) if area else None});i+=1
 if proc.wait()!=0:raise RuntimeError('FFmpeg decode failed')
summary={}
for name in ['control','bounce']:
 rr=[v for v in rows if v['name']==name];final=float(np.median([v['cy'] for v in rr if 4<v['seconds']<5 and v['cy'] is not None]));full=[v for v in rr if v['area']>15000 and v['top']>0 and v['bottom']<719 and 1.0<=v['seconds']<3.8]
 summary[name]={'finalCy':final,'maxCy':max(v['cy'] for v in full),'maxPositiveDeviationFromFinalPixels':max(v['cy'] for v in full)-final,'fullVisibleFrames':len(full)}
control={v['frame']:v for v in rows if v['name']=='control' and 1.0<=v['seconds']<3.8 and v['area']>15000 and v['top']>0 and v['bottom']<719}
diff=[abs(v['cy']-control[v['frame']]['cy']) for v in rows if v['name']=='bounce' and v['frame'] in control and v['area']>15000]
summary['maxPairedTrajectoryDifferencePixels']=max(diff)
report={'movie':str(Path(a.movie).resolve()),'frameMetadata':metadata['streams'][0],'uniqueRecordedFrames':len(timestamps),'analysisWindowSeconds':[1.0,3.8],'finalReferenceWindowSeconds':[4.0,5.0],'boundary':'Native 1280x720 fixed test layout; RGB mask threshold35, area>15000, top>0/bottom<719 for full-circle comparisons. Uses actual frame PTS; no decoder duplication. Partially clipped circles excluded from comparison. Not a general object tracker.','rows':rows,'summary':summary}
out.parent.mkdir(parents=True,exist_ok=True);out.write_text(json.dumps(report,indent=2),encoding='utf-8');print(json.dumps(summary,indent=2))
