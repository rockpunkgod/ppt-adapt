"""Native actual-frame gauge trajectory, 2s timing and held-end audit."""
import argparse,subprocess,json,importlib.util,hashlib
from pathlib import Path
import numpy as np
from PIL import Image
p=argparse.ArgumentParser();p.add_argument('--directory',required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--ffprobe',required=True);p.add_argument('--frames');a=p.parse_args();o=Path(a.directory);live=o/'live';frames=Path(a.frames) if a.frames else live/'actual-frames'
if frames.exists():raise FileExistsError('Refusing overwrite decoded frames')
frames.mkdir();helper=Path(__file__).with_name('gauge_sync_pixels.py');helper=helper if helper.exists() else Path(__file__).with_name('pixels.py');sp=importlib.util.spec_from_file_location('gaugePixels',helper);gp=importlib.util.module_from_spec(sp);sp.loader.exec_module(gp)
movie=live/'gauge-live.mp4';r=subprocess.run([a.ffprobe,'-v','error','-select_streams','v:0','-show_frames','-show_entries','frame=best_effort_timestamp_time','-of','json',str(movie)],capture_output=True,text=True);assert r.returncode==0;ts=[float(q['best_effort_timestamp_time']) for q in json.loads(r.stdout)['frames']]
subprocess.run([a.ffmpeg,'-hide_banner','-loglevel','error','-i',str(movie),'-fps_mode','passthrough',str(frames/'f-%05d.png')],check=True);files=sorted(frames.glob('f-*.png'));assert len(files)==len(ts)>150
events=json.loads((live/'clicks.json').read_text(encoding='utf-8-sig'));dispatch=next((v for v in events if v['kind']=='advance'),None);assert dispatch is None or (dispatch['slide'] in [1,2] and dispatch['clickIndex']==0);expected=[('before',1,0),('trigger',1,1),('end',1,1)]+([('advance',dispatch['slide'],0)] if dispatch else [])+[('next',2,0),('before',2,0),('trigger',2,1),('end',2,1)];assert [(v['kind'],v['slide'],v['clickIndex']) for v in events]==expected
values=[v['value'] for v in json.loads((o/'model.json').read_text())];theta=np.arange(0,180.01,0.25);xx=np.round(gp.cx-gp.radius*0.9*np.cos(np.radians(theta))).astype(int);yy=np.round(gp.cy-gp.radius*0.9*np.sin(np.radians(theta))).astype(int)
rows=[];initialDigit=None
for t,f in zip(ts,files):
 im=gp.normalize(Image.open(f));ar=np.asarray(im,dtype=float);r,g,b=[ar[:,:,i] for i in range(3)];coral=(r>180)&(g>60)&(g<160)&(b<140)&(r>g*1.4);ys,xs=np.where(coral);dist=np.hypot(xs-gp.cx,ys-gp.cy);far=dist>150;assert np.count_nonzero(far)>100
 needle=float(np.median(np.degrees(np.arctan2(-(ys[far]-gp.cy),-(xs[far]-gp.cx)))));teal=(r<100)&(g>130)&(b>100)&(g>r+60);angles=theta[teal[yy,xx]];arc=float(max(angles)) if len(angles) else 0
 crop=ar[529:605,777:830];digit=(crop[:,:,0]<100)&(crop[:,:,1]<120)&(crop[:,:,2]<150)
 if initialDigit is None:initialDigit=digit.copy()
 rows.append({'time':t,'needle':needle,'arc':arc,'digitMaskHash':hashlib.sha256(digit.tobytes()).hexdigest(),'digitPixels':int(digit.sum()),'digitDifference':int(np.count_nonzero(digit!=initialDigit))})
seenEnd=False;cut=None
for i,row in enumerate(rows):
 if row['needle']>values[0]*1.8-0.5:seenEnd=True
 if seenEnd and row['needle']<1:cut=i;break
assert cut is not None and cut>30;segments=[rows[:cut],rows[cut:]]
(live/'observations.json').write_text(json.dumps({'rows':rows,'cut':cut,'events':events},indent=2),encoding='utf-8')
# Align presentation click timestamps with actual first frame of the natural next
# page. Older research logs stamped end immediately before advance; exact scripts
# add an immediate advance event. This is a frame-boundary estimate, not2.000s metrology.
advance=next((v for v in events if v['kind']=='advance'),None)
if advance is None:advance=next(v for v in events if v['kind']=='end' and v['slide']==1)
offset=rows[cut]['time']-advance['time'];assert abs(offset)<0.2
triggers=[v for v in events if v['kind']=='trigger'];summary=[]
for i,(seq,value,event) in enumerate(zip(segments,values,triggers),1):
 start=event['time']+offset
 def closest(at):return min(seq,key=lambda q:abs(q['time']-at))
 pre=[q for q in seq if q['time']<start-0.1];assert pre and all(abs(q['needle'])<0.6 and q['arc']<0.6 for q in pre)
 middle=closest(start+1.5);assert value*1.8-middle['needle']>0.6 and value*1.8-middle['arc']>0.6,'Already final by1.5s'
 late=[q for q in seq if q['time']>=start+2.1];assert len(late)>6
 for q in late[::3]:
  idx=min(range(len(rows)),key=lambda j:abs(rows[j]['time']-q['time']));gp.analyze(Image.open(files[idx]),value)
 endState=gp.analyze(Image.open(live/f's{i}-end.png'),value);gp.analyze(Image.open(live/f's{i}-before.png'),0)
 early=[q for q in seq if start+0.05<q['time']<start+0.95];assert len(early)>8 and max(q['digitDifference'] for q in early)>300 and np.ptp([q['digitPixels'] for q in early])>80,'Fast scrolling missing'
 summary.append({'value':value,'estimatedTriggerPts':start,'notFinalAt1_5s':middle,'finalBy2_1s':True,'heldLateSamples':len(late[::3]),'endState':endState})
report={'passed':True,'actualPtsFrames':len(rows),'usedPassthrough':True,'clockOffsetEstimate':offset,'clockAlignment':'Natural page advance frame boundary; precision bounded by frame/dispatch','states':summary,'timingEvidence':'Native Timing2s/Delay1s separate from screen bounds1.5s not final,2.1s final and held','intermediateDigits':'Decorative scrolling, not live instantaneous value','heldEndVerified':True}
(live/'motion-audit.json').write_text(json.dumps(report,indent=2),encoding='utf-8');print(json.dumps(report,indent=2))
