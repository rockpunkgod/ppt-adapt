"""Independent actual-frame PTS/trajectory audit, not a COM object-count proxy."""
import argparse,json,subprocess
from pathlib import Path
import numpy as np
from PIL import Image
p=argparse.ArgumentParser();p.add_argument('--live',required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--ffprobe',required=True);p.add_argument('--frames');a=p.parse_args();o=Path(a.live);movie=o/'master-loop-live.mp4';frames=Path(a.frames) if a.frames else o/'actual-frames'
if frames.exists():raise FileExistsError('Refusing overwrite frame evidence')
frames.mkdir()
r=subprocess.run([a.ffprobe,'-v','error','-select_streams','v:0','-show_frames','-show_entries','frame=best_effort_timestamp_time','-of','json',str(movie)],capture_output=True,text=True);assert r.returncode==0
times=[float(q['best_effort_timestamp_time']) for q in json.loads(r.stdout)['frames']]
subprocess.run([a.ffmpeg,'-hide_banner','-loglevel','error','-i',str(movie),'-fps_mode','passthrough',str(frames/'f-%05d.png')],check=True)
files=sorted(frames.glob('f-*.png'));assert len(files)==len(times) and len(files)>250
first=Image.open(files[0]);w,h=first.size
# Projection windows can be 16:10 with a letterboxed16:9 slide. Calibrate visible
# slide height from the known solid background; do not call black bars slide space.
cal=np.asarray(first.convert('RGB').resize((1600,900)));column=cal[:,1450,:];bg=(abs(column[:,0].astype(float)-16)<6)&(abs(column[:,1].astype(float)-34)<6)&(abs(column[:,2].astype(float)-54)<6);ys=np.where(bg)[0];canvasHeight=int(ys.max()-ys.min()+1);assert 650<canvasHeight<=900
def title(im):return np.asarray(im.resize((1600,900)).crop((80,60,1000,160))).mean(axis=2)>180
templates=[title(Image.open(o/f'page{i}-before.png')) for i in [1,2,3]]
def centroid(im,kind):
 if kind=='orange':x0,y0,x1,y1=40,160,600,350
 elif kind in ['teal','purple']:x0,y0,x1,y1=580,180,850,700
 else:x0,y0,x1,y1=1000,500,1450,760
 arr=np.asarray(im.crop((x0,y0,x1,y1)),dtype=float);r,g,b=[arr[:,:,k] for k in range(3)]
 if kind=='orange':m=(r>180)&(g>70)&(b<100)&(r>g*1.3)
 elif kind=='teal':m=(r<90)&(g>140)&(b>100)
 elif kind=='purple':m=(r>130)&(g<160)&(b>160)&(r>g+30)
 else:m=(b>120)&(b>g*1.3)&(r<100)
 ys,xs=np.where(m);return {'count':len(xs),'x':float(xs.mean()+x0) if len(xs)>50 else None,'y':float(ys.mean()+y0) if len(xs)>50 else None}
rows=[]
for t,f in zip(times,files):
 im=Image.open(f).convert('RGB').resize((1600,900));white=title(im);label=int(np.argmin([np.count_nonzero(white!=v) for v in templates]))+1
 rows.append(dict(time=t,slide=label,**{k:centroid(im,k) for k in ['orange','teal','purple','blue']}))
events=json.loads((o/'clicks.json').read_text(encoding='utf-8-sig'));assert [(v['kind'],v['slide'],v['clickIndex']) for v in events]==[('initial',1,0),('local-click',1,1),('next',2,0),('local-click',2,1),('next',3,0),('local-click',3,1),('explicit-reset-return',1,0)]
# Four actual visual segments: A1, A2, B3, explicit-reset A1.
runs=[]
for row in rows:
 if not runs or runs[-1]['slide']!=row['slide']:runs.append({'slide':row['slide'],'rows':[]})
 runs[-1]['rows'].append(row)
(o/'observations.json').write_text(json.dumps({'times':times,'movieSize':[w,h],'normalizedSlideHeight':canvasHeight,'rows':rows}),encoding='utf-8')
assert [v['slide'] for v in runs]==[1,2,3,1],[(v['slide'],len(v['rows'])) for v in runs]
assert all(len(v['rows'])>30 for v in runs)
def fit(seq,kind,axis,expectedAmplitude,initial=False):
 ts=np.array([q['time']-seq[0]['time'] for q in seq]);ys=np.array([q[kind][axis] for q in seq]);assert not np.any(np.equal(ys,None))
 ys=ys.astype(float);best=None
 # Fixed tolerance: measured period within 2±0.05s; RMSE<6px; amplitude within3%.
 for period in np.arange(1.95,2.051,0.005):
  for phase in np.arange(0,period,0.01):
   tri=1-np.abs((((ts-phase)%period)/(period/2))-1);X=np.column_stack([np.ones(len(ts)),tri]);co=np.linalg.lstsq(X,ys,rcond=None)[0];rmse=float(np.sqrt(np.mean((X@co-ys)**2)))
   if co[1]*expectedAmplitude<=0:continue # Remove the half-period/sign alias using the known path direction.
   if best is None or rmse<best['rmse']:best=dict(period=float(period),phase=float(phase),base=float(co[0]),amplitude=float(co[1]),rmse=rmse)
 assert abs(best['amplitude']/expectedAmplitude-1)<0.03,(kind,best,expectedAmplitude)
 assert best['rmse']<6,(kind,best)
 if not initial:assert min(best['phase'],best['period']-best['phase'])<0.2,('not reset near entry',kind,best)
 return best
summary=[]
for n,run in enumerate(runs):
 seq=run['rows'];kind='purple' if run['slide']==3 else 'teal';other='teal' if kind=='purple' else 'purple'
 assert all(q[kind]['count']>2000 and q[other]['count']<100 for q in seq),'Layout visibility leaked'
 master=fit(seq,'orange','x',320,initial=n==0);layout=fit(seq,kind,'y',(-0.15 if kind=='purple' else 0.15)*canvasHeight,initial=n==0)
 assert seq[0]['blue']['count']<100,'Local Fade did not reset hidden at entry'
 if n<3:assert any(q['blue']['count']>10000 for q in seq),'Physical local click did not reveal local object'
 else:assert all(q['blue']['count']<100 for q in seq),'Explicit reset retained completed local Fade'
 summary.append({'slide':run['slide'],'start':seq[0]['time'],'end':seq[-1]['time'],'master':master,'layout':layout})
first=runs[0]['rows'];assert first[-1]['time']-first[0]['time']>7
for start,end in [(0,2),(2,4),(4,6),(6.1,6.9)]:
 window=[v['orange']['x'] for v in first if start<=v['time']-first[0]['time']<=end];assert np.ptp(window)>70,'Stopped before page end'
report={'passed':True,'actualFrameCount':len(rows),'usedActualPts':True,'movieSize':[w,h],'normalizedSlideHeight':canvasHeight,'segments':summary,'firstPageBeyondThreeCycles':True,'naturalPhysicalNextSlides':[2,3],'return':'explicit GotoSlide(1,ResetSlide=true), not default backward behavior','rows':rows}
(o/'motion-audit.json').write_text(json.dumps(report,indent=2),encoding='utf-8');print(json.dumps({k:v for k,v in report.items() if k!='rows'},indent=2))
