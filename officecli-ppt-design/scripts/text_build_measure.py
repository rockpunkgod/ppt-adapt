"""Measure real paragraph states and per-unit Swivel phases, without creating PPT."""
import argparse,json,subprocess,math
from pathlib import Path
import numpy as np
from PIL import Image
p=argparse.ArgumentParser();p.add_argument('--directory',required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--ffprobe',required=True);p.add_argument('--report',required=True);a=p.parse_args()
d=Path(a.directory);out=Path(a.report)
if out.exists():p.error('New report required')
data=json.loads((d/'report.json').read_text(encoding='utf-8-sig'));movie=d/'text-build-live.mp4'
meta=json.loads(subprocess.check_output([a.ffprobe,'-v','error','-select_streams','v:0','-show_frames','-show_entries','stream=width,height,avg_frame_rate:frame=best_effort_timestamp_time','-of','json',str(movie)],encoding='utf-8'))
w,h=meta['streams'][0]['width'],meta['streams'][0]['height'];pts=[float(f['best_effort_timestamp_time']) for f in meta['frames']]
assert all(y>x for x,y in zip(pts,pts[1:]))
sx=w/data['client'][0];sy=h/data['client'][1]
color=np.array([28,101,97],dtype=np.int16)
def load(name):return np.asarray(Image.open(d/(name+'.png')).convert('RGB').resize((w,h),Image.Resampling.BICUBIC))
finals={1:load('paragraph-3'),2:load('character-final'),3:load('word-final')}
def green(im):return np.max(np.abs(im.astype(np.int16)-color),axis=-1)<18
units={i:[] for i in (1,2,3)}
paragraph_rows=np.flatnonzero(green(finals[1]).any(axis=1))
paragraph_bands=np.split(paragraph_rows,np.flatnonzero(np.diff(paragraph_rows)>2)+1)
assert len(paragraph_bands)==3,'Expected three independently separated rendered paragraph lines'
for unit in data['unitBoxes']:
 x,y,bw,bh=unit['box'];box=(max(0,math.floor(x*sx)-2),max(0,math.floor(y*sy)-16),min(w,math.ceil((x+bw)*sx)+2),min(h,math.ceil((y+bh)*sy)+2))
 if unit['slide']==1:
  band=paragraph_bands[unit['unit']-1];box=(0,int(band[0]),w,int(band[-1])+1)
 l,t,r,b=box;mask=green(finals[unit['slide']][t:b,l:r])
 # Compare solid interior pixels, excluding resampling/chroma-subsampled edges.
 padded=np.pad(mask,1,constant_values=False)
 mask=mask & padded[:-2,1:-1] & padded[2:,1:-1] & padded[1:-1,:-2] & padded[1:-1,2:]
 assert mask.sum()>15,unit
 units[unit['slide']].append({'unit':unit['unit'],'text':unit['text'],'box':box,'mask':mask,'pixels':int(mask.sum())})
def ratios(im,page):
 return [float((green(im[t:b,l:r])&u['mask']).sum()/u['pixels']) for u in units[page] for l,t,r,b in [u['box']]]
paragraph=[]
for step,name in enumerate(['paragraph-initial','paragraph-1','paragraph-2','paragraph-3']):
 r=ratios(load(name),1);actual=[v>.75 for v in r];expected=[i<step for i in range(3)]
 paragraph.append({'step':step,'ratios':r,'actual':actual,'expected':expected,'passed':actual==expected and all(v<.05 for v in r[step:])})
headers={i:np.max(im[45:220,35:1190],axis=-1)<100 for i,im in finals.items()}
process=subprocess.Popen([a.ffmpeg,'-v','error','-i',str(movie),'-fps_mode','passthrough','-f','rawvideo','-pix_fmt','rgb24','pipe:1'],stdout=subprocess.PIPE)
rows=[];n=0
try:
 while True:
  raw=process.stdout.read(w*h*3)
  if not raw:break
  assert len(raw)==w*h*3
  im=np.frombuffer(raw,dtype=np.uint8).reshape(h,w,3);header=np.max(im[45:220,35:1190],axis=-1)<100
  distances={i:float(np.mean(header!=m)) for i,m in headers.items()};page=min(distances,key=distances.get)
  if distances[page]<.015:rows.append({'pts':pts[n],'page':page,'ratios':ratios(im,page)})
  n+=1
finally:process.stdout.close();code=process.wait()
assert code==0 and n==len(pts)
motion=[]
for page in (2,3):
 page_rows=[r for r in rows if r['page']==page]
 onsets=[next((r['pts'] for r in page_rows if r['ratios'][j]>.15),None) for j in range(len(units[page]))]
 partial=[r for r in page_rows if r['ratios'][0]>.2 and r['ratios'][-1]<.1]
 complete=[r for r in page_rows if min(r['ratios'])>.75]
 passed=bool(partial and complete and all(t is not None for t in onsets) and onsets[-1]-onsets[0]>.02)
 motion.append({'page':page,'units':[u['text'] for u in units[page]],'onsetPTS':onsets,'firstLastSeparation':None if None in onsets else onsets[-1]-onsets[0],'partialFrameCount':len(partial),'samplePartial':partial[:4],'completeFrameCount':len(complete),'bestMinimumRatio':max((min(r['ratios']) for r in page_rows),default=0),'lastPageSamples':page_rows[-4:],'passed':passed})
result={'passed':all(r['passed'] for r in paragraph+motion),'paragraph':paragraph,'motion':motion,'actualDecodedFrames':n,'actualPTS':pts,'stream':meta['streams'][0],'sampling':'Requested60fps, actual independent VFR frames and FFprobe PTS; no decoder duplication. Fixed sample unit boxes/masks, not a general OCR/motion tracker.'}
out.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps({'passed':result['passed'],'actualFrames':n,'motion':[{k:r[k] for k in ('page','firstLastSeparation','partialFrameCount','passed')} for r in motion]}))
