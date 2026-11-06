"""Pixel evidence for this original zigzag recipe, not a generic Ink validator."""
import argparse,json,math,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('movie',type=Path);p.add_argument('--ffmpeg',required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
probe=Path(a.ffmpeg).with_name('ffprobe.exe' if Path(a.ffmpeg).suffix.lower()=='.exe' else 'ffprobe')
meta=json.loads(subprocess.run([str(probe),'-v','error','-select_streams','v:0','-show_entries','stream=width,height','-of','json',str(a.movie)],capture_output=True,text=True,check=True).stdout)['streams'][0]
width,height=meta['width'],meta['height'];scale=min(width/33.87,height/19.05);ox=(width-33.87*scale)/2;oy=(height-19.05*scale)/2
roi=[max(0,math.floor((ox+6.75*scale)*640/width)-2),max(0,math.floor((oy+7.75*scale)*400/height)-2),min(640,math.ceil((ox+27.75*scale)*640/width)+3),min(400,math.ceil((oy+12.75*scale)*400/height)+3)]
r=subprocess.run([a.ffmpeg,'-hide_banner','-loglevel','error','-i',str(a.movie),'-vf','fps=30,scale=640:400','-f','rawvideo','-pix_fmt','rgb24','-'],capture_output=True,check=True)
size=640*400*3;counts=[]
for frame in range(len(r.stdout)//size):
 data=r.stdout[frame*size:(frame+1)*size];white=0
 for y in range(roi[1],roi[3]):
  for x in range(roi[0],roi[2]):
   rgb=data[(y*640+x)*3:(y*640+x)*3+3]
   white+=min(rgb)>155 and max(rgb)-min(rgb)<28
 counts.append(white)
span=max(counts)-min(counts);smooth=[sum(counts[max(0,i-1):i+2])/len(counts[max(0,i-1):i+2]) for i in range(len(counts))]
resets=[]
for i in range(1,len(counts)):
 if span>150 and counts[i-1]-counts[i]>span*0.4 and (not resets or i/30-resets[-1]>1):resets.append(round(i/30,3))
holds=[];start=None
for i,value in enumerate(smooth+[float('-inf')]):
 high=value>=max(counts)-span*0.04
 if high and start is None:start=i
 if not high and start is not None:
  if i-start>=15:holds.append([round(start/30,3),round(i/30,3)])
  start=None
intervals=[round(y-x,3) for x,y in zip(resets,resets[1:])]
rises=[]
for reset in resets:
 next_holds=[h for h in holds if h[0]>reset and h[0]<reset+7]
 if next_holds:rises.append(round(next_holds[0][0]-reset,3))
result={'passed':span>150 and len(resets)>=3 and all(abs(x-7)<0.2 for x in intervals) and sum(y-x>=1.5 for x,y in holds)>=2 and len(rises)>=2 and all(4.4<x<5.3 for x in rises),'captureSize':[width,height],'roiAt640x400':roi,'whitePixelRange':[min(counts),max(counts)],'whitePixelSpan':span,'resetTimesSeconds':resets,'resetIntervalsSeconds':intervals,'nearCompleteHoldsSeconds':holds,'riseToNearCompleteAfterResetSeconds':rises,'movieSeconds':len(counts)/30,'note':'Specific original Ink zigzag ROI with a centered 33.87x19.05cm slide; complements visual review. Movie zero is after slideshow Run. Near-complete threshold includes the final 4 percent of line growth, so its plateau begins slightly before the exact full drawing.'}
a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2))
