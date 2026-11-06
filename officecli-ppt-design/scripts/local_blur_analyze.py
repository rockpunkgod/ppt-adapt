import argparse,json,math,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--native',type=Path,required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
def decode(path):return subprocess.run([a.ffmpeg,'-hide_banner','-loglevel','error','-i',str(path),'-f','rawvideo','-pix_fmt','rgb24','-'],capture_output=True,check=True).stdout
imgs=[decode(a.native/f'slide-{i}.png') for i in [1,2,3,4]];W,H=1600,900
def rect_cm(x,y,w,h):return tuple(round(q) for q in (x/33.87*W,y/19.05*H,(x+w)/33.87*W,(y+h)/19.05*H))
def rms(a,b,box):
 x0,y0,x1,y1=box;total=n=0
 for y in range(y0,y1):
  aa=a[(y*W+x0)*3:(y*W+x1)*3];bb=b[(y*W+x0)*3:(y*W+x1)*3];total+=sum((v-w)**2 for v,w in zip(aa,bb));n+=len(aa)
 return math.sqrt(total/n)
def freq(data,box):
 x0,y0,x1,y1=box;out=[]
 for y in range(y0,y1):
  for x in range(x0,x1-1):out.append(abs(data[(y*W+x+1)*3+2]-data[(y*W+x)*3+2]))
 return sum(out)/len(out)
box=rect_cm(4,9,7,3);clear_freq=freq(imgs[0],box);blur_freq=freq(imgs[1],box)
rows=[];checks={'fullBlurSuppressesHighFrequency':blur_freq<clear_freq*0.3}
for s,positions in [(3,[(2,5),(19.87,5)]),(4,[(7,2),(17,7)])]:
 for i,(x,y) in enumerate(positions,1):
  region=(x+1,y+1,8,3) if s==4 and i==1 else (x+1,y+1,10,7)
  box=rect_cm(*region);match=rms(imgs[s-1],imgs[1],box);different=rms(imgs[s-1],imgs[0],box)
  rows.append({'slide':s,'panel':i,'analysisRegionCM':region,'laterPanelBorderExcluded':s==4 and i==1,'sameWorldBlurRMS':match,'clearDifferenceRMS':different});checks[f's{s}panel{i}SameWorldBlur']=match<2;checks[f's{s}panel{i}ActuallyChanged']=different>10
 # A lower-left corner is outside all panels in both placements.
 outside=rms(imgs[s-1],imgs[0],rect_cm(0.5,17,5,1));checks[f's{s}OutsideRemainsClear']=outside<0.5
# Horizontal black→white step: compare the width between 10% and 90% intensity.
def edge(data):
 y=round(87.5/540*H);line=[data[(y*W+x)*3] for x in range(600,1000)];indices10=[i for i,v in enumerate(line) if v>=26];indices90=[i for i,v in enumerate(line) if v>=230]
 return indices90[0]-indices10[0] if indices10 and indices90 else None
edge_clear=edge(imgs[0]);edge_blur=edge(imgs[1]);checks['blurBroadensEdge']=edge_clear is not None and edge_blur is not None and edge_blur>=edge_clear+5
result={'passed':all(checks.values()),'checks':checks,'frequencyBlueAdjDifference':{'clear':clear_freq,'blur':blur_freq,'ratio':blur_freq/clear_freq},'edge10To90Pixels':{'clear':edge_clear,'blur':edge_blur},'worldCoordinates':rows,'note':'Static native PNG evidence for the implementation identified in accompanying provenance. No Gaussian kernel identity, artistic-radius equivalence to other apps, or directional motion-blur claim. World-coordinate windows compare independent full-slide blur at identical coordinates. The moved first panel uses its upper-left interior to exclude the later second panel white border; the earlier full-interior false is retained. Original thresholds are unchanged.'}
a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2))
