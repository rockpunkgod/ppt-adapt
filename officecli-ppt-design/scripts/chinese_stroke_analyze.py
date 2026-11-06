import argparse,json,subprocess,math
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--native',type=Path,required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
def decode(n):return subprocess.run([a.ffmpeg,'-v','error','-i',str(a.native/f'slide-{n}.png'),'-f','rawvideo','-pix_fmt','rgb24','-'],capture_output=True,check=True).stdout
imgs=[decode(n) for n in [2,3,4]];checks={};rows=[]
def components(data):
 mask=set()
 for y in range(190,730):
  for x in range(140,1490):
   j=(y*1600+x)*3;r,g,b=data[j:j+3]
   if r<80 and g<170 and b<170:mask.add(y*1600+x)
 sizes=[]
 while mask:
  q=[mask.pop()];size=0
  while q:
   v=q.pop();size+=1
   for w in [v-1,v+1,v-1600,v+1600]:
    if w in mask:mask.remove(w);q.append(w)
  if size>50:sizes.append(size)
 return sorted(sizes,reverse=True)
for i,data in zip([2,3,4],imgs):
 sizes=components(data);checks[f's{i}VisibleConnectedComponents']=len(sizes)==(4 if i==2 else 3);rows.append({'slide':i,'componentAreas':sizes})
# Replacement is isolated to the middle-river region; exclude title/footer and compare real rendered body pixels.
outside=[];inside=[]
for y in range(190,730):
 for x in range(140,1490):
  j=(y*1600+x)*3;diff=sum((imgs[1][j+c]-imgs[2][j+c])**2 for c in range(3))/3
  (inside if 955<=x<=1105 else outside).append(diff)
rmso=math.sqrt(sum(outside)/len(outside));rmsi=math.sqrt(sum(inside)/len(inside));checks['unchangedStrokesRenderIdentically']=rmso<0.1;checks['replacementVisiblyChangesMiddle']=rmsi>10
r={'passed':all(checks.values()),'checks':checks,'components':rows,'unchangedRegionRMS':rmso,'replacementRegionRMS':rmsi,'scope':'Native raster continuity/isolation; not automatic Chinese recognition. Human must inspect readable word and theme fit.'};a.report.write_text(json.dumps(r,indent=2),encoding='utf-8');print(json.dumps(r,indent=2))
