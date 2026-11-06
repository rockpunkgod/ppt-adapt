import argparse,json,subprocess,math
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--extraction',type=Path,required=True);p.add_argument('--native',type=Path,required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
def dec(f,fmt):return subprocess.run([a.ffmpeg,'-v','error','-i',str(f),'-f','rawvideo','-pix_fmt',fmt,'-'],capture_output=True,check=True).stdout
src=dec(a.extraction/'extracted.png','rgba');checks={};rows=[]
for n,bg in [(2,(22,42,56)),(3,(215,173,89)),(4,(113,142,121))]:
 d=dec(a.native/f'slide-{n}.png','rgb24');errors=[];faint=[];paper=[]
 def alpha(x,y):return src[(y*640+x)*4+3]
 for y in sorted(set(range(20,340,3))|{300}):
  for x in range(20,620,3):
   v=alpha(x,y);nb=[alpha(x+dx,y+dy) for dx in [-1,0,1] for dy in [-1,0,1]]
   if max(nb)-min(nb)>2:continue
   expect=[round(255*v/255+c*(1-v/255)) for c in bg];j=((120+2*y+1)*1600+160+2*x+1)*3;actual=d[j:j+3];err=[abs(actual[c]-expect[c]) for c in range(3)];errors.extend(err)
   if v==12:faint.extend(err)
   if v==0:paper.extend(err)
 rms=lambda v:math.sqrt(sum(x*x for x in v)/len(v))
 checks[f's{n}NativeAlphaRGBMatch']=rms(errors)<=2;checks[f's{n}FaintTraceVisibleAndAccurate']=len(faint)>=300 and rms(faint)<=2;checks[f's{n}NoPaperRectangle']=rms(paper)<=0.5;rows.append({'slide':n,'sampleCount':len(errors)//3,'rms':rms(errors),'faintSamples':len(faint)//3,'faintRMS':rms(faint),'paperRMS':rms(paper)})
r={'passed':all(checks.values()),'checks':checks,'measurements':rows,'limits':'Native composition against 3 static solid backgrounds; does not establish arbitrary scan extraction or source Adobe Screen equivalence.'};a.report.write_text(json.dumps(r,indent=2),encoding='utf-8');print(json.dumps(r,indent=2))
