import argparse,json,subprocess,math
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--assets',type=Path,required=True);p.add_argument('--native',type=Path,required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
def decode(f,fmt):return subprocess.run([a.ffmpeg,'-v','error','-i',str(f),'-f','rawvideo','-pix_fmt',fmt,'-'],capture_output=True,check=True).stdout
checks={};rows=[];bg=(16,28,42)
for n,name in enumerate(['original.png','channels.png','slices.png','streak-approx.png'],1):
 src=decode(a.assets/name,'rgba');native=decode(a.native/f'slide-{n}.png','rgb24');errors=[];semi=[]
 def sp(x,y):return tuple(src[(y*640+x)*4:(y*640+x+1)*4])
 for y in range(31,330,7):
  for x in range(31,609,7):
   v=sp(x,y);nb=[sp(x+dx,y+dy) for dy in [-1,0,1] for dx in [-1,0,1]]
   if any(max(q[c] for q in nb)-min(q[c] for q in nb)>8 for c in range(4)):continue
   alpha=v[3]/255;expected=[round(v[c]*alpha+bg[c]*(1-alpha)) for c in range(3)];j=((120+2*y+1)*1600+160+2*x+1)*3;actual=native[j:j+3];err=[abs(actual[c]-expected[c]) for c in range(3)];errors.extend(err)
   if v[3]==128:semi.extend(err)
 rms=math.sqrt(sum(e*e for e in errors)/len(errors));semirms=math.sqrt(sum(e*e for e in semi)/len(semi))
 # Image's transparent corner is actually rendered over the page background, not blackened/flattened.
 j=(130*1600+170)*3;corner=list(native[j:j+3]);checks[f's{n}SufficientStableSamples']=len(errors)//3>=500;checks[f's{n}NativeMatchesEmbeddedPixels']=rms<=3;checks[f's{n}SemiAlphaCompositesCorrectly']=semirms<=3;checks[f's{n}TransparentCornerUsesSlideBackground']=corner==list(bg)
 rows.append({'slide':n,'sampleCount':len(errors)//3,'rms':rms,'semiAlphaSampleCount':len(semi)//3,'semiAlphaRMS':semirms,'transparentCorner':corner})
r={'passed':all(checks.values()),'checks':checks,'measurements':rows,'scope':'Native PowerPoint still-image rendering of supplied channel/slice assets; excludes exact Adobe Wind equivalence or editable channel layers in PPT.'};a.report.write_text(json.dumps(r,indent=2),encoding='utf-8');print(json.dumps(r,indent=2))
