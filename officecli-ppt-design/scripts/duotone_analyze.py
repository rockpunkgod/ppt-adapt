import argparse,json,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--native',type=Path,required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
def decode(path):return subprocess.run([a.ffmpeg,'-hide_banner','-loglevel','error','-i',str(path),'-f','rawvideo','-pix_fmt','rgb24','-'],capture_output=True,check=True).stdout
imgs=[decode(a.native/f'slide-{i}.png') for i in [1,2,3]]
def rgb(data,x,y):
 j=(round(y)*1600+round(x))*3;return list(data[j:j+3])
left=1.5/33.87*1600;top=3.1/19.05*900;width=30.87/33.87*1600;height=15.435/19.05*900
rows=[];checks={};colors=[[32,16,68],[242,215,135]]
for state in [1,2]:
 c0,c1=colors if state==1 else colors[::-1];samples=[]
 for i in range(16):
  x=left+width*(i+0.5)/16;y=top+height*32/256;original=rgb(imgs[0],x,y);actual=rgb(imgs[state],x,y);g=original[0];expected=[round(v+(w-v)*g/255) for v,w in zip(c0,c1)]
  samples.append({'inputGray':g,'actual':actual,'encodedRGBLinearDiagnostic':expected,'diagnosticError':max(abs(x-y) for x,y in zip(actual,expected))})
 ramp=[rgb(imgs[state],left+width*(2*g+0.5)/512,top+height*96/256) for g in range(256)]
 checks[f's{state+1}blackEndpoint']=max(abs(v-w) for v,w in zip(samples[0]['actual'],c0))<=2
 checks[f's{state+1}whiteEndpoint']=max(abs(v-w) for v,w in zip(samples[-1]['actual'],c1))<=2
 checks[f's{state+1}allMidtonesBetweenEndpoints']=all(all(min(v,w)<q<max(v,w) for q,v,w in zip(s['actual'],c0,c1)) for s in samples[1:-1])
 checks[f's{state+1}continuousRampHas128Colors']=len(set(map(tuple,ramp)))>=128
 direction=1 if state==1 else -1;checks[f's{state+1}monotoneRamp']=all(all(direction*(ramp[j+1][c]-ramp[j][c])>=-1 for c in range(3)) for j in range(255))
 rows.append({'slide':state+1,'colors':[c0,c1],'steps':samples,'rampUniqueColors':len(set(map(tuple,ramp))),'encodedRGBLinearDiagnosticMaxError':max(s['diagnosticError'] for s in samples)})
result={'passed':all(checks.values()),'checks':checks,'states':rows,'note':'Endpoint/interior/continuity checks are independent of any assumed gamma model. Encoded-RGB linear errors are a diagnostic for this neutral-gray source only, not a universal color-luma formula. Measurements sample actual native original and effected PNGs at the same interior coordinates.'}
a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2))
