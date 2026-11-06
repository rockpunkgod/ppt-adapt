import argparse,json,math,subprocess,statistics
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--evidence',type=Path,required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
meta=json.loads((a.evidence/'events.json').read_text(encoding='utf-8-sig'));events=meta['events']
raw=subprocess.run([a.ffmpeg,'-hide_banner','-loglevel','error','-i',str(a.evidence/'native.mp4'),'-vf',"fps=20,crop='min(iw,ih*16/9)':'min(ih,iw*9/16)',scale=640:360",'-f','rawvideo','-pix_fmt','rgb24','-'],capture_output=True,check=True).stdout
frame_size=640*360*3;centers=[(x/33.87*640,7.8/19.05*360) for x in [6.9,16.935,26.97]];rois=[(round(x-32),round(y-32),round(x+32),round(y+32)) for x,y in centers];frames=[]
for k in range(len(raw)//frame_size):
 f=raw[k*frame_size:(k+1)*frame_size];phases=[];summaries=[]
 for (cx,cy),(x0,y0,x1,y1) in zip(centers,rois):
  pts=[];roi=bytearray()
  for y in range(y0,y1):
   roi+=f[(y*640+x0)*3:(y*640+x1)*3]
   for x in range(x0,x1):
    r,g,b=f[(y*640+x)*3:(y*640+x)*3+3]
    if r<130 and g>140 and b>130:pts.append((x,y))
  if len(pts)>=3:
   x=statistics.mean(q[0] for q in pts);y=statistics.mean(q[1] for q in pts);phase=round(((math.atan2(y-cy,x-cx)+math.pi/2)%(2*math.pi))/(math.pi/6))%12
  else:phase=None
  phases.append(phase);summaries.append(bytes(roi))
 frames.append({'t':k/20,'phases':phases,'rois':summaries})
states=[];checks={}
def circularmean(vals):
 c=sum(math.cos(v*math.pi/1.2) for v in vals);s=sum(math.sin(v*math.pi/1.2) for v in vals);return (math.atan2(s,c)*1.2/math.pi)%2.4
for i,e in enumerate(events):
 start=e['clickElapsed']+1.0;end=min(events[i+1]['clickElapsed']-0.25,e['clickElapsed']+2.7) if i+1<len(events) else 19.8;idx=e['actualSlide']-2;rows=[f for f in frames if start<=f['t']<=end]
 phases=[f['phases'][idx] for f in rows if f['phases'][idx] is not None];others=[j for j in range(3) if j!=idx]
 # Codec noise is allowed only as a mean absolute RGB difference; visible phase detection is independent.
 noise=[]
 for j in others:
  anchor=rows[0]['rois'][j];noise.append(max(sum(abs(x-y) for x,y in zip(anchor,f['rois'][j]))/len(anchor) for f in rows))
 origins=[(f['t']-f['phases'][idx]*0.2)%2.4 for f in rows if f['phases'][idx] is not None]
 origin=circularmean(origins) if origins else None
 row={'name':e['name'],'slide':e['actualSlide'],'window':[start,end],'uniqueActivePhases':sorted(set(phases)),'activeDetectionFraction':len(phases)/len(rows),'inactiveMaxMeanRGBDifference':noise,'phaseOriginModulo2_4':origin,'originMinusClickModulo2_4':(origin-e['clickElapsed'])%2.4 if origin is not None else None,'samples':[{'t':f['t'],'phase':f['phases'][idx]} for f in rows]}
 states.append(row);checks[e['name']+'ActiveAnimates']=len(set(phases))>=5 and len(phases)/len(rows)>0.9;checks[e['name']+'OthersStatic']=all(n<1.0 for n in noise);checks[e['name']+'OnlySelectedBright']=all(f['phases'][j] is None for f in rows for j in others)
restarts=[]
for j in range(3):
 first,second=states[j],states[j+3];delta=(second['originMinusClickModulo2_4']-first['originMinusClickModulo2_4']+1.2)%2.4-1.2
 restarts.append({'icon':j+1,'phaseResetDifferenceSeconds':delta,'restartConsistentWithinOneFrame':abs(delta)<0.2,'note':'Revisit phase-origin relative to click, after excluding Morph. 200ms GIF frames; not a guarantee for all PowerPoint/GIF decoders.'})
checks['sixPhysicalLinks']=meta['passed'] and len(events)==6 and all(e['actualSlide']==e['expectedSlide'] for e in events)
last_phases=[x['phase'] for x in states[-1]['samples'] if x['phase'] is not None];checks['lastGifObservedLoop']=any(x==11 and y==0 for x,y in zip(last_phases,last_phases[1:]))
result={'passedStateAndMotion':all(checks.values()),'checks':checks,'states':states,'reentryPhase':restarts,'note':'Activity/static state and reentry phase are separate conclusions. 50ms analysis sampling/200ms GIF frames; encoder startup creates a common clock offset. Current-slide self-link, stop/restart of entire show and other players remain untested.'}
a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps({k:v for k,v in result.items() if k!='states'},indent=2))
