import argparse,collections,json,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--native',type=Path,required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
checks={};rows=[]
for s in [2,3,4]:
 data=subprocess.run([a.ffmpeg,'-hide_banner','-loglevel','error','-i',str(a.native/f'slide-{s}.png'),'-vf','crop=1420:500:90:210','-f','rawvideo','-pix_fmt','rgb24','-'],capture_output=True,check=True).stdout;w,h=1420,500;mask=bytearray(1 if max(data[j:j+3])<90 else 0 for j in range(0,len(data),3))
 def components(value):
  seen=bytearray(w*h);result=[]
  for start in range(w*h):
   if seen[start] or mask[start]!=value:continue
   q=collections.deque([start]);seen[start]=1;area=0;border=False
   while q:
    k=q.popleft();y,x=divmod(k,w);area+=1;border=border or x in [0,w-1] or y in [0,h-1]
    for nxt in ([k-1] if x else [])+([k+1] if x<w-1 else [])+([k-w] if y else [])+([k+w] if y<h-1 else []):
     if not seen[nxt] and mask[nxt]==value:seen[nxt]=1;q.append(nxt)
   result.append({'area':area,'touchesBoundary':border})
  return result
 black=[c for c in components(1) if c['area']>100];holes=[c for c in components(0) if not c['touchesBoundary'] and c['area']>50]
 checks[f's{s}fourConnectedLetters']=len(black)==4;checks[f's{s}fourOpenCounters']=len(holes)==4;rows.append({'slide':s,'letterComponents':black,'enclosedCounters':holes})
result={'passed':all(checks.values()),'checks':checks,'rows':rows,'note':'Independent native PNG connected-component/counter audit at fixed dark-color threshold; B has2 counters,O1,D1. Readability is also inspected visually, not established by the component count alone.'}
a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2))
