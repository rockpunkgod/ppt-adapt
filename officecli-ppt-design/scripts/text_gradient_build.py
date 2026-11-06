"""Editable glyph gradient; OfficeCLI L2 only, standard library."""
import argparse,json,subprocess,shutil
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():p.error('Fresh output directory required')
a.outdir.mkdir(parents=True);log=[];deck=a.outdir/'text-gradient.pptx'
def cli(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace');log.append({'args':list(map(str,args)),'code':r.returncode,'stdout':r.stdout,'stderr':r.stderr})
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
cli('create',deck,'--locale','en-US');cli('set',deck,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
for i,bg in enumerate(['101A2A','EFE8DA'],1):
 cli('add',deck,'/','--type','slide','--prop','layout=blank','--prop','background='+bg)
 cli('add',deck,f'/slide[{i}]','--type','shape','--prop','name=glyph-gradient','--prop','text=OPEN','--prop','x=2.9cm','--prop','y=5.7cm','--prop','width=28.1cm','--prop','height=7.65cm','--prop','font=Arial Black','--prop','size=170','--prop','bold=true','--prop','fill=none','--prop','line=none','--prop','margin=0cm','--prop','align=center','--prop','valign=middle','--prop','textFill=F26F4E@0-9E81EA@30-67CFD2@55-67CFD2+alpha0@85-67CFD2+alpha0@100-0')
cli('close',deck)
mask=a.outdir/'text-mask.pptx';shutil.copyfile(deck,mask)
for i in [1,2]:cli('set',mask,f'/slide[{i}]/shape[1]','--prop','textFill=FFFFFF')
cli('close',mask)
(a.outdir/'build-log.json').write_text(json.dumps(log,ensure_ascii=False,indent=2),encoding='utf-8')
print(deck)
