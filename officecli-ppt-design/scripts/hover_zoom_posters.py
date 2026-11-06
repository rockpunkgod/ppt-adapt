"""Embed exact native target posters; all deck mutations through OfficeCLI."""
import argparse,subprocess,json,shutil
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--input',type=Path,required=True);p.add_argument('--output',type=Path,required=True);p.add_argument('--native-dir',type=Path,required=True);a=p.parse_args()
if a.output.exists():p.error('Fresh output required')
a.output.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(a.input,a.output);log=[]
def cli(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8',errors='replace');log.append({'args':list(map(str,args)),'code':r.returncode,'stdout':r.stdout})
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
for s in range(1,4):
 for j in [1,2]:
  name=f'poster-cache-{j}';cli('add',a.output,f'/slide[{s}]','--type','picture','--prop','name='+name,'--prop','src='+str(a.native_dir/f'slide-{j+3}.png'),'--prop','x=0cm','--prop','y=0cm','--prop','width=10.3cm','--prop','height=5.79375cm')
  d=json.loads(cli('get',a.output,f'/slide[{s}]/picture[@name={name}]','--json'))['data']['results'][0];rid=d['format']['relId']
  cli('raw-set',a.output,f'/slide[{s}]','--xpath',f"//p:pic/p:nvPicPr/p:cNvPr[@name='{name}']",'--action','setattr','--xml','hidden=1')
  cli('raw-set',a.output,f'/slide[{s}]','--xpath',f"//mc:AlternateContent[mc:Choice/p:graphicFrame/p:nvGraphicFramePr/p:cNvPr[@name='!!zoom-{j}']]//a:blip",'--action','setattr','--xml','r:embed='+rid)
cli('close',a.output);a.output.with_suffix('.poster-log.json').write_text(json.dumps(log,indent=2),encoding='utf-8');print(a.output)
