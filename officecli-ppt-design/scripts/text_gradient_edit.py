"""Edit actual text after native roundtrip, OfficeCLI only; preserve font gradient."""
import argparse,json,shutil,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--input',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args()
if a.output.exists():p.error('Fresh output required')
a.output.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(a.input,a.output)
for i in [1,2]:
 r=subprocess.run([a.officecli,'set',str(a.output),f'/slide[{i}]/shape[1]','--prop','text=OPAL'],capture_output=True,text=True,encoding='utf-8',errors='replace')
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
r=subprocess.run([a.officecli,'close',str(a.output)],capture_output=True,text=True,encoding='utf-8',errors='replace')
if r.returncode:raise RuntimeError(r.stdout+r.stderr)
print(a.output)
