"""Change the arithmetic text via OfficeCLI; preserve old gate to test real rebinding."""
import argparse,shutil,subprocess,json,zipfile
from pathlib import Path
import xml.etree.ElementTree as E
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--input',type=Path,required=True);p.add_argument('--output',type=Path,required=True);p.add_argument('--addend',type=int,choices=[2,3,4],required=True);a=p.parse_args();logfile=a.output.with_suffix('.variant-log.json')
if a.output.exists() or logfile.exists():p.error('Fresh outputs required')
n={'a':'http://schemas.openxmlformats.org/drawingml/2006/main','p':'http://schemas.openxmlformats.org/presentationml/2006/main'}
ids={}
with zipfile.ZipFile(a.input) as z:
 for i,name in [(1,'question'),(2,'result-math')]:
  r=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'));ids[i]=next(x.get('id') for x in r.findall('.//p:cNvPr',n) if x.get('name')==name)
with zipfile.ZipFile(a.input) as z:
 r=E.fromstring(z.read('ppt/slides/slide1.xml'));next_shape=next(q for q in r.findall('.//p:sp',n) if q.find('p:nvSpPr/p:cNvPr',n).get('name')=='next');has_run_link=bool(next_shape.findall('p:txBody//a:hlinkClick',n))
a.output.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(a.input,a.output);logs=[]
for i,text in [(1,f'2 + {a.addend} = ?'),(2,f'2 + {a.addend} = {2+a.addend}')]:
 args=[a.officecli,'set',str(a.output),f'/slide[{i}]/shape[@id={ids[i]}]','--prop','text='+text,'--json'];r=subprocess.run(args,capture_output=True,text=True,encoding='utf-8',errors='replace');logs.append({'args':args,'returncode':r.returncode,'stdout':r.stdout})
 if r.returncode or json.loads(r.stdout).get('success') is False:raise RuntimeError(r.stdout+r.stderr)
if has_run_link:
 args=[a.officecli,'raw-set',str(a.output),'/slide[1]','--xpath',"//p:sp[p:nvSpPr/p:cNvPr[@name='next']]/p:txBody//a:hlinkClick",'--action','remove','--json'];r=subprocess.run(args,capture_output=True,text=True,encoding='utf-8',errors='replace');logs.append({'args':args,'returncode':r.returncode,'stdout':r.stdout})
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
r=subprocess.run([a.officecli,'close',str(a.output)],capture_output=True,text=True,encoding='utf-8',errors='replace')
if r.returncode:raise RuntimeError(r.stdout+r.stderr)
logfile.write_text(json.dumps(logs,indent=2,ensure_ascii=False),encoding='utf-8');print(a.output)
