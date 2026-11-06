"""Supply actual native-rendered target posters using only OfficeCLI mutations.

Requires native exports in artifacts/zoom-lab-native/ from native-verify.ps1.
The hidden picture carriers own valid media relationships for zoom blipFill.
"""
import argparse,json,subprocess,shutil
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',default='officecli');p.add_argument('--artifacts',type=Path,required=True);p.add_argument('--out',type=Path,required=True);args=p.parse_args()
root=args.artifacts.resolve();file=args.out.resolve()
if file.exists():p.error('Refusing to overwrite '+str(file))
file.parent.mkdir(parents=True,exist_ok=True)
shutil.copy2(root/'zoom-lab.pptx',file)
def call(*a):
    r=subprocess.run([args.officecli,*map(str,a)],capture_output=True,text=True,encoding='utf8',errors='replace')
    if r.returncode or 'UNSUPPORTED' in r.stdout+r.stderr:raise RuntimeError(r.stdout+r.stderr)
    print(r.stdout[:300])
    return r.stdout
for i in range(1,4):
    name=f'poster-cache-{i}'
    call('add',file,'/slide[1]','--type','picture','--prop',f'name={name}','--prop',f'src={root/"zoom-lab-native"/f"slide-{i+1}.png"}','--prop','x=0cm','--prop','y=0cm','--prop','width=8cm','--prop','height=4.5cm','--prop',f'alt=Zoom target {i+1} poster')
    node=json.loads(call('get',file,f'/slide[1]/picture[@name={name}]','--json'))['data']['results'][0]
    rid=node['format']['relId']
    call('raw-set',file,'/slide[1]','--xpath',f"//p:pic/p:nvPicPr/p:cNvPr[@name='{name}']",'--action','setattr','--xml','hidden=1')
    xpath=f"//mc:AlternateContent[mc:Choice/p:graphicFrame/p:nvGraphicFramePr/p:cNvPr[@name='zoom-{i}']]//a:blip"
    call('raw-set',file,'/slide[1]','--xpath',xpath,'--action','setattr','--xml',f'r:embed={rid}')
call('close',file)
call('validate',file)
