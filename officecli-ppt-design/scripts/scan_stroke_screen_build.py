import argparse,json,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--fixture',type=Path,required=True);p.add_argument('--screen-byte-compensation',type=int,default=0,choices=[0,1]);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);deck=a.outdir/'scan-screen.pptx'
def cli(*xs):
 r=subprocess.run([a.officecli,*map(str,xs)],capture_output=True,text=True,encoding='utf-8')
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
cli('create',deck);cli('open',deck)
try:
 cli('set',deck,'/','--prop','slideWidth=40cm','--prop','slideHeight=22.5cm');ops=[]
 for i,bg in enumerate(['E6E1D7','162A38','D7AD59','718E79'],1):
  title='原创模拟扫描：中性白纸诊断图，非实物扫描' if i==1 else '原生反相 → 滤色：固定底色合成，不是真alpha'
  ops.extend([dict(command='add',parent='/',type='slide',props=dict(layout='blank',background=bg)),dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name='title',text=title,x='1cm',y='0.5cm',width='38cm',height='1.6cm',font='Microsoft YaHei',size=22,color='FFFFFF' if i!=1 else '283C42',fill='none',line='none')),dict(command='add',parent=f'/slide[{i}]',type='picture',props=dict(name='simulated-neutral-scan',src=str((a.fixture/'neutral-probe.png').resolve()),x='4cm',y='3cm',width='32cm',height='18cm'))])
  if i>1:
   ops.append(dict(command='raw-set',part=f'/slide[{i}]',xpath='//p:pic/p:blipFill/a:blip',action='append',xml=f'<a:duotone xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="FFFFFF"/><a:srgbClr val="000000"/></a:duotone>'))
   screen=''.join(f'{min(255,int(bg[j:j+2],16)+a.screen_byte_compensation):02X}' for j in [0,2,4])
   ops.append(dict(command='raw-set',part=f'/slide[{i}]',xpath='//p:pic/p:blipFill/a:blip',action='append',xml=f'<a:fillOverlay xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" blend="screen"><a:solidFill><a:srgbClr val="{screen}"/></a:solidFill></a:fillOverlay>'))
 f=a.outdir/'batch.json';f.write_text(json.dumps(ops,ensure_ascii=False),encoding='utf-8');cli('batch',deck,'--input',f,'--json')
finally:cli('close',deck)
print(cli('validate',deck,'--json'))
