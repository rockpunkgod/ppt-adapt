import argparse,json,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--assets',type=Path,required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);deck=a.outdir/'image-glitch.pptx'
def cli(*xs):
 r=subprocess.run([a.officecli,*map(str,xs)],capture_output=True,text=True,encoding='utf-8')
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
cli('create',deck);cli('open',deck)
try:
 cli('set',deck,'/','--prop','slideWidth=40cm','--prop','slideHeight=22.5cm');ops=[]
 for i,(name,title) in enumerate([('original.png','原创图像：彩色通道与半透明对照'),('channels.png','真实图像采样：R右移10px，GB左移10px'),('slices.png','选区副本：三段窄条分别错移'),('streak-approx.png','近似实验：自定义右向拖线，非Adobe Wind等价')],1):
  ops.extend([dict(command='add',parent='/',type='slide',props=dict(layout='blank',background='101C2A')),dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name='title',text=title,x='1.5cm',y='0.5cm',width='37cm',height='1.6cm',font='Microsoft YaHei',size=23,color='F4F0E8',fill='none',line='none')),dict(command='add',parent=f'/slide[{i}]',type='picture',props=dict(name='image-channel-diagnostic',src=str((a.assets/name).resolve()),x='4cm',y='3cm',width='32cm',height='18cm',alt=title)),dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name='scope',text='确定性图像处理 → OfficeCLI嵌入 · 原图/公式/透明度保留 · 静态样本',x='1.5cm',y='21.3cm',width='37cm',height='0.9cm',font='Microsoft YaHei',size=14,color='C0C9CD',fill='none',line='none'))])
 batch=a.outdir/'batch.json';batch.write_text(json.dumps(ops,ensure_ascii=False),encoding='utf-8');cli('batch',deck,'--input',batch,'--json')
finally:cli('close',deck)
print(cli('validate',deck,'--json'))
