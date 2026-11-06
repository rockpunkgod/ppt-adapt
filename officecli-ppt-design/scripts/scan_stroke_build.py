import argparse,json,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--extraction',type=Path,required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);deck=a.outdir/'scan-strokes.pptx'
def cli(*xs):
 r=subprocess.run([a.officecli,*map(str,xs)],capture_output=True,text=True,encoding='utf-8')
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
cli('create',deck);cli('open',deck)
try:
 cli('set',deck,'/','--prop','slideWidth=40cm','--prop','slideHeight=22.5cm');ops=[]
 for i,(name,bg,title) in enumerate([('original.png','E6E1D7','原创模拟扫描诊断：不是实物扫描或字体描边'),('extracted.png','162A38','透明提取：深色底，灰边与淡笔触保留'),('extracted.png','D7AD59','透明提取：暖色底，无白纸矩形残留'),('extracted.png','718E79','透明提取：绿色底，同一RGBA原字节')],1):
  ops.extend([dict(command='add',parent='/',type='slide',props=dict(layout='blank',background=bg)),dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name='title',text=title,x='1cm',y='0.5cm',width='38cm',height='1.6cm',font='Microsoft YaHei',size=22,color='FFFFFF' if i!=1 else '283C42',fill='none',line='none')),dict(command='add',parent=f'/slide[{i}]',type='picture',props=dict(name='original-scan' if i==1 else 'true-alpha-strokes',src=str((a.extraction/name).resolve()),x='4cm',y='3cm',width='32cm',height='18cm')),dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name='scope',text='均匀纸色/单墨色校准 · 无二值阈值 · 保原图/参数 · 不外推任意拍照纸纹',x='1cm',y='21.3cm',width='38cm',height='0.9cm',font='Microsoft YaHei',size=14,color='FFFFFF' if i!=1 else '283C42',fill='none',line='none'))])
 f=a.outdir/'batch.json';f.write_text(json.dumps(ops,ensure_ascii=False),encoding='utf-8');cli('batch',deck,'--input',f,'--json')
finally:cli('close',deck)
print(cli('validate',deck,'--json'))
