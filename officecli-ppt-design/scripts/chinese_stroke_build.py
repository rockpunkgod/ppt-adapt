import argparse,json,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);deck=a.outdir/'chinese-strokes.pptx';A='http://schemas.openxmlformats.org/drawingml/2006/main'
def point(x,y):return f'<a:pt x="{round(x*1000000)}" y="{round(y*1000000)}"/>'
def path(commands):
 out=[f'<a:custGeom xmlns:a="{A}"><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="w" b="h"/><a:pathLst><a:path w="2000000" h="1000000">']
 for kind,pts in commands:out.append(f'<a:{kind}>'+''.join(point(x,y) for x,y in pts)+f'</a:{kind}>')
 out.append('<a:close/></a:path></a:pathLst></a:custGeom>');return ''.join(out)
def rr(x0,y0,x1,y1,rx,ry):
 k=0.5522847498
 return path([('moveTo',[(x0+rx,y0)]),('lnTo',[(x1-rx,y0)]),('cubicBezTo',[(x1-rx+k*rx,y0),(x1,y0+ry-k*ry),(x1,y0+ry)]),('lnTo',[(x1,y1-ry)]),('cubicBezTo',[(x1,y1-ry+k*ry),(x1-rx+k*rx,y1),(x1-rx,y1)]),('lnTo',[(x0+rx,y1)]),('cubicBezTo',[(x0+rx-k*rx,y1),(x0,y1-ry+k*ry),(x0,y1-ry)]),('lnTo',[(x0,y0+ry)]),('cubicBezTo',[(x0,y0+ry-k*ry),(x0+rx-k*rx,y0),(x0+rx,y0)])])
def vertical(x,y0=0.12,y1=0.88):return rr(x-0.0375,y0,x+0.0375,y1,0.0375,0.035)
def shared():
 x=0.85;r=0.0375
 return path([('moveTo',[(x-r,0.24)]),('cubicBezTo',[(x-r+0.01,0.45),(x-r-0.005,0.70),(x-r-0.03,0.86)]),('cubicBezTo',[(x-r-0.035,0.91),(x+r-0.03,0.91),(x+r-0.03,0.86)]),('cubicBezTo',[(x+r-0.005,0.70),(x+r+0.01,0.45),(x+r,0.24)]),('cubicBezTo',[(x+r,0.20),(x-r,0.20),(x-r,0.24)])])
def river():
 x=1.17;r=0.0375
 return path([('moveTo',[(x-r,0.18)]),('cubicBezTo',[(x-r-0.18,0.40),(x-r+0.16,0.53),(x-r,0.70)]),('cubicBezTo',[(x-r-0.03,0.78),(x-r-0.01,0.83),(x-r,0.86)]),('cubicBezTo',[(x-r,0.91),(x+r,0.91),(x+r,0.86)]),('cubicBezTo',[(x+r-0.01,0.83),(x+r-0.03,0.78),(x+r,0.70)]),('cubicBezTo',[(x+r+0.16,0.53),(x+r-0.18,0.40),(x+r,0.18)]),('cubicBezTo',[(x+r,0.13),(x-r,0.13),(x-r,0.18)])])
common=[('shan-left',vertical(0.15,0.35,0.88),['shan.left']),('shan-middle',vertical(0.5,0.12,0.88),['shan.middle']),('shan-bottom',rr(0.15,0.8525,0.85,0.9075,0.025,0.0275),['shan.bottom'])]
states={2:common+[('shan-right',vertical(0.85,0.35,0.88),['shan.right']),('chuan-left',vertical(1.30,0.20,0.88),['chuan.left']),('chuan-middle',vertical(1.62,0.18,0.88),['chuan.middle']),('chuan-right',vertical(1.94),['chuan.right'])],3:common+[('shared-stroke',shared(),['shan.right','chuan.left']),('chuan-middle',vertical(1.17,0.18,0.88),['chuan.middle']),('chuan-right',vertical(1.49),['chuan.right'])],4:common+[('shared-stroke',shared(),['shan.right','chuan.left']),('river-symbol',river(),['chuan.middle']),('chuan-right',vertical(1.49),['chuan.right'])]}
def cli(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8')
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
cli('create',deck);cli('open',deck);cli('set',deck,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm');ops=[];geometry=[];meta=[]
for s,title in enumerate(['原文字备份：山川','按笔画走势，手工建立闭合骨架','一笔共用：山的右竖，也是川的左竖','主题替换：以河流曲线替换川的中竖'],1):
 ops.append(dict(command='add',parent='/',type='slide',props=dict(layout='blank',background='F4F0E8')))
 ops.append(dict(command='add',parent=f'/slide[{s}]',type='shape',props=dict(name='title',text=title,x='1.5cm',y='0.8cm',width='30.87cm',height='1.6cm',font='Microsoft YaHei',size=25,color='172A35',fill='none',line='none')))
 if s==1:ops.append(dict(command='add',parent='/slide[1]',type='shape',props=dict(name='editable-original',text='山川',x='2cm',y='5cm',width='29.87cm',height='9cm',font='Microsoft YaHei',bold=True,size=180,color='172A35',fill='none',line='none',align='center',valign='middle')))
 else:
  x,y,w,h=(4.5,5,24,9) if s==2 else (5.66,4.3,27.5,10.5)
  for name,xml,roles in states[s]:
   color='287D75' if name in ['shared-stroke','river-symbol'] else '172A35'
   ops.append(dict(command='add',parent=f'/slide[{s}]',type='shape',props=dict(name=name,geometry='rect',x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',fill=color,line='none')))
   geometry.append(dict(command='raw-set',part=f'/slide[{s}]',xpath=f'//p:sp[p:nvSpPr/p:cNvPr[@name="{name}"]]/p:spPr/a:prstGeom',action='replace',xml=xml));meta.append(dict(slide=s,name=name,roles=roles,designer='original hand-selected stroke geometry',symbolReplaces='chuan.middle' if name=='river-symbol' else None))
  ops.append(dict(command='add',parent=f'/slide[{s}]',type='shape',props=dict(name='scope',text='原创字形实验 · 每个笔画均可编辑 · 识读与主题适配需人工判断',x='2cm',y='16.4cm',width='29.87cm',height='1cm',font='Microsoft YaHei',size=17,color='6C655B',fill='none',line='none',align='center')))
batch=a.outdir/'base-batch.json';batch.write_text(json.dumps(ops,ensure_ascii=False),encoding='utf-8');cli('batch',deck,'--input',batch,'--json');gb=a.outdir/'geometry-batch.json';gb.write_text(json.dumps(geometry),encoding='utf-8');cli('batch',deck,'--input',gb,'--json');cli('close',deck);print(cli('validate',deck,'--json'))
(a.outdir/'design-model.json').write_text(json.dumps({'text':'山川','fontBackup':'Microsoft YaHei Bold','method':'Human-authored closed stroke geometry; no font outline conversion','sharedRoles':['shan.right','chuan.left'],'symbolMeaning':'river / outdoor landscape; original middle vertical is removed, not overlaid','states':meta,'limits':'Designed for this two-character word only, not a Chinese font generator; readability and semantic fitness require human review.'},ensure_ascii=False,indent=2),encoding='utf-8')
