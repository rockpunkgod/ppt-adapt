"""OfficeCLI authors two real native charts with series/category build animations."""
import argparse,json,subprocess,zipfile
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);a=p.parse_args();out=Path(a.out)
if any(q.exists() for q in [out,*[out.with_suffix(x) for x in ('.batch.json','.build-log.json','.validation.json')]]):raise FileExistsError('Refusing overwrite')
out.parent.mkdir(parents=True,exist_ok=True)
def run(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8');
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
run('create',out);run('set',out,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
ops=[]
for i,title in enumerate(('Series: finish Alpha, then Beta','Category: compare both products each quarter'),1):
 ops.append(dict(command='add',parent='/',type='slide',props={'layout':'blank','background':'F6F8FC'}))
 def txt(name,text,x,y,w,h,size,color):
  ops.append(dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name=name,text=text,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',size=str(size),font='Arial',color=color,fill='none',line='none',margin='0cm')))
 txt('title',title,1.7,1.3,30.5,2,27,'17324E')
 txt('unit','Orders (demonstration data)',1.8,3.6,29,1,16,'527086')
 txt('guide','Click to reveal the next data group. Same chart; different explanation order.',1.8,17.5,30,1,13,'527086')
 ops.append(dict(command='add',parent=f'/slide[{i}]',type='chart',props=dict(name='sales-chart',chartType='column',categories='Q1,Q2,Q3,Q4',data='Alpha:24,38,46,60;Beta:36,29,54,42',x='1.8cm',y='4.8cm',width='30cm',height='12cm',title='none',legend='bottom',legendFont='16:17324E:Arial',axisfont='16:527086:Arial',axismin='0',axismax='70',majorunit='10',chartareafill='none',chartborder='none',plotareafill='none',plotborder='none',gridlines='DCE4EE:0.5',colors='007F86,E88735',overlap='0')))
 ops.append(dict(command='add',parent=f'/slide[{i}]/chart[1]',type='animation',props={'effect':'fade','class':'entrance','chartBuild':'series' if i==1 else 'category','trigger':'onClick','duration':'500'}))
 ops.append(dict(command='add',parent=f'/slide[{i}]',type='notes',props={'text':'BV1VTDGYsEmd P21 2181-2322s. Original synthetic data. Native chart build: series versus category. Native literal values; this animation recipe does not promise PowerPoint Edit Data. Static bars are not a substitute for native chart subelement targets.'}))
b=out.with_suffix('.batch.json');b.write_text(json.dumps(ops),encoding='utf-8');r=run('batch',out,'--input',b,'--json');out.with_suffix('.build-log.json').write_text(r,encoding='utf-8');assert json.loads(r)['success'];run('close',out)
with zipfile.ZipFile(out) as z:assert len([n for n in z.namelist() if '/charts/chart' in n and n.endswith('.xml')])==2
out.with_suffix('.validation.json').write_text(run('validate',out,'--json'),encoding='utf-8');print(out)
