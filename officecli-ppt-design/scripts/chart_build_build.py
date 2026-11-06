"""OfficeCLI authors native series/category builds; optional per-point Wipe builds."""
import argparse,json,subprocess,zipfile
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);p.add_argument('--elements',action='store_true',help='Per-point seriesEl/categoryEl builds with Wipe from bottom, 500ms; default stays grouped Fade');a=p.parse_args();out=Path(a.out)
if any(q.exists() for q in [out,*[out.with_suffix(x) for x in ('.batch.json','.build-log.json','.validation.json')]]):raise FileExistsError('Refusing overwrite')
out.parent.mkdir(parents=True,exist_ok=True)
def run(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8');
 if r.returncode:raise RuntimeError(r.stdout+r.stderr)
 return r.stdout
run('create',out);run('set',out,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
ops=[]
titles=('Series elements: one product, quarter by quarter','Category elements: one quarter, product by product') if a.elements else ('Series: finish Alpha, then Beta','Category: compare both products each quarter')
for i,title in enumerate(titles,1):
 ops.append(dict(command='add',parent='/',type='slide',props={'layout':'blank','background':'F6F8FC'}))
 def txt(name,text,x,y,w,h,size,color):
  ops.append(dict(command='add',parent=f'/slide[{i}]',type='shape',props=dict(name=name,text=text,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',size=str(size),font='Arial',color=color,fill='none',line='none',margin='0cm')))
 txt('title',title,1.7,1.3,30.5,2,27,'17324E')
 txt('unit','Orders (demonstration data)',1.8,3.6,29,1,16,'527086')
 txt('guide','Click to reveal the next data group. Same chart; different explanation order.',1.8,17.5,30,1,13,'527086')
 ops.append(dict(command='add',parent=f'/slide[{i}]',type='chart',props=dict(name='sales-chart',chartType='column',categories='Q1,Q2,Q3,Q4',data='Alpha:24,38,46,60;Beta:36,29,54,42',x='1.8cm',y='4.8cm',width='30cm',height='12cm',title='none',legend='bottom',legendFont='16:17324E:Arial',axisfont='16:527086:Arial',axismin='0',axismax='70',majorunit='10',chartareafill='none',chartborder='none',plotareafill='none',plotborder='none',gridlines='DCE4EE:0.5',colors='007F86,E88735',overlap='0')))
 animation={'effect':'wipe' if a.elements else 'fade','class':'entrance','chartBuild':('seriesEl' if i==1 else 'categoryEl') if a.elements else ('series' if i==1 else 'category'),'trigger':'onClick','duration':'500'}
 if a.elements:animation['direction']='down'
 ops.append(dict(command='add',parent=f'/slide[{i}]/chart[1]',type='animation',props=animation))
 ops.append(dict(command='add',parent=f'/slide[{i}]',type='notes',props={'text':'BV1VTDGYsEmd P21 2181-2322s. Original synthetic data. Native chart build: '+('seriesEl versus categoryEl; Wipe from bottom 500ms.' if a.elements else 'series versus category; Fade 500ms.')+' Native literal values; this animation recipe does not promise PowerPoint Edit Data. Static bars are not a substitute for native chart subelement targets.'}))
b=out.with_suffix('.batch.json');b.write_text(json.dumps(ops),encoding='utf-8');r=run('batch',out,'--input',b,'--json');out.with_suffix('.build-log.json').write_text(r,encoding='utf-8');assert json.loads(r)['success'];run('close',out)
if a.elements:
 # OfficeCLI 1.0.160 writes preset 20 + wipe(left), ignoring direction.
 # Native PowerPoint EffectType=22, Direction=3 saves subtype4 + wipe(down).
 # Playback confirms this reveals from bottom; wipe(up) reveals from top.
 # Apply that native representation only to this fresh per-point fixture.
 for i in [1,2]:
  run('raw-set',out,f'/slide[{i}]','--xpath',"//p:cTn[@presetID and @presetClass='entr']",'--action','setattr','--xml','presetID=22')
  run('raw-set',out,f'/slide[{i}]','--xpath',"//p:cTn[@presetID and @presetClass='entr']",'--action','setattr','--xml','presetSubtype=4')
  run('raw-set',out,f'/slide[{i}]','--xpath','//p:animEffect','--action','setattr','--xml','filter=wipe(down)')
 run('close',out)
with zipfile.ZipFile(out) as z:assert len([n for n in z.namelist() if '/charts/chart' in n and n.endswith('.xml')])==2
out.with_suffix('.validation.json').write_text(run('validate',out,'--json'),encoding='utf-8');print(out)
