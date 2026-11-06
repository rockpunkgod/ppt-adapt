import argparse,json,math
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as E
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
p=argparse.ArgumentParser();p.add_argument('--deck',required=True);p.add_argument('--measure',required=True);p.add_argument('--report',required=True);a=p.parse_args();out=Path(a.report)
if out.exists():p.error('Fresh report required')
checks=[]
def check(name,value):checks.append({'name':name,'passed':bool(value)})
with ZipFile(a.deck) as z:
 root=E.fromstring(z.read('ppt/slides/slide1.xml'));pres=E.fromstring(z.read('ppt/presentation.xml'));width=int(pres.find('p:sldSz',N).get('cx'))
 groups=root.findall('p:cSld/p:spTree/p:grpSp',N);check('one true grid group',len(groups)==1);g=groups[0]
 gid=g.find('p:nvGrpSpPr/p:cNvPr',N).get('id');members=g.findall('p:sp',N);check('forty editable shapes',len(members)==40)
 names={s.find('p:nvSpPr/p:cNvPr',N).get('name'):s for s in members}
 check('thirty column edges and nine row edges',len([n for n in names if n.startswith('column-edge-')])==30 and len([n for n in names if n.startswith('row-edge-')])==9)
 check('two-slide coordinate bounds',int(names['period-bounds'].find('p:spPr/a:xfrm/a:ext',N).get('cx'))==2*width)
 cols=[int(names[f'column-edge-{j}'].find('p:spPr/a:xfrm/a:off',N).get('x')) for j in range(30)]
 check('two identical spatial periods',all(abs((cols[j+15]-cols[j])-width)<=1 for j in range(15)))
 check('regular fifteen-column pitch',all(abs((b-a)-width/15)<1 for a,b in zip(cols,cols[1:])))
 check('actual twenty percent alpha',all(s.find('p:spPr/a:solidFill/a:srgbClr',N).get('val')=='FFFFFF' and s.find('p:spPr/a:solidFill/a:srgbClr/a:alpha',N).get('val')=='20000' for n,s in names.items() if n!='period-bounds'))
 check('all animation targets are the whole group',all(s.get('spid')==gid for s in root.findall('.//p:spTgt',N)))
 effects=[n for n in root.findall('.//p:cTn',N) if n.get('presetClass')];motion=[n for n in effects if n.get('presetClass')=='path'];fade=[n for n in effects if n.get('presetClass')=='entr']
 check('one fade and one motion',len(effects)==2 and len(motion)==len(fade)==1)
 check('motion repeats indefinitely',motion[0].get('repeatCount')=='indefinite')
 check('seventeen second motion','17000' in [n.get('dur') for n in motion[0].findall('.//p:cTn',N)])
 path=motion[0].find('.//p:animMotion',N);check('exact one-slide translation',path.get('path').split()==['M','0','0','L','-1','0','E'])
 check('no acceleration or autoreverse',all(float(n.get('accel','0'))==0 and float(n.get('decel','0'))==0 and n.get('autoRev','0')=='0' for n in motion[0].iter() if n.tag.endswith('cTn')))
 check('fade once for one second',fade[0].get('repeatCount') is None and '1000' in [n.get('dur') for n in fade[0].findall('.//p:cTn',N)])
 check('native vectors without raster',not root.findall('.//p:pic',N))
m=json.loads(Path(a.measure).read_text(encoding='utf-8-sig'))
check('actual continuous late-cycle motion',m['passed'] and all(f['passed'] for f in m['fits']))
check('actual samples over three seventeen-second cycles',m['actualPTS'][-1]>53 and m['actualFrames']==len(m['actualPTS']))
result={'passed':all(c['passed'] for c in checks),'count':len(checks),'checks':checks};out.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps({'passed':result['passed'],'checks':len(checks)}));raise SystemExit(0 if result['passed'] else 1)
