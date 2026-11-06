"""Check structural invariants on an explicitly supplied recipe directory."""
import argparse,json,zipfile
from pathlib import Path
from xml.etree import ElementTree as E
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
p=argparse.ArgumentParser();p.add_argument('--artifacts',required=True,type=Path);p.add_argument('--report',required=True,type=Path);a=p.parse_args()
if a.report.exists():p.error('Refusing to overwrite report: '+str(a.report))
rows=[]
def ck(n,v):
 rows.append({'name':n,'passed':bool(v)})
 if not v:raise AssertionError(n)
f=a.artifacts/'alignment.pptx'
if f.exists():
 z=zipfile.ZipFile(f);r=E.fromstring(z.read('ppt/slides/slide1.xml'));b=[]
 for i in range(3):
  s=next(s for s in r.findall('.//p:sp',ns)if s.find('p:nvSpPr/p:cNvPr',ns).get('name')=='card-'+str(i));t=s.find('p:spPr/a:xfrm',ns);o=t.find('a:off',ns);e=t.find('a:ext',ns);b.append([int(o.get('x')),int(o.get('y')),int(e.get('cx')),int(e.get('cy'))])
 ck('Common bottom edge',len(set(y+h for x,y,w,h in b))==1);g=[b[i+1][0]-b[i][0]-b[i][2]for i in range(2)];ck('Equal edge gaps',abs(g[0]-g[1])<=1)
 r=E.fromstring(z.read('ppt/slides/slide2.xml'));q=r.find('.//p:grpSp',ns);ck('Real group with three members',q is not None and len(q.findall('p:sp',ns))==3)
 t=q.find('p:grpSpPr/a:xfrm',ns);o=t.find('a:off',ns);e=t.find('a:ext',ns);ck('Group centered on page',abs(2*int(o.get('x'))+int(e.get('cx'))-round(33.87*360000))<=1 and abs(2*int(o.get('y'))+int(e.get('cy'))-round(19.05*360000))<=1)
f=a.artifacts/'visual-design.pptx'
if f.exists():
 z=zipfile.ZipFile(f);ck('Vector SVG part retained',any(n.endswith('.svg')for n in z.namelist()));ck('Native chart retained',any('/charts/chart' in n and n.endswith('.xml')for n in z.namelist()))
if not rows:p.error('No known layout story recipe decks in supplied directory')
a.report.parent.mkdir(parents=True,exist_ok=True);a.report.write_text(json.dumps({'checks':rows,'passed':all(x['passed']for x in rows),'limitation':'Structural checks do not replace native visual review.'},indent=2),encoding='utf-8')
print(json.dumps({'checks':len(rows),'passed':True}))
