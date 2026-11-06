import json,zipfile,subprocess,shutil,xml.etree.ElementTree as E,posixpath,argparse
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);a=p.parse_args();r=Path(a.out);f=r/'summary-final.pptx'
for target in (f,r/'click-plan.json',r/'final-labels.json'):
 if target.exists():raise FileExistsError('Refusing overwrite: '+str(target))
shutil.copyfile(r/'summary-edited.pptx',f);cli=a.officecli
ops=[]
for name,text,x,y,w,h,size in [('summary-title','EXPLORE / THREE CHAPTERS',1.6,.6,30,1.2,24),('summary-footer','NOVA / SELECT A CHAPTER. RETURN HERE AT ITS END.',1.7,18.2,30,.6,11)]:ops.append(dict(command='add',parent='/slide[1]',type='shape',props=dict(name=name,text=text,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',font='Arial',size=str(size),color='17324E',fill='none',line='none',margin='0cm')))
b=r/'final-labels.json';b.write_text(json.dumps(ops));p=subprocess.run([cli,'batch',str(f),'--input',str(b),'--json'],capture_output=True,text=True,encoding='utf-8');assert p.returncode==0 and json.loads(p.stdout)['success'];subprocess.run([cli,'close',str(f)],capture_output=True,check=True)
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','s':'http://schemas.microsoft.com/office/powerpoint/2016/summaryzoom','p14':'http://schemas.microsoft.com/office/powerpoint/2010/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
with zipfile.ZipFile(f) as z:
 p=E.fromstring(z.read('ppt/presentation.xml'));s=E.fromstring(z.read('ppt/slides/slide1.xml'));order=[x.get('id') for x in p.findall('p:sldIdLst/p:sldId',N)];sections={x.get('id'):{'name':x.get('name'),'slides':[e.get('id') for e in x.findall('p14:sldIdLst/p14:sldId',N)]} for x in p.findall('.//p14:section',N)};frame=s.find('.//a:graphicData[@uri="'+N['s']+'"]/../..',N);xf=frame.find('p:xfrm',N);off=xf.find('a:off',N);ox,oy=int(off.get('x')),int(off.get('y'));cards=[];steps=[]
 for i,item in enumerate(s.findall('.//s:summaryZmObj',N)):
  xf=item.find('.//a:xfrm',N);of=xf.find('a:off',N);ex=xf.find('a:ext',N);sec=sections[item.get('sectionId')];first=order.index(sec['slides'][0])+1;last=order.index(sec['slides'][-1])+1;x=(ox+int(of.get('x'))+int(ex.get('cx'))/2)/360000;y=(oy+int(of.get('y'))+int(ex.get('cy'))/2)/360000;card=dict(name=sec['name'],first=first,last=last,x=x,y=y);cards.append(card)
  steps.extend([dict(name=f'chapter-{i+1}-enter',x=x,y=y,expected=first),dict(name=f'chapter-{i+1}-content',x=31,y=18,expected=last),dict(name=f'chapter-{i+1}-return',x=31,y=18,expected=1)])
(r/'click-plan.json').write_text(json.dumps(dict(widthCm=33.87,heightCm=19.05,cards=cards,steps=steps),indent=2));print(json.dumps(cards))
