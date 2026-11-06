"""Independent OPC placeholder linkage, style cascade, content and native state audit."""
import argparse,json,posixpath,zipfile,xml.etree.ElementTree as E
from pathlib import Path
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
PHASES={'before':'01-before.pptx','layout':'02-layout-update.pptx','body':'03-body-role-update.pptx','swap':'04-layout-swap.pptx','reset':'placeholder-final.pptx','reopen':'placeholder-final.pptx'}
def text(sp):return '\n'.join(''.join(t.text or '' for t in q.findall('.//a:t',N)) for q in sp.findall('p:txBody/a:p',N))
def props(node):
 if node is None:return {}
 out={}
 if 'sz' in node.attrib:out['size']=int(node.get('sz'))/100
 c=node.find('a:solidFill/a:srgbClr',N)
 if c is not None:out['color']=c.get('val').upper()
 f=node.find('a:latin',N)
 if f is not None:out['font']=f.get('typeface')
 return out
def resolve(z,part,kind):
 relpart=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
 rel=next(q for q in E.fromstring(z.read(relpart)) if q.get('Type').endswith('/'+kind))
 return posixpath.normpath(posixpath.join(posixpath.dirname(part),rel.get('Target'))).lstrip('/')
def ph(root,role):
 found=[sp for sp in root.findall('p:cSld/p:spTree/p:sp',N) if (q:=sp.find('p:nvSpPr/p:nvPr/p:ph',N)) is not None and q.get('type','obj')==role]
 assert len(found)==1,(role,len(found));return found[0]
def key(sp):
 q=sp.find('p:nvSpPr/p:nvPr/p:ph',N);return q.get('type','obj'),q.get('idx','0')
def font(sp,layout,master,role,para):
 paragraphs=sp.findall('p:txBody/a:p',N);p=paragraphs[para];ppr=p.find('a:pPr',N);level=int(ppr.get('lvl','0')) if ppr is not None else 0
 out=props(master.find(f'p:txStyles/p:{role}Style/a:lvl{level+1}pPr/a:defRPr',N))
 for obj in (layout,sp):out.update(props(obj.find(f'p:txBody/a:lstStyle/a:lvl{level+1}pPr/a:defRPr',N)))
 out.update(props(p.find('a:pPr/a:defRPr',N)))
 for run in p.findall('a:r',N):out.update(props(run.find('a:rPr',N)))
 out['level']=level+1;return out
def geometry(sp,layout):
 x=sp.find('p:spPr/a:xfrm',N)
 if x is None:x=layout.find('p:spPr/a:xfrm',N)
 assert x is not None
 o=x.find('a:off',N);e=x.find('a:ext',N)
 return {'left':int(o.get('x'))/12700,'top':int(o.get('y'))/12700,'width':int(e.get('cx'))/12700}
def bgr(n):return ''.join(f'{(int(n)>>s)&255:02X}' for s in (0,8,16))
def verify(base,evidence):
 original=[]
 with zipfile.ZipFile(base) as z:
  for i in range(1,5):
   root=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'));names={sp.find('p:nvSpPr/p:cNvPr',N).get('name'):sp for sp in root.findall('p:cSld/p:spTree/p:sp',N)}
   original.append({'title':text(names['source-title']),'body':text(names['source-body']),'footer':text(names['local-caption'])})
 states={s['phase']:s['slides'] for s in json.loads((evidence/'states.json').read_text(encoding='utf-8-sig'))};assert set(states)==set(PHASES)
 results=[]
 for phase,filename in PHASES.items():
  with zipfile.ZipFile(evidence/filename) as z:
   rows=[];layouts=[]
   for i in range(1,5):
    part=f'ppt/slides/slide{i}.xml';root=E.fromstring(z.read(part));lp=resolve(z,part,'slideLayout');layout=E.fromstring(z.read(lp));mp=resolve(z,lp,'slideMaster');master=E.fromstring(z.read(mp));layouts.append(lp)
    assert len(root.findall('p:cSld/p:spTree/p:sp',N))==3,'Only two true placeholders and one local footer are allowed'
    t=ph(root,'title');body=ph(root,'body');lt=ph(layout,'title');lb=ph(layout,'body');mb=ph(master,'body')
    assert key(t)==key(lt) and key(body)==key(lb),'Slide placeholders must link to actual matching layout placeholders'
    assert text(t)==original[i-1]['title'] and text(body)==original[i-1]['body'],'Original local text must survive every operation'
    footer=next(q for q in root.findall('p:cSld/p:spTree/p:sp',N) if q.find('p:nvSpPr/p:cNvPr',N).get('name')=='local-caption')
    assert text(footer)==original[i-1]['footer']
    assert [int(q.find('a:pPr',N).get('lvl','0')) if q.find('a:pPr',N) is not None else 0 for q in body.findall('p:txBody/a:p',N)]==[0,1]
    assert body.find('p:txBody/a:lstStyle',N) is not None and len(body.find('p:txBody/a:lstStyle',N))==0
    tf=font(t,lt,master,'title',0);bf=[font(body,lb,master,'body',j) for j in (0,1)];geo=geometry(t,lt)
    direct=(i==2 and phase not in ('reset','reopen'));early=phase in ('before','layout');swapped=phase in ('swap','reset','reopen')
    expectedTitle=(40,'B7633C') if direct else ((28,'467665') if i==4 or (i==1 and swapped) or (i==2 and swapped) else ((30,'3867A8') if phase=='before' else (36,'1B7775')))
    assert (tf['size'],tf['color'])==expectedTitle,(phase,i,tf,expectedTitle)
    assert tf['font']=='Arial' and all(q['font']=='Arial' for q in bf)
    expectedBody=[(24,'17324E'),(20,'527086')] if early else [(26,'1B7775'),(18,'2D6972')]
    if direct:expectedBody[0]=(expectedBody[0][0],'B7633C')
    assert [(q['size'],q['color']) for q in bf]==expectedBody,(phase,i,bf,expectedBody)
    if direct and not swapped:assert abs(geo['left']-120)<0.01 and abs(geo['width']-700)<0.01
    elif i==4 or (swapped and i in (1,2)):assert abs(geo['left']-80)<0.01 and abs(geo['width']-800)<0.01
    else:assert abs(geo['left']-(50 if phase=='before' else 65))<0.01
    # Native data are checked against the independently resolved actual XML cascade.
    native=states[phase][i-1];assert native['shapeCount']==3 and native['title']['placeholderType']==1
    assert native['title']['size']==tf['size'] and bgr(native['title']['color'])==tf['color']
    for a,b in zip(native['body'],bf):assert a['level']==b['level'] and a['size']==b['size'] and bgr(a['color'])==b['color']
    assert all(abs(native['title'][k]-geo[k])<0.02 for k in ('left','top','width'))
    rows.append({'slide':i,'layoutPart':lp,'masterPart':mp,'titlePh':key(t),'bodyPh':key(body),'masterBodyPh':key(mb),'title':tf,'body':bf,'titleGeometry':geo,'localTextPreserved':True})
   assert layouts[2]!=layouts[3]
   if not swapped:assert layouts[:3]==[layouts[0]]*3
   else:assert layouts[0]==layouts[1]==layouts[3]
   results.append({'phase':phase,'slides':rows})
 return {'passed':True,'phases':results,'scope':'True Title1/Body2, two body levels, four original slides, layout/master role inheritance, direct per-property overlays, B layout swap, native SlideReset and reopen; theme references are outside scope.'}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--base',required=True);p.add_argument('--evidence',required=True);p.add_argument('--report',required=True);a=p.parse_args();out=Path(a.report)
 if out.exists():raise FileExistsError('Refusing overwrite')
 result=verify(a.base,Path(a.evidence));out.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(out)
