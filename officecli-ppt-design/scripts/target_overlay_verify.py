"""Independent OPC checks for zero-axis target/actual overlap."""
import argparse,json,posixpath,zipfile,xml.etree.ElementTree as E
from pathlib import Path
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
def verify(deck):
 rows=[]
 with zipfile.ZipFile(deck) as z:
  for page,target in [(1,100),(2,120)]:
   slide=E.fromstring(z.read(f'ppt/slides/slide{page}.xml'))
   rels={q.get('Id'):q.get('Target') for q in E.fromstring(z.read(f'ppt/slides/_rels/slide{page}.xml.rels'))}
   refs=slide.findall('.//c:chart',N);assert len(refs)==1
   part=posixpath.normpath(posixpath.join('ppt/slides',rels[refs[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')])).lstrip('/')
   root=E.fromstring(z.read(part));bar=root.find('.//c:barChart',N);assert bar is not None
   for key,value in [('barDir','col'),('grouping','clustered'),('overlap','100'),('gapWidth','300')]:assert bar.find('c:'+key,N).get('val')==value
   axes=root.findall('.//c:valAx',N);assert len(axes)==1
   assert axes[0].find('c:scaling/c:min',N).get('val')=='0'
   assert axes[0].find('c:scaling/c:max',N).get('val')=='140'
   series=bar.findall('c:ser',N);assert len(series)==2
   for j,ser in enumerate(series):
    assert ser.findtext('c:tx/c:v',namespaces=N)==['Target','Actual'][j]
    assert [q.findtext('c:v',namespaces=N) for q in ser.findall('c:cat/c:strLit/c:pt',N)]==['Q1','Q2','Q3','Q4']
    assert [float(q.findtext('c:v',namespaces=N)) for q in ser.findall('c:val/c:numLit/c:pt',N)]==([target]*4 if j==0 else [45,60,75,90])
    fill=ser.find('c:spPr/a:solidFill/a:srgbClr',N);assert fill is not None and fill.get('val')=='1C6561'
    alpha=fill.find('a:alpha',N)
    assert (alpha.get('val') if alpha is not None else '100000')==('15000' if j==0 else '100000')
   assert root.find('c:externalData',N) is None
   rows.append({'page':page,'chartPart':part,'target':target,'actual':[45,60,75,90],'primaryAxis':[0,140],'targetAlpha':.15,'actualAlpha':1,'overlapPercent':100,'gapWidth':300})
 assert rows[0]['chartPart']!=rows[1]['chartPart'],'Clone must have independent chart data'
 return {'passed':True,'pages':rows,'scope':'Native clustered chart literal data; no embedded-workbook/Edit Data claim.'}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('deck');p.add_argument('--report',required=True);a=p.parse_args();out=Path(a.report)
 if out.exists():raise FileExistsError(out)
 out.write_text(json.dumps(verify(a.deck),indent=2),encoding='utf-8');print(out)
