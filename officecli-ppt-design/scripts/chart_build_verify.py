"""Independent real chart subelement targets, order, durations and click evidence."""
import argparse,json,zipfile,posixpath,xml.etree.ElementTree as E
from pathlib import Path
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
def verify(deck):
 reports=[]
 with zipfile.ZipFile(deck) as z:
  for slide,mode,total in [(1,'series',2),(2,'category',4)]:
   root=E.fromstring(z.read(f'ppt/slides/slide{slide}.xml'))
   charts=root.findall('.//p:graphicFrame',N);assert len(charts)==1
   sid=charts[0].find('p:nvGraphicFramePr/p:cNvPr',N).get('id')
   build=root.find('.//p:bldGraphic',N);assert build is not None and build.get('spid')==sid
   spec=build.find('p:bldSub/a:bldChart',N);assert spec is not None and spec.get('bld')==mode
   rels={q.get('Id'):q.get('Target') for q in E.fromstring(z.read(f'ppt/slides/_rels/slide{slide}.xml.rels'))}
   rid=charts[0].find('.//c:chart',N).get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
   part=posixpath.normpath(posixpath.join('ppt/slides',rels[rid])).lstrip('/')
   cr=E.fromstring(z.read(part));series=cr.findall('.//c:barChart/c:ser',N);assert len(series)==2
   assert cr.find('.//c:barDir',N).get('val')=='col'
   assert cr.find('.//c:grouping',N).get('val')=='clustered'
   data=[]
   for j,ser in enumerate(series):
    name=ser.findtext('c:tx/c:v',namespaces=N);cats=[q.findtext('c:v',namespaces=N) for q in ser.findall('c:cat/c:strLit/c:pt',N)]
    vals=[float(q.findtext('c:v',namespaces=N)) for q in ser.findall('c:val/c:numLit/c:pt',N)]
    assert name==['Alpha','Beta'][j] and cats==['Q1','Q2','Q3','Q4']
    assert vals==[[24,38,46,60],[36,29,54,42]][j]
    data.append({'name':name,'categories':cats,'values':vals})
   assert cr.find('c:externalData',N) is None,'This tested subset does not materialize a workbook'
   groups=root.findall('.//p:timing//p:cTn[@presetID]',N)
   assert len(groups)==total,'Only the authored series/category data subtargets are expected in the OfficeCLI file'
   steps=[]
   for idx,g in enumerate(groups):
    assert g.get('presetID')=='10' and g.get('presetClass')=='entr' and g.get('nodeType')=='clickEffect'
    assert g.get('fill')=='hold'
    targets=g.findall('.//p:spTgt',N);assert len(targets)==2 and all(t.get('spid')==sid for t in targets)
    expected={'bldStep':mode,'seriesIdx':str(idx) if mode=='series' else '-4','categoryIdx':str(idx) if mode=='category' else '-4'}
    for target in targets:assert target.find('p:graphicEl/a:chart',N).attrib==expected
    duration=g.find('p:childTnLst/p:animEffect/p:cBhvr/p:cTn',N);assert duration is not None and duration.get('dur')=='500'
    steps.append(expected)
   reports.append({'slide':slide,'shapeId':sid,'build':mode,'nativePlayerDefaultFrameClicks':0,'dataClicks':total,'steps':steps,'durationMs':500,'series':data})
 return {'passed':True,'slides':reports,'scope':'OfficeCLI-authored literal native chart + Fade series/category builds. Does not claim embedded workbook or PowerPoint Edit Data capability.'}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('deck');p.add_argument('--report',required=True);p.add_argument('--desktop');p.add_argument('--pixels');a=p.parse_args();out=Path(a.report)
 if out.exists():raise FileExistsError('Refusing overwrite')
 result=verify(a.deck)
 if a.desktop:
  clicks=json.loads((Path(a.desktop)/'clicks.json').read_text(encoding='utf-8-sig'));assert len(clicks)==8
  for q in clicks:assert q['slide']==q['actualSlide'] and q['click']==q['index']
  result['physicalClickStates']=8
 if a.pixels:
  pixel=json.loads(Path(a.pixels).read_text(encoding='utf-8'));assert pixel['passed'] and len(pixel['checks'])==8 and all(q['passed'] for q in pixel['checks']);result['pixelStates']=8
 out.write_text(json.dumps(result,indent=2),encoding='utf-8');print(out)
