import argparse,json,zipfile,xml.etree.ElementTree as E,posixpath
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('folder');a=p.parse_args();o=Path(a.folder);P='http://schemas.openxmlformats.org/presentationml/2006/main';ns={'p':P};R='{http://schemas.openxmlformats.org/package/2006/relationships}'
with zipfile.ZipFile(o/'master-loop.pptx') as z:
 expected=[('ppt/slideMasters/slideMaster1.xml','master-orange','M 0 0 L 0.2 0 E'),('ppt/slideLayouts/slideLayout1.xml','layout-A-teal','M 0 0 L 0 0.15 E'),('ppt/slideLayouts/slideLayout2.xml','layout-B-purple','M 0 0 L 0 -0.15 E')];rows=[]
 for part,name,path in expected:
  r=E.fromstring(z.read(part));targets=[v for v in r.findall('.//p:cNvPr',ns) if v.get('name')==name];assert len(targets)==1;id=targets[0].get('id')
  effects=r.findall('.//p:cTn[@presetClass]',ns);assert len(effects)==1;e=effects[0];assert e.get('presetClass')=='path' and e.get('autoRev')=='1' and e.get('repeatCount')=='indefinite'
  assert {v.get('spid') for v in r.find('p:timing',ns).iter() if v.get('spid')}=={id}
  motion=e.find('.//p:animMotion',ns);assert motion.get('path')==path and motion.find('p:cBhvr/p:cTn',ns).get('dur')=='1000';assert e.get('accel','0')=='0' and e.get('decel','0')=='0'
  rows.append({'part':part,'targetName':name,'targetId':id,'repeat':'indefinite','autoReverse':True,'forwardMs':1000})
 for i,layout in enumerate([1,1,2],1):
  r=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'));names=[v.get('name') for v in r.findall('p:cSld/p:spTree/p:sp/p:nvSpPr/p:cNvPr',ns)];assert names==['local-title','local-note','local-blue'],'Inherited shapes must not be duplicated into pages'
  effects=r.findall('.//p:cTn[@presetClass]',ns);assert len(effects)==1 and effects[0].get('presetClass')=='entr' and effects[0].get('autoRev','0')=='0' and effects[0].get('repeatCount','1000')=='1000'
  rel=E.fromstring(z.read(f'ppt/slides/_rels/slide{i}.xml.rels'));lr=next(v for v in rel if v.get('Type').endswith('/slideLayout'));resolved=lr.get('Target').lstrip('/') if lr.get('Target').startswith('/') else posixpath.normpath(posixpath.join('ppt/slides',lr.get('Target')));assert resolved==f'ppt/slideLayouts/slideLayout{layout}.xml'
  rows.append({'slide':i,'layout':layout,'localShapes':names,'localEffectCount':1})
native=o/'native/objects.json'
if native.exists():
 nativeRows=json.loads(native.read_text(encoding='utf-8-sig'));assert [r['owner'] for r in nativeRows]==['master','layout-A','layout-B','slide-1','slide-2','slide-3']
 for i,r in enumerate(nativeRows):
  assert len(r['effects'])==1;e=r['effects'][0]
  if i<3:assert e['autoReverse']==-1 and e['repeatCount']==-2147483648 and abs(e['duration']-1)<0.0001 and e['trigger']==2
  else:assert r['shapeNames']==['local-title','local-note','local-blue'] and e['autoReverse']==0 and e['repeatCount']==1 and abs(e['duration']-0.4)<0.0001 and e['trigger']==1
 rows.append({'nativeIndependentTimelines':True,'untilEndNativeSentinel':-2147483648})
(o/'structure-audit.json').write_text(json.dumps({'passed':True,'checks':rows,'scope':'Structure/native object checks; actual live pixels are separate'},indent=2),encoding='utf-8');print('Independent master/layout/local timing and relationship audit passed')
