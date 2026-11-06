import argparse,hashlib,json,posixpath,zipfile,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--deck',type=Path,required=True);p.add_argument('--original',type=Path,required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','a14':'http://schemas.microsoft.com/office/drawing/2010/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'};original=a.original.read_bytes();checks={};rows=[]
with zipfile.ZipFile(a.deck) as z:
 names=set(z.namelist());missing=[]
 for relfile in names:
  if not relfile.endswith('.rels'):continue
  base='' if relfile=='_rels/.rels' else posixpath.dirname(posixpath.dirname(relfile))
  for rel in E.fromstring(z.read(relfile)):
   if rel.get('TargetMode')=='External':continue
   target=rel.get('Target');resolved=target.lstrip('/') if target.startswith('/') else posixpath.normpath(base+'/'+target if base else target)
   if resolved not in names:missing.append({'relationshipPart':relfile,'id':rel.get('Id'),'target':resolved})
 checks['allInternalRelationshipTargetsReachable']=not missing
 def part(s,rid):
  rels={r.get('Id'):r.get('Target') for r in E.fromstring(z.read(f'ppt/slides/_rels/slide{s}.xml.rels'))};target=rels[rid];name=target.lstrip('/') if target.startswith('/') else posixpath.normpath('ppt/slides/'+target);return name,z.read(name)
 slides={s:E.fromstring(z.read(f'ppt/slides/slide{s}.xml')) for s in [1,2,3,4]};blip=slides[2].find('.//p:pic/p:blipFill/a:blip',ns);cache_part,cache=part(2,blip.get('{'+ns['r']+'}embed'))
 effect=blip.find('.//a14:artisticBlur',ns);layer=blip.find('.//a14:imgLayer',ns);checks['editableNativeArtisticBlur30']=effect is not None and effect.get('radius')=='30' and layer is not None
 hdpart,hd=part(2,layer.get('{'+ns['r']+'}embed'));checks['nativeOriginalLayerBound']=bool(hd) and hdpart.endswith('.wdp');checks['referenceCacheDifferentFromOriginal']=cache!=original
 for s in [1,3,4]:
  fg=slides[s].find('.//p:pic/p:blipFill/a:blip',ns);_,data=part(s,fg.get('{'+ns['r']+'}embed'));checks[f's{s}ClearOriginalBytes']=data==original
 for s in [3,4]:
  bg=slides[s].find('p:cSld/p:bg/p:bgPr/a:blipFill/a:blip',ns);_,data=part(s,bg.get('{'+ns['r']+'}embed'));checks[f's{s}BackgroundExactNativeCache']=data==cache
  shapes=slides[s].findall('p:cSld/p:spTree/p:sp',ns);checks[f's{s}TwoTrueBackgroundWindows']=len(shapes)==2 and all(q.get('useBgFill')=='1' and q.find('.//a:blipFill',ns) is None for q in shapes)
  rows.append([{'name':q.find('p:nvSpPr/p:cNvPr',ns).get('name'),'xy':[int(v) for v in q.find('p:spPr/a:xfrm/a:off',ns).attrib.values()]} for q in shapes])
 checks['windowsActuallyMoved']=all(q['xy']!=r['xy'] for q,r in zip(rows[0],rows[1]))
result={'passed':all(checks.values()),'checks':checks,'missingTargets':missing,'originalSHA256':hashlib.sha256(original).hexdigest(),'nativeCacheSHA256':hashlib.sha256(cache).hexdigest(),'nativeCachePart':cache_part,'nativeOriginalLayerPart':hdpart,'windows':rows,'note':'Blur is editable on reference picture; window background is exact native-produced cache snapshot, regenerated via native polish+CLI bridge. Not external blur; not a live Background.Fill blur parameter.'}
a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2))
