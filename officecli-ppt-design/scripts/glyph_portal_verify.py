import argparse,hashlib,json,zipfile,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--deck',type=Path,required=True);p.add_argument('--video',type=Path,required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
with zipfile.ZipFile(a.deck) as z:
 s=E.fromstring(z.read('ppt/slides/slide2.xml'));tree=s.find('p:cSld/p:spTree',ns);items=[]
 for q in tree:
  name=q.find('.//p:cNvPr',ns)
  if name is not None:items.append((name.get('name'),q))
 names=[n for n,q in items];mask=dict(items)['portal-cutout'];paths=mask.findall('.//a:custGeom/a:pathLst/a:path',ns);close=sum(len(x.findall('a:close',ns)) for x in paths)
 to=s.find('.//p:animScale/p:to',ns);grow=s.find('.//p:cTn[@presetID="6"]',ns);exits=s.findall('.//p:cTn[@presetClass="exit"]',ns)
 hashes=[hashlib.sha256(z.read(n)).hexdigest() for n in z.namelist() if n.startswith('ppt/media/') and n.endswith('.mp4')]
 checks={'realNativeCompoundHole':bool(paths) and close>=2,'videoBehindMaskBehindText':names.index('portal-video')<names.index('portal-cutout')<names.index('portal-title-overlay')<names.index('portal-caption'),'matchingForegroundText':dict(items)['portal-title-overlay'].find('.//a:t',ns).text=='H','native90xScale':to is not None and to.get('x')==to.get('y')=='9000000','grow2s':s.find('.//p:animScale/p:cBhvr/p:cTn',ns).get('dur')=='2000','maxSmoothStart':grow.get('accel')=='100000','twoUpwardFlyExits':len(exits)==2 and all(x.get('presetID')=='2' and x.get('presetSubtype')=='1' for x in exits),'allEffectsWithPrevious':all(x.get('nodeType')=='withEffect' for x in exits+[grow]),'noTransition':s.find('p:transition',ns) is None or len(s.find('p:transition',ns))==0,'embeddedOriginalVideo':hashes==[hashlib.sha256(a.video.read_bytes()).hexdigest()],'nativeAutoplay':s.find('.//p:cmd[@cmd="playFrom(0)"]',ns) is not None}
 result={'passed':all(checks.values()),'checks':checks,'pathCount':len(paths),'closedContours':close,'stacking':names,'videoSHA256':hashes}
a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2))
