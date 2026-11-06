import argparse,hashlib,json,zipfile,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--deck',type=Path,required=True);p.add_argument('--wav',type=Path,required=True);p.add_argument('--native',type=Path,required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main'}
with zipfile.ZipFile(a.deck) as z:
 slides=[E.fromstring(z.read(f'ppt/slides/slide{i}.xml')) for i in range(1,4)];media=slides[0].find('.//p:audio/p:cMediaNode',ns)
 waves=[hashlib.sha256(z.read(n)).hexdigest() for n in z.namelist() if n.startswith('ppt/media/') and n.endswith('.wav')]
 checks={'threeSlides':len([n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')])==3,'nativeAudioTwoSlides':media is not None and media.get('numSld')=='2','noAudioOnLaterPages':all(s.find('.//p:audio',ns) is None for s in slides[1:]),'embeddedOriginalWav':waves==[hashlib.sha256(a.wav.read_bytes()).hexdigest()]}
native=json.loads(a.native.read_text(encoding='utf-8-sig'));checks['nativeSettings']=all(native[k]==v for k,v in {'playOnEntry':-1,'pauseAnimation':0,'stopAfterSlides':2,'loopUntilStopped':0,'start':0,'end':12000,'fadeIn':0,'fadeOut':0}.items())
result={'passed':all(checks.values()),'checks':checks,'embeddedWavSHA256':waves,'nativeSettings':native}
a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2))
