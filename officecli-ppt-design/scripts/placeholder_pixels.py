"""Native PNG scope isolation and actual inherited color/override visibility; Pillow."""
import argparse,json
from pathlib import Path
from PIL import Image,ImageChops
p=argparse.ArgumentParser();p.add_argument('--evidence',required=True);p.add_argument('--report',required=True);a=p.parse_args();d=Path(a.evidence);out=Path(a.report)
if out.exists():raise FileExistsError('Refusing overwrite')
phases=['before','layout','body','swap','reset','reopen'];images={}
for phase in phases:
 for slide in range(1,5):
  im=Image.open(d/f'{phase}-{slide}.png').convert('RGB');assert im.size==(1600,900);images[phase,slide]=im
checks=[]
def equal(p,q,box,label):
 diff=ImageChops.difference(images[p].crop(box),images[q].crop(box));same=diff.getbbox() is None
 checks.append({'check':label,'passed':same,'region':box});assert same,label
def color_count(im,box,hex):
 rgb=tuple(bytes.fromhex(hex));return sum(pixel==rgb for pixel in im.crop(box).getdata())
equal(('before',4),('layout',4),(0,0,1600,900),'A layout update leaves sibling B page pixel-identical')
for slide in range(1,5):
 equal(('before',slide),('layout',slide),(0,260,1600,780),f'title-only layout update preserves body pixels on page {slide}')
 equal(('layout',slide),('body',slide),(0,0,1600,260),f'body role update preserves title pixels on page {slide}')
 equal(('reset',slide),('reopen',slide),(0,0,1600,900),f'native reopen preserves all pixels on page {slide}')
 for phase in phases[1:]:equal(('before',slide),(phase,slide),(0,800,1600,900),f'local footer retained in {phase} page {slide}')
for slide in (1,3,4):equal(('swap',slide),('reset',slide),(0,0,1600,900),f'resetting page 2 leaves page {slide} unchanged')
counts=[]
for phase in phases:
 for slide in range(1,5):
  early=phase in ('before','layout');direct=slide==2 and phase not in ('reset','reopen')
  primary='B7633C' if direct else ('17324E' if early else '1B7775');secondary='527086' if early else '2D6972'
  c1=color_count(images[phase,slide],(0,260,1600,780),primary);c2=color_count(images[phase,slide],(0,260,1600,780),secondary)
  assert c1>100 and c2>100,(phase,slide,c1,c2)
  counts.append({'phase':phase,'slide':slide,'primary':primary,'primaryPixels':c1,'secondary':secondary,'secondaryPixels':c2})
assert color_count(images['swap',2],(0,0,1600,780),'B7633C')>500
assert color_count(images['reset',2],(0,0,1600,780),'B7633C')==0
out.write_text(json.dumps({'passed':True,'isolationChecks':checks,'nativeColorCounts':counts,'resetRemovesOrangeOverride':True,'scope':'Fixed original text fixture and colors; XML/native font/placeholder audit is independently required.'},indent=2),encoding='utf-8');print(out)
