"""Research-only native PNG comparison against the intact reference image (Pillow)."""
import argparse,json,math
from pathlib import Path
from PIL import Image,ImageChops,ImageStat
p=argparse.ArgumentParser();p.add_argument('evidence');p.add_argument('--structure',required=True);p.add_argument('--report',required=True);a=p.parse_args();d=Path(a.evidence);out=Path(a.report)
if out.exists():raise FileExistsError('Refusing overwrite')
m=json.loads(Path(a.structure).read_text(encoding='utf-8'));ref=Image.open(d/'slide-1.png').convert('RGB');rb=m['referenceBoxEmu']
def pixels(b):return [round(b[0]/12193200*1600),round(b[1]/6858000*900),round((b[0]+b[2])/12193200*1600),round((b[1]+b[3])/6858000*900)]
def metric(img,box):
    x,y,r,b=pixels(box);region=(x+5,y+5,r-5,b-5);diff=ImageChops.difference(img.crop(region),ref.crop(region));rms=math.sqrt(sum(v*v for v in ImageStat.Stat(diff).rms)/3);return {'box':region,'rms':rms}
groups=[]
for g in m['mappedGroups']:
    img=Image.open(d/f'slide-{g["slide"]}.png').convert('RGB');parts=[metric(img,p['worldBoxEmu']) for p in g['pieces']]
    assert all(p['rms']<0.2 for p in parts),parts
    groups.append({'slide':g['slide'],'passed':True,'parts':parts})
probes=[]
for i in (4,5):
    path=d/f'slide-{i}.png'
    if path.exists():
        img=Image.open(path).convert('RGB');parts=[metric(img,p['worldBoxEmu']) for p in m['mappedGroups'][0]['pieces']]
        probes.append({'slide':i,'matchesSharedImage':all(p['rms']<0.2 for p in parts),'parts':parts})
out.write_text(json.dumps({'passed':True,'mappedGroups':groups,'probes':probes,'rmsThreshold':0.2,'edgePixelsExcluded':5},indent=2),encoding='utf-8');print(out)
