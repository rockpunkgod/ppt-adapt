"""Independent geometric/pixel reconstruction of local stroke order and clipping."""
import argparse,json,math,hashlib,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import numpy as np
from PIL import Image,ImageDraw,ImageFilter
from stroke_depth_geometry import STROKES,BRIDGE,LOGICAL_BRIDGE,PATCHES,PATCH_RECEIVERS,CURVED_LEFT,flatten,inside,area
p=argparse.ArgumentParser();p.add_argument('--build',type=Path,required=True);p.add_argument('--assets',type=Path,required=True);p.add_argument('--native',type=Path,required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():p.error('Fresh report required')
plan=json.loads((a.build/'plan.json').read_text(encoding='utf-8'));manifest=json.loads((a.assets/'manifest.json').read_text(encoding='utf-8'));shades={r['file']:r for r in manifest['sourceFields']};images={f.name:Image.open(f).convert('RGBA') for f in a.assets.glob('*.png')};checks=[];stats=[]
def check(name,ok,**meta):checks.append({'name':name,'pass':bool(ok),**meta})
def distance(point,poly):
 x,y=point;best=1e9
 for (ax,ay),(bx,by) in zip(poly,poly[1:]+poly[:1]):
  dx,dy=bx-ax,by-ay;t=max(0,min(1,((x-ax)*dx+(y-ay)*dy)/(dx*dx+dy*dy))) if dx*dx+dy*dy else 0;best=min(best,math.hypot(x-ax-t*dx,y-ay-t*dy))
 return best
def expected(slide,cmx,cmy):
 rgb=(16,30,37);owner=None;safe=True
 for glyph in [g for g in plan if g['slide']==slide]:
  gx,gy,gw,gh=glyph['rectCm'];x=(cmx-gx)/gw*600;y=(cmy-gy)/gh*600
  for item in glyph['order']:
   path=BRIDGE if item['name']=='bridge' else PATCHES[item['name']] if item['name'] in PATCHES else STROKES[item['name']]
   if distance((x,y),path)<4:safe=False
   if not inside((x,y),path):continue
   base=images[item['base']].getpixel((0,0))[:3];rgb=base;owner=glyph['prefix']+':'+item['name']
   if item['effect']:
    src=images[item['source']];l,t,r,b=shades[item['source']]['bounds'];sx=max(0,min(src.width-1,(x-l)/(r-l)*src.width-0.5));sy=max(0,min(src.height-1,(y-t)/(b-t)*src.height-0.5));ix=int(sx);iy=int(sy);fx=sx-ix;fy=sy-iy
    s=[sum(src.getpixel((min(src.width-1,ix+dx),min(src.height-1,iy+dy)))[c]*wx*wy for dx,wx in [(0,1-fx),(1,fx)] for dy,wy in [(0,1-fy),(1,fy)]) for c in range(4)];alpha=s[3]/255;rgb=tuple(round(v*((1-alpha)+alpha*s[c]/255)) for c,v in enumerate(base))
 return rgb,owner,safe
for slide in [1,2,3]:
 im=np.asarray(Image.open(a.native/f'native-{slide}.png').convert('RGB'));re=np.asarray(Image.open(a.native/f'reopened-{slide}.png').convert('RGB'));check('Native reopen identical '+str(slide),bool(np.array_equal(im,re)))
 errors=[];owners={};outside=0;worst=[]
 for py in range(245,905,10):
  for px in range(110,1790,10):
   ex,owner,safe=expected(slide,px/1920*33.87,py/1080*19.05)
   if not safe:continue
   error=int(np.abs(im[py,px].astype(int)-np.array(ex)).max());errors.append(error);owners[owner or 'background']=owners.get(owner or 'background',0)+1
   if owner is None:outside=max(outside,error)
   if error>3:worst.append({'xy':[px,py],'owner':owner,'actual':im[py,px].tolist(),'expected':ex,'error':error})
 stats.append({'slide':slide,'sampleCount':len(errors),'maxError':max(errors),'outsideMaxError':outside,'ownerCounts':owners,'bad':worst[:20]});check('Path/layer/gradient reconstruction '+str(slide),max(errors)<=3,maximum=max(errors),samples=len(errors));check('No stroke-exterior leak '+str(slide),outside==0,actual=outside)
 mask=Image.new('L',(1920,1080),0);draw=ImageDraw.Draw(mask)
 for g in [g for g in plan if g['slide']==slide]:
  gx,gy,gw,gh=g['rectCm']
  for poly in STROKES.values():draw.polygon([((gx+px/600*gw)/33.87*1920,(gy+py/600*gh)/19.05*1080) for px,py in poly],fill=255)
 envelope=np.asarray(mask.filter(ImageFilter.MaxFilter(7)))>0;roi=np.zeros_like(envelope);roi[245:905,110:1790]=True;outside_pixels=roi&~envelope;d=np.abs(im.astype(int)-np.array([16,30,37]));leak=int(d[outside_pixels].max());check('All-pixel exterior outside 3px boundary '+str(slide),leak==0,pixels=int(outside_pixels.sum()),maxDifference=leak)
def nativepoint(slide,prefix,xy):
 g=next(g for g in plan if g['slide']==slide and g['prefix']==prefix);x,y,w,h=g['rectCm'];px=round((x+xy[0]/600*w)/33.87*1920);py=round((y+xy[1]/600*h)/19.05*1080);return Image.open(a.native/f'native-{slide}.png').convert('RGB').getpixel((px,py))
before=nativepoint(3,'without-bridge',(300,315));after=nativepoint(3,'with-bridge',(300,315));check('Local bridge changes front owner at same crossing',before[1]>before[0]*2 and after[0]>after[1],before=before,after=after)
seam=[nativepoint(3,'with-bridge',(x,315)) for x in [325,327,329,330,331,333,335]];seam_error=max(max(abs(c-e) for c,e in zip(rgb,(255,161,66))) for rgb in seam);check('Bridge-to-original-stroke continuity',seam_error<=3,colors=seam,maximum=seam_error)
fade=[nativepoint(3,'with-bridge',point) for point in [(166,400),(168,400),(170,400),(172,400),(174,400),(190,380),(210,360),(230,340),(250,320),(260,310)]];check('Local gradient starts continuously and darkens',all(fade[i][2]>=fade[i+1][2] for i in range(len(fade)-1)) and max(abs(a-b) for a,b in zip(fade[0],fade[2]))<=3 and fade[-1][2]<fade[0][2]/2,colors=fade)
check('Real logical intersection has area, not bbox',1000<abs(area(LOGICAL_BRIDGE))<4000,intersectionArea=abs(area(LOGICAL_BRIDGE)),bboxArea=60*(354.7826086956522-260))
check('Seam bleed remains inside right stroke',all(inside(point,STROKES['right']) or distance(point,STROKES['right'])<0.01 for point in BRIDGE) and abs(area(BRIDGE))>abs(area(LOGICAL_BRIDGE)),logicalArea=abs(area(LOGICAL_BRIDGE)),renderArea=abs(area(BRIDGE)))
for name,poly in PATCHES.items():
 receiver=STROKES[PATCH_RECEIVERS[name]];check('Closed patch lies within actual receiving stroke '+name,abs(area(poly))>50 and all(inside(point,receiver) or distance(point,receiver)<0.01 for point in poly),area=abs(area(poly)),points=len(poly))
curve_error=max(distance(point,STROKES['left']) for point in flatten(CURVED_LEFT,160));check('Curve subdivision converges below 0.05 path units',curve_error<0.05,maximum=curve_error,coarseSteps=80,fineSteps=160)
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
counts=[]
for deck in [a.build/'stroke-depth.pptx',a.native/'roundtrip.pptx']:
 with zipfile.ZipFile(deck) as z:
  deckcounts=[]
  for i in [1,2,3]:
   root=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'));check('Original text preserved '+deck.parent.name+str(i),'木木' in ''.join(root.itertext()));deckcounts.append({'custom':len(root.findall('.//a:custGeom',ns)),'cubic':len(root.findall('.//a:cubicBezTo',ns)),'mult':len(root.findall('.//a:fillOverlay',ns))})
  counts.append(deckcounts)
check('Native path/Bezier/mult survive roundtrip',counts[0]==counts[1],counts=counts)
check('Original closed vector geometry and native multiply present',all(c['custom']>=8 and c['cubic']>=4 and c['mult']>=4 for c in counts[0]))
for g in plan:
 if g['detail']:check('Independent closed patch objects '+str(g['slide'])+g['prefix'],sum(bool(i.get('closedPatch')) for i in g['order'])==3)
readback=json.loads((a.native/'native-readback.json').read_text(encoding='utf-8'));check('Input SHA unchanged',readback['unchanged'])
for name,sha in manifest['files'].items():check('Asset SHA '+name,hashlib.sha256((a.assets/name).read_bytes()).hexdigest()==sha)
out={'passed':sum(c['pass'] for c in checks),'count':len(checks),'checks':checks,'pixels':stats};a.report.write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8');print(out['passed'],'/',out['count']);print([(r['slide'],r['sampleCount'],r['maxError'],r['outsideMaxError']) for r in stats]);assert all(c['pass'] for c in checks)
