"""Original glyph contour/source gradient assets; native PPT performs the blend."""
import argparse,json,hashlib
from pathlib import Path
from PIL import Image
from stroke_depth_geometry import STROKES,BRIDGE,LOGICAL_BRIDGE,CURVED_LEFT,PATCHES,PATCH_RECEIVERS
p=argparse.ArgumentParser();p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():p.error('Fresh output required')
a.outdir.mkdir(parents=True)
colors={'green':(18,222,158),'left':(169,110,245),'right':(255,161,66),'vertical':(57,219,169),'horizontal':(85,187,245)}
for name,rgb in colors.items():Image.new('RGB',(600,600),rgb).save(a.outdir/(name+'.png'))
rows=[]
for back in [False,True]:
 for detail in [False,True]:
  for name,poly in {**STROKES,'bridge':BRIDGE,**PATCHES}.items():
   bounds=[min(p[0] for p in poly),min(p[1] for p in poly),max(p[0] for p in poly),max(p[1] for p in poly)];l,t,r,b=bounds;w=max(2,round((r-l)*2));h=max(2,round((b-t)*2));full=Image.new('RGBA',(w,h),(255,255,255,0));px=full.load();receiver=PATCH_RECEIVERS.get(name,name)
   for iy in range(h):
    for ix in range(w):
     x=l+(ix+0.5)/w*(r-l);y=t+(iy+0.5)/h*(b-t)
     factor=(0.2+0.8*x/599) if back else 1;active=back
     if detail and receiver=='vertical':
      if 230<=y<300:factor*=0.35+0.65*(y-230)/70;active=True
      edge=300-(x-270)*40/55
      if 270<=x<=330 and edge-90<=y<edge:factor*=0.35+0.65*(edge-y)/90;active=True
     if detail and receiver=='left' and 170<=x<270:factor*=0.35+0.65*(270-x)/100;active=True
     if active:g=round(255*factor);px[ix,iy]=(g,g,g,255)
   filename=f'{"back" if back else "front"}-{"detail" if detail else "flat"}-{name}.png';full.save(a.outdir/filename);rows.append({'file':filename,'bounds':bounds,'path':name,'back':back,'detail':detail,'pixelCenters':'world bounds mapped to 2x samples per path unit'})
manifest={'originalText':'木木','originalDesign':'Hand-authored 4 closed strokes per 木; not a font extraction or author template','canvas':[600,600],'paths':STROKES,'leftNativeCubicCommands':CURVED_LEFT,'bridgeOperation':'logical right ∩ vertical; render right ∩ (vertical expanded 2 path units horizontally), still confined to right stroke','logicalBridge':LOGICAL_BRIDGE,'bridge':BRIDGE,'patches':PATCHES,'patchReceivers':PATCH_RECEIVERS,'sourceCoordinates':'Each native fillOverlay source image sampled inside actual fractional geometry bounds at 2x resolution; not full path viewport','sourceFields':rows,'files':{f.name:hashlib.sha256(f.read_bytes()).hexdigest() for f in a.outdir.glob('*.png')}}
(a.outdir/'manifest.json').write_text(json.dumps(manifest,indent=2,ensure_ascii=False),encoding='utf-8')
