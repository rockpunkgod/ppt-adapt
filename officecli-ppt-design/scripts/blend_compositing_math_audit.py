"""Independent scalar fixtures and locality checks for original compositing fields."""
import argparse,json,math
from pathlib import Path
from PIL import Image
p=argparse.ArgumentParser();p.add_argument('--fields',type=Path,required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():p.error('Fresh report required')
checks=[]
def check(name,ok,**meta):checks.append({'name':name,'pass':bool(ok),**meta})
def pixel(name,x,y):return Image.open(a.fields/(name+'.png')).convert('RGB').getpixel((x,y))
fixtures=[('multiply-encoded',(128,128,128)),('overlay-encoded',(255,255,255)),('softlight-encoded',(181,181,181)),('normal-50-encoded',(192,192,192)),('multiply-50-encoded',(128,128,128))]
for name,expected in fixtures:
 actual=pixel(name,160,288);check(name+' B128 S255',actual==expected,actual=actual,expected=expected)
for mode in ['multiply','overlay','softlight']:
 b=Image.open(a.fields/'backdrop.png').convert('RGB');m=Image.open(a.fields/'local-mask.png');im=Image.open(a.fields/('local-'+mode+'-encoded.png')).convert('RGB');outside=inside=different=0
 for y in range(b.height):
  for x in range(b.width):
   if m.getpixel((x,y))==0:outside+=im.getpixel((x,y))!=b.getpixel((x,y))
   else:inside+=1;different+=im.getpixel((x,y))!=b.getpixel((x,y))
 check(mode+' local mask outside unchanged',outside==0,changedOutside=outside,insidePixels=inside,changedInside=different)
 check(mode+' local effect present',different>10000,changedInside=different)
check('Multiply50 differs from Normal50',pixel('multiply-50-encoded',160,288)!=pixel('normal-50-encoded',160,288))
check('SoftLight differs from Overlay',pixel('softlight-encoded',160,288)!=pixel('overlay-encoded',160,288))
check('Encoding-space alternatives differ',pixel('softlight-encoded',160,288)!=pixel('softlight-linear-light',160,288),encoded=pixel('softlight-encoded',160,288),linear=pixel('softlight-linear-light',160,288))
a.report.write_text(json.dumps({'checks':checks,'passed':sum(x['pass'] for x in checks),'count':len(checks)},indent=2),encoding='utf-8')
assert all(c['pass'] for c in checks)
print(len(checks),'passed')
