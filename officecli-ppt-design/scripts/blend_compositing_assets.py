"""Deterministic original RGB fields and explicitly defined blend references."""
import argparse,json,hashlib
from pathlib import Path
import numpy as np
from PIL import Image,ImageDraw,ImageCms
p=argparse.ArgumentParser();p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():p.error('Fresh output directory required')
a.outdir.mkdir(parents=True);profile=ImageCms.ImageCmsProfile(ImageCms.createProfile('sRGB')).tobytes()
bp=np.array([[0,0,0],[64,64,64],[128,128,128],[192,192,192],[255,255,255],[32,96,224],[224,64,96],[96,224,32]],dtype=np.uint8)
sp=np.array([[0,0,0],[64,64,64],[128,128,128],[192,192,192],[255,255,255],[255,128,64]],dtype=np.uint8)
b=np.repeat(np.repeat(np.tile(bp[None,:,:],(6,1,1)),64,axis=0),64,axis=1);s=np.repeat(np.repeat(np.tile(sp[:,None,:],(1,8,1)),64,axis=0),64,axis=1)
def write(name,data):Image.fromarray(data).save(a.outdir/name,icc_profile=profile)
def quant(data):return np.floor(np.clip(data,0,1)*255+0.5).astype(np.uint8)
def decode(x):return np.where(x<=0.04045,x/12.92,((x+0.055)/1.055)**2.4)
def encode(x):return np.where(x<=0.0031308,x*12.92,1.055*np.maximum(x,0)**(1/2.4)-0.055)
def blend(B,S,mode):
 if mode=='multiply':return B*S
 if mode=='screen':return 1-(1-B)*(1-S)
 if mode=='normal':return S
 if mode=='overlay':return np.where(B<=0.5,2*B*S,1-2*(1-B)*(1-S))
 if mode=='softlight':
  D=np.where(B<=0.25,((16*B-12)*B+4)*B,np.sqrt(B));return np.where(S<=0.5,B-(1-2*S)*B*(1-B),B+(2*S-1)*(D-B))
 raise ValueError(mode)
def composite(B,S,mode,alpha=1,linear=False):
 B=B.astype(float)/255;S=S.astype(float)/255
 if linear:B,S=decode(B),decode(S)
 C=(1-alpha)*B+alpha*blend(B,S,mode)
 return quant(encode(C) if linear else C)
write('backdrop.png',b);write('source.png',s)
for mode in ['normal','multiply','screen','overlay','softlight']:
 for space,linear in [('encoded',False),('linear-light',True)]:
  write(f'{mode}-{space}.png',composite(b,s,mode,linear=linear));write(f'{mode}-50-{space}.png',composite(b,s,mode,alpha=0.5,linear=linear))
gray=np.full_like(b,128);write('gray128.png',gray)
for linear in [False,True]:write('gray-multiply-'+('linear-light' if linear else 'encoded')+'.png',composite(b,gray,'multiply',linear=linear))
mask=Image.new('L',(512,384),0);ImageDraw.Draw(mask).polygon([(70,70),(445,115),(380,330),(95,285)],fill=255);m=np.asarray(mask,dtype=np.uint8);mask.save(a.outdir/'local-mask.png')
gradient=np.tile(np.rint(np.linspace(0,255,512)).astype(np.uint8)[None,:,None],(384,1,3));rgba=np.concatenate([gradient,m[...,None]],axis=2);write('local-gradient-rgba.png',rgba)
for mode in ['multiply','overlay','softlight']:
 for space,linear in [('encoded',False),('linear-light',True)]:write(f'local-{mode}-{space}.png',composite(b,gradient,mode,alpha=(m/255)[...,None],linear=linear))
manifest={'inputSpace':'sRGB ICC tagged, RGB code values normalized by /255','encodedMath':'Operate directly on sRGB code values; alpha is linear coverage. Not claimed equivalent to the source Adobe Illustrator ICC/gamma setup.','linearReference':'IEC sRGB decode -> identical blend function -> alpha mix -> sRGB encode','rounding':'clip [0,1], floor(value*255+0.5)','size':[512,384],'tile':64,'backdropPalette':bp.tolist(),'sourcePalette':sp.tolist(),'localMaskPolygon':[[70,70],[445,115],[380,330],[95,285]],'localSource':'RGB horizontal 0..255 ramp, binary alpha polygon; outside mask exact zero','formulaSource':'https://www.w3.org/TR/compositing-1/','files':{f.name:hashlib.sha256(f.read_bytes()).hexdigest() for f in a.outdir.glob('*.png')}}
(a.outdir/'manifest.json').write_text(json.dumps(manifest,indent=2),encoding='utf-8')
