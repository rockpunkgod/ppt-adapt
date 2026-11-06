"""Prepare separate strip / brush assets; actual Blast and Overlay run in an editor.

Reuses the existing image_glitch_assets.py 640x360 fixture and band offsets.
No synthetic Wind replacement or automatic verification of supplied filter output.
"""
import argparse,hashlib,json
from pathlib import Path
import numpy as np
from PIL import Image

def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--channels',type=Path,required=True)
    p.add_argument('--blast',type=Path)
    p.add_argument('--outdir',type=Path,required=True)
    a=p.parse_args()
    if a.outdir.exists():raise FileExistsError(a.outdir)
    src=Image.open(a.channels).convert('RGBA')
    if src.size!=(640,360):raise ValueError('This fixture uses 640x360 pixels; replan bands for other dimensions.')
    strips=Image.new('RGBA',src.size)
    for y0,y1,dx in [(90,101,28),(166,174,-36),(247,261,20)]:
        strips.alpha_composite(src.crop((0,y0,640,y1)),(dx,y0))
    blast=None
    if a.blast:
        blast=Image.open(a.blast).convert('RGBA')
        if blast.size!=src.size:raise ValueError('Blast output must retain canvas size.')
        before=np.asarray(strips);after=np.asarray(blast);changed=np.any(before!=after,axis=2)
        mask=np.zeros(360,dtype=bool);mask[90:101]=mask[166:174]=mask[247:261]=True
        if not changed.any() or changed[~mask].any():raise ValueError('Expected a changed horizontal-strip layer only; inspect editor output.')
    a.outdir.mkdir(parents=True);src.save(a.outdir/'channels.png');strips.save(a.outdir/'strip-copy.png')
    if blast:
        blast.save(a.outdir/'blast-right.png');merged=src.copy();merged.alpha_composite(blast)
        base=Image.new('RGBA',src.size,'#10141E');base.alpha_composite(merged);base.convert('RGB').save(a.outdir/'blast-base.png')
    y,x=np.mgrid[:360,:640]
    for name,cx,cy,color in [('magenta',135,95,(244,28,168)),('cyan',485,265,(0,220,239))]:
        field=np.empty((360,640,4),dtype='uint8');field[:,:,:3]=color
        field[:,:,3]=np.floor(255*.72*np.exp(-((x-cx)**2+(y-cy)**2)/(2*180**2))+.5).astype('uint8')
        Image.fromarray(field).save(a.outdir/(name+'-brush.png'))
    (a.outdir/'preparation.json').write_text(json.dumps({'scope':'Original 640x360 fixture preparation only. Supplied Blast requires actual editor operation; this script does not prove its method or direction.','backend':'external-render','bands':[[90,101,28],[166,174,-36],[247,261,20]],'brushStack':'base, magenta Overlay100%, cyan Overlay100%; do not merge the brush colors first','hashes':{f.name:hashlib.sha256(f.read_bytes()).hexdigest() for f in a.outdir.glob('*.png')}},indent=2),encoding='utf-8')

if __name__=='__main__':main()
