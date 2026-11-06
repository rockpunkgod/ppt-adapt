"""OfficeCLI stroke composition with explicit Gaussian-alpha assets.

Reuses the original four-stroke geometry. Blurred layers are PNG; crisp
strokes and the source shapes remain editable. Not an Artistic Blur preset.
"""
import argparse,hashlib,json,math,subprocess,xml.etree.ElementTree as E
from pathlib import Path
import numpy as np
from PIL import Image,ImageDraw
from stroke_depth_geometry import STROKES,CURVED_LEFT

def main():
 p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--outdir',type=Path,required=True);p.add_argument('--sigma',type=float,default=12);a=p.parse_args()
 if a.outdir.exists():raise FileExistsError(a.outdir)
 if not math.isfinite(a.sigma) or not 2<=a.sigma<=24:raise ValueError('Sigma must be 2..24 source pixels')
 a.outdir.mkdir(parents=True);assets=a.outdir/'assets';assets.mkdir();pad=math.ceil(4*a.sigma);size=600+2*pad
 axis=np.arange(-pad,pad+1,dtype=float);kernel=np.exp(-axis**2/(2*a.sigma*a.sigma));kernel/=kernel.sum()
 masks={};model=[]
 for role,names in [('whole',list(STROKES)),('left',['left'])]:
  large=Image.new('L',(size*4,size*4));draw=ImageDraw.Draw(large)
  for name in names:draw.polygon([((x+pad)*4,(y+pad)*4) for x,y in STROKES[name]],fill=255)
  original=large.resize((size,size),Image.Resampling.LANCZOS);source=np.asarray(original,dtype=float)/255
  blurred=np.apply_along_axis(lambda line:np.convolve(line,kernel,'same'),1,source)
  blurred=np.apply_along_axis(lambda line:np.convolve(line,kernel,'same'),0,blurred)
  alpha=Image.fromarray(np.rint(np.clip(blurred,0,1)*255).astype('uint8'))
  original.save(assets/(role+'-source-alpha.png'));alpha.save(assets/(role+'-gaussian-alpha.png'));masks[role]=(source,blurred)
  for label,rgb in [('ink',(23,42,53)),('white',(255,255,255))]:
   rgba=Image.new('RGBA',(size,size),rgb+(0,));rgba.putalpha(alpha);rgba.save(assets/f'{role}-{label}.png')
  model.append({'role':role,'strokes':names,'sourceMask':role+'-source-alpha.png','filteredMask':role+'-gaussian-alpha.png','sourceAlphaMass':float(source.sum()),'floatFilteredAlphaMass':float(blurred.sum())})
 deck=a.outdir/'glyph-gaussian.pptx'
 def cli(*args):
  r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8')
  if r.returncode:raise RuntimeError(r.stdout+r.stderr)
  return r.stdout
 def shape(s,name,x,y,w,h,text='',size=22,color='172A35',fill='none'):
  return {'command':'add','parent':f'/slide[{s}]','type':'shape','props':dict(name=name,x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm',text=text,font='Microsoft YaHei',size=str(size),color=color,fill=fill,line='none',margin='0cm')}
 def geometry(name):
  ns='http://schemas.openxmlformats.org/drawingml/2006/main';q=lambda t:'{'+ns+'}'+t
  root=E.Element(q('custGeom'))
  for n in ['avLst','gdLst','ahLst','cxnLst']:E.SubElement(root,q(n))
  E.SubElement(root,q('rect'),l='0',t='0',r='w',b='h');path=E.SubElement(E.SubElement(root,q('pathLst')),q('path'),w='600000',h='600000',stroke='false')
  commands=CURVED_LEFT if name=='left' else [('M' if i==0 else 'L',point) for i,point in enumerate(STROKES[name])]+[('Z',())]
  for kind,points in commands:
   node=E.SubElement(path,q({'M':'moveTo','L':'lnTo','C':'cubicBezTo','Z':'close'}[kind]))
   for k in range(0,len(points),2):E.SubElement(node,q('pt'),x=str(round(points[k]*1000)),y=str(round(points[k+1]*1000)))
  return E.tostring(root,encoding='unicode')
 cli('create',deck);cli('set',deck,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
 placements=[]
 for s in [1,2,3]:
  dark=s==3;ink='FFFFFF' if dark else '172A35';bg='101E25' if dark else 'F4F0E8'
  cli('add',deck,'/','--type','slide','--prop','layout=blank','--prop','background='+bg)
  title={1:'让一个字退后，让另一个字清晰',2:'只虚化一笔，保留字的骨架',3:'虚实之间，保持清楚的阅读顺序'}[s]
  ops=[shape(s,'title',1.5,.8,30,1.7,title,29,ink),shape(s,'left-label',2,3,12,1,'完整可编辑轮廓' if s<3 else '全字虚化',17,ink),shape(s,'right-label',19,3,12,1,'全字 Gaussian' if s==1 else '仅左撇 Gaussian',17,ink),shape(s,'original-text',1.5,17,30,1,'木木  /  原创四笔轮廓 · 原文字保留',14,ink)]
  for side,x in [('reference',2),('effect',19)]:
   mode='whole' if s==1 and side=='effect' or s==3 and side=='reference' else 'left' if s>=2 and side=='effect' else 'none'
   crisp=[] if mode=='whole' else [n for n in STROKES if mode!='left' or n!='left']
   # Blurred selected stroke goes behind the remaining crisp contours.
   if mode!='none':
    factor=12/600;asset=assets/f'{mode}-{"white" if dark else "ink"}.png'
    ops.append({'command':'add','parent':f'/slide[{s}]','type':'picture','props':dict(name=side+'-gaussian',src=str(asset.resolve()),x=f'{x-pad*factor}cm',y=f'{4.5-pad*factor}cm',width=f'{size*factor}cm',height=f'{size*factor}cm')})
   for name in crisp:ops.append(shape(s,side+'-'+name,x,4.5,12,12,fill=ink))
   placements.append({'slide':s,'side':side,'x':x,'y':4.5,'size':12,'mode':mode,'crispStrokes':crisp,'ink':ink})
  f=a.outdir/'batch.json';f.write_text(json.dumps(ops,ensure_ascii=False),encoding='utf-8');result=json.loads(cli('batch',deck,'--input',f,'--json'));assert result['success']
  for side in ['reference','effect']:
   item=next(q for q in placements if q['slide']==s and q['side']==side)
   for name in item['crispStrokes']:cli('raw-set',deck,f'/slide[{s}]','--xpath',f'//p:sp[p:nvSpPr/p:cNvPr[@name="{side}-{name}"]]/p:spPr/a:prstGeom','--action','replace','--xml',geometry(name))
 cli('close',deck);validation=cli('validate',deck,'--json');(a.outdir/'validation.json').write_text(validation,encoding='utf-8')
 manifest={'sigmaSourcePixels':a.sigma,'sourceDomain':600,'paddingPixels':pad,'kernel':kernel.tolist(),'kernelOffsets':axis.tolist(),'kernelVariance':float((kernel*axis**2).sum()),'algorithm':'Normalized discrete Gaussian, separable x/y convolution of alpha, truncated at 4 sigma; constant RGB, unassociated alpha. Source mask is 4x supersampled original stroke contours. Not Adobe radius/algorithm equivalence.','layers':model,'placements':placements,'assets':{f.name:hashlib.sha256(f.read_bytes()).hexdigest() for f in assets.glob('*.png')}}
 (a.outdir/'gaussian-model.json').write_text(json.dumps(manifest,indent=2),encoding='utf-8');print(deck)

if __name__=='__main__':main()
