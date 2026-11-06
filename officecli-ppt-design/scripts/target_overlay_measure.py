"""Fixture-only native image measurement; does not infer editable-data support."""
import argparse,json
from pathlib import Path
import numpy as np
from PIL import Image
def measure(folder):
 pages=[];images=[]
 for page,target in [(1,100),(2,120)]:
  a=np.array(Image.open(folder/f'page-{page}.png').convert('RGB'));images.append(a)
  # Axis bounds come from the native black axis, independently of colour bars.
  black=(a[:,:,:3].max(axis=2)<30);region=black[200:720,120:220]
  axis_x=120+int(region.sum(axis=0).argmax());yy=np.where(black[200:720,axis_x])[0]+200
  top=int(yy.min());bottom=200+int(black[200:720,120:1500].sum(axis=1).argmax());assert bottom-top>400
  scale=(bottom-top)/140
  dark=(a==[28,101,97]).all(axis=2);pale=(a==[213,226,228]).all(axis=2)
  dark[:230]=False;dark[690:]=False;pale[:230]=False;pale[690:]=False
  xs=np.where(dark.sum(axis=0)>20)[0];groups=np.split(xs,np.where(np.diff(xs)>1)[0]+1);assert len(groups)==4
  bars=[]
  for group,value in zip(groups,[45,60,75,90]):
   x0,x1=int(group.min()),int(group.max());x=(x0+x1)//2
   dy=np.where(dark[:,x])[0];py=np.where(pale[:,x])[0];assert len(dy)>100 and len(py)>15
   actual_top=int(dy.min());target_top=int(py.min());base=int(dy.max())+1
   assert abs(base-bottom)<=1
   actual_height=base-actual_top;target_height=base-target_top
   assert abs(actual_height/scale-value)<0.5
   assert abs(target_height/scale-target)<0.5
   # Pale target and opaque actual must share the same x interval.
   target_x=np.where(pale[target_top+3])[0];local=target_x[(target_x>=x0-5)&(target_x<=x1+5)]
   assert int(local.min())==x0 and int(local.max())==x1
   bars.append({'actual':value,'actualHeightPx':actual_height,'targetHeightPx':target_height,'x':[x0,x1],'center':x,'zeroY':base,'actualTop':actual_top,'targetTop':target_top})
  pages.append({'page':page,'target':target,'axisY':[top,bottom],'pixelsPerUnit':scale,'bars':bars})
 for a,b in zip(pages[0]['bars'],pages[1]['bars']):
  assert a['x']==b['x'] and a['actualHeightPx']==b['actualHeightPx'] and a['zeroY']==b['zeroY']
  assert b['targetHeightPx']>a['targetHeightPx']
 # All opaque actual pixels in the plot are unchanged after target-only edit.
 for a,b in zip(images,images[1:]):
  mask=(a==[28,101,97]).all(axis=2);mask[:230]=False;mask[690:]=False
  assert np.array_equal(a[mask],b[mask])
 return {'passed':True,'pages':pages,'opaqueActualPixelsUnchanged':True,'scope':'1600x900 original fixed target-overlay fixture; not arbitrary chart styles.'}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('folder',type=Path);p.add_argument('--report',required=True,type=Path);a=p.parse_args()
 if a.report.exists():raise FileExistsError(a.report)
 a.report.write_text(json.dumps(measure(a.folder),indent=2),encoding='utf-8');print(a.report)
