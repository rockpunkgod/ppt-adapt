"""Check true chart series fill/data semantics and fixed native-fixture pixels."""
import argparse,json,posixpath,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import numpy as np
from PIL import Image
N={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def charts(deck):
 out=[]
 with zipfile.ZipFile(deck) as z:
  for page in [1,2]:
   path=f'ppt/slides/_rels/slide{page}.xml.rels';rels=E.fromstring(z.read(path));targets=[q.get('Target') for q in rels if q.get('Type','').endswith('/chart')];assert len(targets)==1
   part=posixpath.normpath(posixpath.join('ppt/slides',targets[0]));root=E.fromstring(z.read(part));ser=root.findall('.//c:barChart/c:ser',N);assert len(ser)==1
   vals=[float(q.findtext('c:v',namespaces=N)) for q in ser[0].findall('c:val/c:numLit/c:pt',N)]
   if not vals:vals=[float(q.findtext('c:v',namespaces=N)) for q in ser[0].findall('c:val/c:numRef/c:numCache/c:pt',N)]
   assert vals==[45,60,75,90]
   assert root.find('.//c:valAx/c:scaling/c:min',N).get('val')=='0'
   assert root.find('.//c:valAx/c:scaling/c:max',N).get('val')=='100'
   fill=ser[0].find('c:spPr',N)
   if page==1:
    grad=fill.find('a:gradFill',N);assert grad is not None
    assert [(q.get('pos'),q.find('a:srgbClr',N).get('val')) for q in grad.findall('a:gsLst/a:gs',N)]==[('0','28D9E7'),('100000','102338')]
    assert grad.find('a:lin',N).get('ang')=='5400000'
   else:
    patt=fill.find('a:pattFill',N);assert patt is not None and patt.get('prst')=='narHorz'
    assert patt.find('a:fgClr/a:srgbClr',N).get('val')=='78DB65'
    assert patt.find('a:bgClr/a:srgbClr',N).get('val')=='102338'
   out.append({'page':page,'chart':part,'data':vals,'fill':'gradient' if page==1 else 'narHorz'})
 return out
def pixels(folder):
 grad=np.array(Image.open(folder/'gradient.png').convert('RGB'));pattern=np.array(Image.open(folder/'pattern.png').convert('RGB'))
 black=grad.max(axis=2)<30;zero=230+int(black[230:740,120:1500].sum(axis=1).argmax());x=120+int(black[230:740,120:220].sum(axis=0).argmax());top=230+int(np.where(black[230:740,x])[0].min());scale=(zero-top)/100
 fg=(pattern==[120,219,101]).all(axis=2);fg[:260]=False;fg[zero:]=False
 xs=np.where(fg.sum(axis=0)>20)[0];groups=np.split(xs,np.where(np.diff(xs)>1)[0]+1);assert len(groups)==4
 profiles=[]
 for group,value in zip(groups,[45,60,75,90]):
  center=int((group.min()+group.max())//2);yy=np.where(fg[:,center])[0];actual_top=int(yy.min())
  # The first foreground stripe may follow the actual bar edge by one pattern
  # period; infer that period from native rows rather than treating it as height.
  period=int(np.max(np.diff(yy)));assert 2<=period<=4
  expected_top=zero-value*scale
  assert -1<=actual_top-expected_top<=period+1
  samples=[grad[round(zero-value*scale+f*value*scale),center].astype(int).tolist() for f in [.1,.35,.65,.9]];profiles.append(samples)
  # Actual horizontal pattern has many alternating foreground/background runs.
  flags=fg[actual_top:zero,center];switches=int(np.count_nonzero(flags[1:]!=flags[:-1]));assert switches>value
  near=fg[actual_top+5:zero-5,int(group.min())+5:int(group.max())-5]
  assert np.all(near==near[:,0,None]),'Pattern must be horizontal inside each bar'
 assert np.abs(np.array(profiles)-np.array(profiles[0])).max()<=3,'Gradient normalized to individual bars'
 assert profiles[0][0][1]>profiles[0][-1][1]+100
 reopen=[]
 for i,img in [(1,grad),(2,pattern)]:
  same=np.array_equal(img,np.array(Image.open(folder/f'reopen-{i}.png').convert('RGB')));assert same;reopen.append(same)
 return {'passed':True,'axisY':[top,zero],'barGradientProfiles':profiles,'reopenPixelEqual':reopen}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('folder',type=Path);p.add_argument('--report',required=True,type=Path);a=p.parse_args()
 if a.report.exists():raise FileExistsError(a.report)
 result={'passed':True,'charts':charts(a.folder/'styled.pptx'),'pixels':pixels(a.folder)};a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(a.report)
