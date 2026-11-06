"""Independent OPC, native angle, true-gap and label consistency audit."""
import argparse,json,zipfile,posixpath,xml.etree.ElementTree as E
from pathlib import Path
import numpy as np
from PIL import Image
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
PAL=np.array([[28,101,97],[96,127,154],[213,173,110],[150,132,174]])
def structure(deck):
 pages=[];parts=[]
 with zipfile.ZipFile(deck) as z:
  for i in [1,2,3]:
   root=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'));rels={q.get('Id'):q.get('Target') for q in E.fromstring(z.read(f'ppt/slides/_rels/slide{i}.xml.rels'))};charts=[]
   for gf in root.findall('p:cSld/p:spTree/p:graphicFrame',N):
    cref=gf.find('.//c:chart',N)
    if cref is None:continue
    part=posixpath.normpath(posixpath.join('ppt/slides',rels[cref.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')])).lstrip('/');parts.append(part)
    cr=E.fromstring(z.read(part));plot=cr.find('c:chart/c:plotArea',N);nodes=plot.findall('c:doughnutChart',N)+plot.findall('c:pieChart',N);assert len(nodes)==1
    node=nodes[0];kind=node.tag.split('}')[-1];ser=node.findall('c:ser',N);assert len(ser)==1
    vals=[float(q.findtext('c:v',namespaces=N)) for q in ser[0].findall('c:val/c:numLit/c:pt',N)];assert abs(sum(vals)-100)<1e-8
    assert node.find('c:varyColors',N).get('val')=='1' and node.find('c:firstSliceAng',N).get('val')=='0'
    assert cr.find('c:externalData',N) is None
    if kind=='doughnutChart':assert node.find('c:holeSize',N).get('val')==('85' if i<3 else '68')
    transform=gf.find('p:xfrm',N);off=transform.find('a:off',N);ext=transform.find('a:ext',N);x,y,cx,cy=[int(v) for v in [off.get('x'),off.get('y'),ext.get('cx'),ext.get('cy')]]
    charts.append({'part':part,'kind':kind,'values':vals,'centerEMU':[x+cx/2,y+cy/2]})
   assert len(charts)==(2 if i<3 else 3)
   texts={sp.find('p:nvSpPr/p:cNvPr',N).get('name'):''.join(sp.itertext()) for sp in root.findall('p:cSld/p:spTree/p:sp',N)}
   if i<3:
    expected=[55,20,15,10] if i==1 else [40,30,20,10]
    assert charts[0]['kind']=='doughnutChart' and charts[1]['kind']=='pieChart'
    assert all(c['values']==expected for c in charts)
    assert charts[0]['centerEMU']==charts[1]['centerEMU']
    for j,v in enumerate(expected):assert texts[f'value-{j}']==f'{v}%'
   else:
    for j,v in enumerate([65,25,10]):assert charts[j]['values']==[v,100-v] and texts[f'card-value-{j}']==f'{v}%'
   pages.append({'page':i,'charts':charts})
 assert len(parts)==len(set(parts))==7,'Independent chart parts needed'
 return pages
def classify(rgb,palette=PAL):
 delta=np.abs(rgb.astype(int)[:,None,:]-palette[None,:,:]).max(axis=2);best=delta.argmin(axis=1);best[delta.min(axis=1)>18]=-1;return best
def rays(a,c,r):
 theta=np.arange(720)*np.pi/360;xx=np.rint(c[0]+r*np.sin(theta)).astype(int);yy=np.rint(c[1]-r*np.cos(theta)).astype(int);return a[yy,xx]
def proportions(labels,n):
 valid=labels>=0;assert valid.mean()>.99
 return [float((labels[valid]==j).mean()) for j in range(n)]
def pixels(folder,pages):
 checks=[]
 for i in [1,2]:
  a=np.array(Image.open(folder/f'page-{i}.png').convert('RGB'));region=a[200:800,100:780];mask=np.any(np.all(region[:,:,None,:]==PAL[None,None,:,:],axis=3),axis=2);yy,xx=np.where(mask)
  c=[(float(xx.min()+xx.max())/2)+100,(float(yy.min()+yy.max())/2)+200];r=(float(xx.max()-xx.min())+float(yy.max()-yy.min()))/4
  outer=classify(rays(a,c,r*.94));inner=classify(rays(a,c,r*.55));op=proportions(outer,4);ip=proportions(inner,4)
  target=np.array(pages[i-1]['charts'][0]['values'])/100
  assert np.max(np.abs(np.array(op)-target))<.006 and np.max(np.abs(np.array(ip)-target))<.006
  good=(outer>=0)&(inner>=0);agreement=float((outer[good]==inner[good]).mean());assert agreement>.99
  # A real empty annulus must show two different native backgrounds.
  gap=classify(rays(a,c,r*.8),np.array([[230,239,237],[238,223,203]]));assert (gap>=0).mean()>.99
  assert abs((gap==0).mean()-.5)<.02 and abs((gap==1).mean()-.5)<.02
  checks.append({'page':i,'centerPx':c,'outerRadiusPx':r,'outerFractions':op,'innerFractions':ip,'angularColorAgreement':agreement,'trueTwoBackgroundGap':True})
 a=np.array(Image.open(folder/'page-3.png').convert('RGB'))
 for j,v in enumerate([65,25,10]):
  x0=round((1.8+j*10.6)*1600/33.87);region=a[340:730,x0:x0+445];mask=np.any(np.all(region[:,:,None,:]==PAL[None,None,:2,:],axis=3),axis=2);yy,xx=np.where(mask)
  c=[float(xx.min()+xx.max())/2+x0,float(yy.min()+yy.max())/2+340];r=(float(xx.max()-xx.min())+float(yy.max()-yy.min()))/4
  values=proportions(classify(rays(a,c,r*.84),PAL[:2]),2);assert abs(values[0]-v/100)<.006
  checks.append({'card':j+1,'share':v,'centerPx':c,'radiusPx':r,'fractions':values})
 return {'passed':True,'checks':checks,'scope':'Original fixed 3-page diagnostic; no workbook/Edit Data claim.'}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('deck');p.add_argument('--native',type=Path,required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
 if a.report.exists():raise FileExistsError(a.report)
 pages=structure(a.deck);result={'passed':True,'structure':pages,'pixels':pixels(a.native,pages)};a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(a.report)
