"""Preserve the native Intersect path and image mapping while removing srcRect."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import argparse,xml.etree.ElementTree as E
p=argparse.ArgumentParser();p.add_argument('--input',type=Path,required=True);p.add_argument('--out',type=Path,required=True);a=p.parse_args()
if a.out.exists():raise FileExistsError(a.out)
n={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
with ZipFile(a.input) as z:data={v:z.read(v) for v in z.namelist()}
x=E.fromstring(data['ppt/slides/slide2.xml']);pics=[v for v in x.findall('.//p:pic',n) if v.find('p:nvPicPr/p:cNvPr',n).get('name')=='texture-gif']
if len(pics)!=1:raise ValueError('Expected the native Intersect output')
pic=pics[0]
if pic.find('p:spPr/a:custGeom',n) is None:raise ValueError('Do not substitute a rectangular GIF')
f=pic.find('p:blipFill',n);crop=f.find('a:srcRect',n)
if crop is None:raise ValueError('Expected native crop; no repair applied')
l,t,r,b=[int(crop.get(k,'0')) for k in ['l','t','r','b']];w=100000-l-r;h=100000-t-b
if min(l,t,r,b)<0 or w<=0 or h<=0:raise ValueError('Example covers positive, nonempty native crop only')
fill=f.find('a:stretch/a:fillRect',n)
if fill is None or fill.attrib:raise ValueError('Unexpected existing destination fill; preserve it rather than overwrite')
fill.attrib.update({k:str(round(v)) for k,v in dict(l=-100000*l/w,r=-100000*r/w,t=-100000*t/h,b=-100000*b/h).items()});f.remove(crop)
data['ppt/slides/slide2.xml']=E.tostring(x,encoding='utf-8',xml_declaration=True)
a.out.parent.mkdir(parents=True,exist_ok=True)
with ZipFile(a.out,'w',ZIP_DEFLATED) as z:
 for k,v in data.items():z.writestr(k,v)
print(dict(nativeGeometryRetained=True,fillRect=fill.attrib,output=str(a.out)))
