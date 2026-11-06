import argparse,json,subprocess,hashlib,zipfile,posixpath,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--assets',type=Path,required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--deck',type=Path,required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
W,H=640,360
def decode(f):return subprocess.run([a.ffmpeg,'-v','error','-i',str(f),'-f','rawvideo','-pix_fmt','rgba','-'],capture_output=True,check=True).stdout
names=['original.png','channels.png','slices.png','streak-approx.png'];data=[decode(a.assets/n) for n in names];src,ch,sl,wind=data
def pixel(d,x,y):return tuple(d[(y*W+x)*4:(y*W+x+1)*4]) if 0<=x<W and 0<=y<H else (0,0,0,0)
channelbad=0;alphabad=0;slicebad=0;unchanged=0;changed=0
bands=[(90,101,28),(166,174,-36),(247,261,20)]
for y in range(H):
 for x in range(W):
  r=pixel(src,x-10,y);gb=pixel(src,x+10,y);v=pixel(ch,x,y);alpha=max(r[3],gb[3]);alphabad+=v[3]!=alpha
  expected=(round(r[0]*r[3]/alpha),round(gb[1]*gb[3]/alpha),round(gb[2]*gb[3]/alpha),alpha) if alpha else (0,0,0,0);channelbad+=v!=expected
  band=next((b for b in bands if b[0]<=y<b[1]),None);actual=pixel(sl,x,y)
  if not band:unchanged+=actual!=v
  else:
   moved=pixel(ch,x-band[2],y);sa=moved[3]/255;da=v[3]/255;oa=sa+da*(1-sa);expected=tuple(round((moved[c]*sa+v[c]*da*(1-sa))/oa) for c in range(3))+(round(oa*255),) if oa else (0,0,0,0);slicebad+=actual!=expected;changed+=actual!=v
checks={'RGBA640x360':all(len(d)==W*H*4 for d in data),'exactIndependentChannelSampling':channelbad==0,'unionAlphaPreserved':alphabad==0,'semiTransparentInputPresent':any(src[j]==128 for j in range(3,len(src),4)),'transparentInputPresent':any(src[j]==0 for j in range(3,len(src),4)),'outsideBandsByteIdentical':unchanged==0,'insideBandsCopyOverCorrect':slicebad==0,'bandsActuallyChangeImage':changed>3000,'streakOnlyDeclaredRows':all(wind[(y*W)*4:((y+1)*W)*4]==sl[(y*W)*4:((y+1)*W)*4] for y in range(H) if not(80<=y<280 and y%9<2)),'streakHasVisibleChange':wind!=sl}
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'};R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'
with zipfile.ZipFile(a.deck) as z:
 for i,name in enumerate(names,1):
  root=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'));pics=root.findall('.//p:pic',ns);checks[f's{i}OneRealPicture']=len(pics)==1
  rels={r.get('Id'):r.get('Target') for r in E.fromstring(z.read(f'ppt/slides/_rels/slide{i}.xml.rels'))};rid=pics[0].find('.//a:blip',ns).get(R);target=rels[rid];part=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('ppt/slides',target));checks[f's{i}EmbeddedBytesUnchanged']=z.read(part)==(a.assets/name).read_bytes()
r={'passed':all(checks.values()),'checks':checks,'mismatches':{'channels':channelbad,'alpha':alphabad,'bands':slicebad,'outsideBands':unchanged},'changedBandPixels':changed,'hashes':{n:hashlib.sha256((a.assets/n).read_bytes()).hexdigest() for n in names},'scope':'True image channel separation and selected-band copy offset; fourth page is declared custom streak approximation, not Adobe Wind algorithm verification.'};a.report.write_text(json.dumps(r,indent=2),encoding='utf-8');print(json.dumps(r,indent=2))
