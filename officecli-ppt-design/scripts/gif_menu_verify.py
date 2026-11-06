import argparse,hashlib,json,posixpath,struct,zipfile,xml.etree.ElementTree as E
from pathlib import Path
def gif_info(data):
 assert data[:6] in [b'GIF87a',b'GIF89a'];packed=data[10];i=13+(3*(2**((packed&7)+1)) if packed&128 else 0);frames=[];delay=0;transparent=False;loops=None
 def blocks(i):
  out=b''
  while data[i]:n=data[i];out+=data[i+1:i+1+n];i+=1+n
  return out,i+1
 while i<len(data):
  kind=data[i];i+=1
  if kind==0x3b:break
  if kind==0x21:
   label=data[i];i+=1;payload,i=blocks(i)
   if label==0xf9:delay=int.from_bytes(payload[1:3],'little')/100;transparent=bool(payload[0]&1)
   if label==0xff and payload.startswith(b'NETSCAPE2.0'):loops=int.from_bytes(payload[-2:],'little')
  elif kind==0x2c:
   packed=data[i+8];i+=9
   if packed&128:i+=3*2**((packed&7)+1)
   i+=1;_,i=blocks(i);frames.append({'delay':delay,'transparent':transparent})
  else:raise ValueError((kind,i))
 return {'frames':len(frames),'duration':sum(f['delay'] for f in frames),'delays':[f['delay'] for f in frames],'transparent':all(f['transparent'] for f in frames),'loopCount':loops}
def verify(deck,assets):
 ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships','rel':'http://schemas.openxmlformats.org/package/2006/relationships'};checks={};states=[]
 with zipfile.ZipFile(deck) as z:
  for s in range(2,5):
   root=E.fromstring(z.read(f'ppt/slides/slide{s}.xml'));rels={x.get('Id'):x.get('Target') for x in E.fromstring(z.read(f'ppt/slides/_rels/slide{s}.xml.rels'))};tree=root.find('p:cSld/p:spTree',ns);pictures={};order=[]
   for q in tree:
    n=q.find('.//p:cNvPr',ns)
    if n is None:continue
    name=n.get('name');order.append(name)
    if q.tag.endswith('}pic'):
     rid=q.find('.//a:blip',ns).get('{'+ns['r']+'}embed');target=rels[rid];path=target.lstrip('/') if target.startswith('/') else posixpath.normpath('ppt/slides/'+target);pictures[name]=z.read(path)
   selected=s-1
   for icon in range(1,4):
    data=pictures[f'!!icon-{icon}'];src=assets/(f'icon-{icon}-active.gif' if icon==selected else f'icon-{icon}-inactive.png');checks[f's{s}icon{icon}originalBytes']=data==src.read_bytes()
    if icon==selected:
     info=gif_info(data);checks[f's{s}actual12FrameTransparentLoop']=info['frames']==12 and abs(info['duration']-2.4)<0.001 and info['loopCount']==0 and info['transparent'];states.append(info)
   checks[f's{s}topmostHitShapes']=all(order.index(f'!!icon-{i}')<order.index(f'!!circle-{j}') for i in range(1,4) for j in range(1,4))
   checks[f's{s}threeLinks']=len(root.findall('.//a:hlinkClick',ns))==3
 return {'passed':all(checks.values()),'checks':checks,'gifStates':states,'note':'Independent GIF block parse and exact original-byte equality, not a native playback assertion.'}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--deck',type=Path,required=True);p.add_argument('--assets',type=Path,required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
 if a.report.exists():raise FileExistsError(a.report)
 result=verify(a.deck,a.assets);a.report.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result,indent=2))
