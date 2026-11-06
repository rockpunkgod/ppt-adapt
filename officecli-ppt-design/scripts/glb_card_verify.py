"""Independent binary/mesh checks, not a native-render success claim."""
import argparse,json,struct,hashlib,zipfile,math,io
from pathlib import Path
from PIL import Image
p=argparse.ArgumentParser();p.add_argument('folder');a=p.parse_args();o=Path(a.folder);b=(o/'card.glb').read_bytes();source=json.loads((o/'texture-source.json').read_text(encoding='utf-8')) if (o/'texture-source.json').exists() else {'kind':'diagnostic'}
magic,ver,total=struct.unpack_from('<III',b);assert (magic,ver,total)==(0x46546c67,2,len(b))
n,t=struct.unpack_from('<II',b,12);assert t==0x4e4f534a and n%4==0;g=json.loads(b[20:20+n]);nb,tb=struct.unpack_from('<II',b,20+n);assert tb==0x004e4942 and nb%4==0;blob=b[28+n:];assert len(blob)==nb
assert g['asset']['version']=='2.0' and len(g['buffers'])==1 and g['buffers'][0]['byteLength']<=nb
assert all(v.get('byteOffset',0)%4==0 and v.get('byteOffset',0)+v['byteLength']<=nb for v in g['bufferViews'])
assert len(g['meshes'])==1 and len(g['meshes'][0]['primitives'])==2 and len(g['images'])==2
assert not any('uri' in q for q in g['buffers']+g['images'])
def array(idx):
 ac=g['accessors'][idx];v=g['bufferViews'][ac['bufferView']];components={'VEC3':3,'VEC2':2,'SCALAR':1}[ac['type']];fmt={5126:'f',5123:'H'}[ac['componentType']];count=ac['count']*components
 vals=struct.unpack_from('<'+fmt*count,blob,v.get('byteOffset',0)+ac.get('byteOffset',0));return [vals[i:i+components] for i in range(0,len(vals),components)]
rows=[]
for side,pr in enumerate(g['meshes'][0]['primitives']):
 assert pr['mode']==4 and pr['material']==side
 mat=g['materials'][side];assert mat['doubleSided'] is False and mat['alphaMode']=='MASK' and mat['alphaCutoff']==0.5
 assert mat['pbrMetallicRoughness']['baseColorTexture']['index']==side
 ps=array(pr['attributes']['POSITION']);norm=array(pr['attributes']['NORMAL']);uv=array(pr['attributes']['TEXCOORD_0']);ind=[v[0] for v in array(pr['indices'])]
 assert len(ps)==4 and len(ind)==6 and max(ind)<4 and all(n==(0,0,1 if side==0 else -1) for n in norm)
 assert uv==([(0.,1.),(1.,1.),(1.,0.),(0.,0.)] if side==0 else [(1.,1.),(0.,1.),(0.,0.),(1.,0.)])
 for j in range(0,6,3):
  pa,pb,pc=[ps[v] for v in ind[j:j+3]];u=[pb[k]-pa[k] for k in range(3)];v=[pc[k]-pa[k] for k in range(3)]
  cross=(u[1]*v[2]-u[2]*v[1],u[2]*v[0]-u[0]*v[2],u[0]*v[1]-u[1]*v[0]);assert sum(cross[k]*norm[0][k] for k in range(3))>0
 view=g['bufferViews'][g['images'][side]['bufferView']];data=blob[view.get('byteOffset',0):view.get('byteOffset',0)+view['byteLength']];assert data==(o/('front.png' if side==0 else 'back.png')).read_bytes()
 im=Image.open(io.BytesIO(data));assert im.mode=='RGBA'
 if source['kind']=='diagnostic':assert im.getpixel((0,0))[3]==0 and im.getpixel((180,320))[3]==255
 rows.append({'side':side,'normal':norm[0],'triangles':2,'pngSha256':hashlib.sha256(data).hexdigest(),'uv':uv})
with zipfile.ZipFile(o/'card.pptx') as z:
 for i,angle in enumerate([0,180,0],1):
  import xml.etree.ElementTree as E
  ns={'m':'http://schemas.microsoft.com/office/drawing/2017/model3d','p159':'http://schemas.microsoft.com/office/powerpoint/2015/09/main'}
  r=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'));m=r.findall('.//m:model3d',ns);assert len(m)==1
  rot=m[0].find('m:trans/m:rot',ns);assert int(rot.get('ay'))==angle*60000 and int(rot.get('ax'))==0
  assert r.find('.//p159:morph',ns) is not None
(o/'structural-audit.json').write_text(json.dumps({'passed':True,'nativePlaybackVerifiedByThisScript':False,'faces':rows,'glbSha256':hashlib.sha256(b).hexdigest()},indent=2),encoding='utf-8');print('GLB mesh/texture and OPC rotation checks passed; native pending')
