"""Original lamp mesh and abstract sensing scene, no downloaded assets."""
import argparse,json,math,struct,io
from pathlib import Path
from PIL import Image,ImageDraw
p=argparse.ArgumentParser();p.add_argument('--outdir',type=Path,required=True);args=p.parse_args()
o=args.outdir;o.mkdir(parents=True,exist_ok=False)
verts=[];norms=[];colors=[]
def tri(a,b,c,color):
 u=[b[i]-a[i] for i in range(3)];v=[c[i]-a[i] for i in range(3)]
 n=[u[1]*v[2]-u[2]*v[1],u[2]*v[0]-u[0]*v[2],u[0]*v[1]-u[1]*v[0]];m=math.sqrt(sum(q*q for q in n));n=[q/m for q in n]
 for p in (a,b,c):verts.extend(p);norms.extend(n);colors.extend(color)
def box(cx,cy,cz,w,h,d,col):
 p=[(cx+x*w/2,cy+y*h/2,cz+z*d/2) for x,y,z in [(-1,-1,-1),(1,-1,-1),(1,1,-1),(-1,1,-1),(-1,-1,1),(1,-1,1),(1,1,1),(-1,1,1)]]
 for ids in [(0,3,2,1),(4,5,6,7),(0,4,7,3),(1,2,6,5),(0,1,5,4),(3,7,6,2)]:
  a,b,c,d=[p[i] for i in ids];tri(a,b,c,col);tri(a,c,d,col)
blue=(.08,.30,.43);metal=(.34,.45,.50);lens=(1,.45,.02)
# Lamp axis +Y; faceted cylindrical housing, lens at upper emitting end.
for k in range(32):
 a=k*math.tau/32;b=(k+1)*math.tau/32
 p=lambda t,y:(.42*math.cos(t),y,.42*math.sin(t))
 tri(p(a,.1),p(b,.1),p(b,.95),blue);tri(p(a,.1),p(b,.95),p(a,.95),blue)
 tri((0,.95,0),p(a,.95),p(b,.95),lens);tri((0,.1,0),p(b,.1),p(a,.1),metal)
box(0,-.08,0,.75,.25,.28,metal);box(0,-.65,0,.12,1.05,.12,metal)
box(0,-1.18,0,1.25,.15,.65,blue)
blob=bytearray();views=[];access=[];prims=[];materials=[]
for col in (blue,metal,lens):
 ids=[k for k in range(len(verts)//3) if tuple(colors[k*3:k*3+3])==col]
 positions=[v for k in ids for v in verts[k*3:k*3+3]];normals=[v for k in ids for v in norms[k*3:k*3+3]]
 start=len(access)
 for vals in (positions,normals):
  data=struct.pack('<'+'f'*len(vals),*vals);views.append(dict(buffer=0,byteOffset=len(blob),byteLength=len(data),target=34962));blob.extend(data)
  item=dict(bufferView=len(views)-1,componentType=5126,count=len(vals)//3,type='VEC3')
  if vals is positions:item.update(min=[min(vals[i::3]) for i in range(3)],max=[max(vals[i::3]) for i in range(3)])
  access.append(item)
 prims.append(dict(attributes=dict(POSITION=start,NORMAL=start+1),material=len(materials),mode=4))
 materials.append(dict(doubleSided=True,pbrMetallicRoughness=dict(baseColorFactor=[*col,1],metallicFactor=.15,roughnessFactor=.65)))
images=[]
for k,col in enumerate((blue,metal,lens)):
 rgba=Image.new('RGBA',(2,2),tuple(round(c*255) for c in col)+(255,));b=io.BytesIO();rgba.save(b,format='PNG');data=b.getvalue()
 while len(blob)%4:blob.append(0)
 images.append(dict(bufferView=len(views),mimeType='image/png'));views.append(dict(buffer=0,byteOffset=len(blob),byteLength=len(data)));blob.extend(data)
 materials[k]['pbrMetallicRoughness'].update(baseColorFactor=[1,1,1,1],baseColorTexture=dict(index=k))
 # Constant UV coordinates make these textures unambiguous on every triangle.
 uv=[0.5]*(access[prims[k]['attributes']['POSITION']]['count']*2);data=struct.pack('<'+'f'*len(uv),*uv)
 while len(blob)%4:blob.append(0)
 views.append(dict(buffer=0,byteOffset=len(blob),byteLength=len(data),target=34962));blob.extend(data)
 prims[k]['attributes']['TEXCOORD_0']=len(access);access.append(dict(bufferView=len(views)-1,componentType=5126,count=len(uv)//2,type='VEC2'))
g=dict(asset=dict(version='2.0',generator='Original LUMA lamp'),scene=0,scenes=[dict(nodes=[0])],nodes=[dict(mesh=0)],meshes=[dict(primitives=prims)],materials=materials,images=images,textures=[dict(source=k) for k in range(3)],buffers=[dict(byteLength=len(blob))],bufferViews=views,accessors=access)
j=json.dumps(g,separators=(',',':')).encode();j+=b' '*(-len(j)%4);blob+=b'\0'*(-len(blob)%4)
(o/'lamp.glb').write_bytes(struct.pack('<III',0x46546c67,2,28+len(j)+len(blob))+struct.pack('<II',len(j),0x4e4f534a)+j+struct.pack('<II',len(blob),0x004e4942)+blob)
im=Image.new('RGB',(1600,900),(24,38,43));d=ImageDraw.Draw(im)
for x in [360,800,1240]:
 d.line((x,250,x,650),fill=(46,69,75),width=1)
 d.ellipse((x-105,295,x+105,505),fill=(57,103,111),outline=(119,203,198),width=4)
 d.ellipse((x-65,335,x+65,465),outline=(204,225,203),width=3)
 for k in range(12):
  t=k*math.tau/12;d.line((x+78*math.cos(t),400+78*math.sin(t),x+91*math.cos(t),400+91*math.sin(t)),fill=(222,176,103),width=3)
d.line((465,400,695,400),fill=(119,160,164),width=3);d.line((905,400,1135,400),fill=(119,160,164),width=3)
im.save(o/'scene.png');print(o)
