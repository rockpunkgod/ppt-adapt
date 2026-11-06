"""Original local double-sided glTF diagnostic; no author template or model."""
import argparse,json,struct,subprocess,hashlib,zipfile,shutil
from pathlib import Path
from PIL import Image,ImageDraw
p=argparse.ArgumentParser();p.add_argument('--out',required=True);p.add_argument('--officecli',required=True);p.add_argument('--front');p.add_argument('--back');a=p.parse_args();o=Path(a.out)
if o.exists():raise FileExistsError('Refusing overwrite')
if bool(a.front)!=bool(a.back):raise ValueError('Supply both front and back PNG or neither')
if a.front:
 for f in [a.front,a.back]:
  with Image.open(f) as im:assert im.format=='PNG' and im.mode=='RGBA','Caller textures: RGBA PNG, prepared locally'
 with Image.open(a.front) as f,Image.open(a.back) as b:assert f.size==b.size and f.width/f.height==360/640,'Tested card aspect is 9:16; matching PNG dimensions required'
o.mkdir(parents=True)
for side,col,letter in [('front',(26,151,143),'F'),('back',(234,117,49),'B')]:
 if a.front:
  shutil.copyfile(a.front if side=='front' else a.back,o/(side+'.png'));continue
 im=Image.new('RGBA',(360,640));d=ImageDraw.Draw(im);d.rounded_rectangle((1,1,358,638),radius=36,fill=(*col,255))
 # Asymmetric corner cues prove texture orientation rather than a repeated face.
 d.rectangle((28,35,100,107),fill=(245,245,245,255));d.ellipse((260,535,330,605),fill=(15,34,53,255))
 d.text((125,230),letter,fill='white',font_size=170,stroke_width=3)
 d.text((60,460),side.upper(),fill='white',font_size=42)
 im.save(o/(side+'.png'))
(o/'texture-source.json').write_text(json.dumps({'kind':'caller' if a.front else 'diagnostic','sha256':{s:hashlib.sha256((o/(s+'.png')).read_bytes()).hexdigest() for s in ['front','back']}},indent=2),encoding='utf-8')
blob=bytearray();views=[];acc=[]
def view(data,target=None):
 while len(blob)%4:blob.append(0)
 v={'buffer':0,'byteOffset':len(blob),'byteLength':len(data)}
 if target:v['target']=target
 views.append(v);blob.extend(data);return len(views)-1
def accessor(vals,typ,component=5126,mins=None,maxs=None):
 n={'VEC3':3,'VEC2':2,'SCALAR':1}[typ];fmt='f' if component==5126 else 'H';v=view(struct.pack('<'+fmt*len(vals),*vals),34963 if component==5123 else 34962)
 q={'bufferView':v,'componentType':component,'count':len(vals)//n,'type':typ}
 if mins is not None:q.update(min=mins,max=maxs)
 acc.append(q);return len(acc)-1
prims=[]
for side in range(2):
 z=0.001 if side==0 else -0.001
 ps=[-.5625,-1,z,.5625,-1,z,.5625,1,z,-.5625,1,z]
 pos=accessor(ps,'VEC3',mins=[-.5625,-1,z],maxs=[.5625,1,z]);norm=accessor([0,0,1 if side==0 else -1]*4,'VEC3')
 uv=accessor(([0,1,1,1,1,0,0,0] if side==0 else [1,1,0,1,0,0,1,0]),'VEC2')
 idx=accessor([0,1,2,0,2,3] if side==0 else [0,2,1,0,3,2],'SCALAR',5123)
 prims.append({'attributes':{'POSITION':pos,'NORMAL':norm,'TEXCOORD_0':uv},'indices':idx,'material':side,'mode':4})
imgs=[{'bufferView':view((o/(s+'.png')).read_bytes()),'mimeType':'image/png'} for s in ['front','back']]
g={'asset':{'version':'2.0','generator':'Original local dual-plane card'},'scene':0,'scenes':[{'nodes':[0]}],'nodes':[{'mesh':0,'name':'Dual-sided card'}],'meshes':[{'primitives':prims}],'materials':[{'name':s,'doubleSided':False,'alphaMode':'MASK','alphaCutoff':0.5,'pbrMetallicRoughness':{'baseColorTexture':{'index':i},'metallicFactor':0,'roughnessFactor':1}} for i,s in enumerate(['front','back'])],'images':imgs,'samplers':[{'magFilter':9729,'minFilter':9729,'wrapS':33071,'wrapT':33071}],'textures':[{'sampler':0,'source':i} for i in range(2)],'accessors':acc,'bufferViews':views,'buffers':[{'byteLength':len(blob)}]}
j=json.dumps(g,separators=(',',':')).encode();j+=b' '*((-len(j))%4);blob+=b'\0'*((-len(blob))%4)
glb=struct.pack('<III',0x46546C67,2,12+8+len(j)+8+len(blob))+struct.pack('<II',len(j),0x4E4F534A)+j+struct.pack('<II',len(blob),0x004E4942)+blob
(o/'card.glb').write_bytes(glb);(o/'gltf.json').write_text(json.dumps(g,indent=2),encoding='utf-8')
deck=o/'card.pptx'
def run(*args):
 r=subprocess.run([a.officecli,*map(str,args)],capture_output=True,text=True,encoding='utf-8');assert r.returncode==0,r.stdout+r.stderr;return r.stdout
run('create',deck);run('set',deck,'/','--prop','slideWidth=33.87cm','--prop','slideHeight=19.05cm')
ops=[]
for n,angle in enumerate([0,180,0],1):
 ops.append(dict(command='add',parent='/',type='slide',props={'layout':'blank','background':'102236','transition':'morph-2000'}))
 ops.append(dict(command='add',parent=f'/slide[{n}]',type='model3d',props={'name':'!!dual-card','src':str((o/'card.glb').resolve()),'x':'12.8cm','y':'2.2cm','width':'8.27cm','height':'14.7cm','rotation':f'0,{angle},0'}))
 ops.append(dict(command='add',parent=f'/slide[{n}]',type='shape',props={'name':'caption','text':'LOCAL 3D / TWO REAL FACES','x':'2cm','y':'17.3cm','width':'29cm','height':'0.7cm','font':'Arial','size':'14','color':'FFFFFF','fill':'none','line':'none'}))
(o/'batch.json').write_text(json.dumps(ops),encoding='utf-8');res=run('batch',deck,'--input',o/'batch.json','--json');(o/'build-log.json').write_text(res,encoding='utf-8');assert json.loads(res)['success'];run('close',deck)
(o/'validation.json').write_text(run('validate',deck,'--json'),encoding='utf-8')
with zipfile.ZipFile(deck) as z:
 embedded=[n for n in z.namelist() if n.endswith('.glb')];assert embedded and all(z.read(n)==glb for n in embedded)
print(json.dumps({'out':str(o),'glbBytes':len(glb),'embedded':embedded}))
