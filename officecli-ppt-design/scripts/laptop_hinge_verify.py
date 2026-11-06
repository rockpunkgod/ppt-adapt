"""Whole-scene audit: editable planes, mirrored hinge helpers, live media and pixels."""
import argparse,json,zipfile,hashlib,posixpath,subprocess,xml.etree.ElementTree as E
from pathlib import Path
import numpy as np
from PIL import Image
p=argparse.ArgumentParser();p.add_argument('--deck',type=Path,required=True);p.add_argument('--native-dir',type=Path,required=True);p.add_argument('--assets-dir',type=Path,required=True);p.add_argument('--live-dir',type=Path,required=True);p.add_argument('--ffprobe',required=True);p.add_argument('--input-plan',type=Path);p.add_argument('--structure-only',action='store_true');p.add_argument('--output',type=Path,required=True);a=p.parse_args()
if a.output.exists():p.error('Fresh verification output required')
n={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships','p14':'http://schemas.microsoft.com/office/powerpoint/2010/main'};checks={};metrics={}
def need(k,v):checks[k]=bool(v)
def canonical(name):return name.lstrip('!')
def descendants(root):
 items={}
 def walk(node,matrix=(1,1,0,0)):
  sx,sy,tx,ty=matrix
  for child in node:
   if child.tag==f"{{{n['p']}}}grpSp":
    name=canonical(child.find('p:nvGrpSpPr/p:cNvPr',n).get('name'));tr=child.find('p:grpSpPr/a:xfrm',n);off=tr.find('a:off',n);ext=tr.find('a:ext',n);co=tr.find('a:chOff',n);ce=tr.find('a:chExt',n);x,y,w,h=[int(z) for z in [off.get('x'),off.get('y'),ext.get('cx'),ext.get('cy')]]
    items[name]={'node':child,'bbox':(tx+sx*x,ty+sy*y,sx*w,sy*h)};rx=w/int(ce.get('cx'));ry=h/int(ce.get('cy'));walk(child,(sx*rx,sy*ry,tx+sx*(x-rx*int(co.get('x'))),ty+sy*(y-ry*int(co.get('y')))))
   elif child.tag==f"{{{n['p']}}}sp":
    name=canonical(child.find('p:nvSpPr/p:cNvPr',n).get('name'));tr=child.find('p:spPr/a:xfrm',n);off=tr.find('a:off',n);ext=tr.find('a:ext',n);items[name]={'node':child,'bbox':(tx+sx*int(off.get('x')),ty+sy*int(off.get('y')),sx*int(ext.get('cx')),sy*int(ext.get('cy')))}
 return walk(root) or items
def near(x,y):return abs(x-y)<200
def part(target):return target.lstrip('/') if target.startswith('/') else posixpath.normpath('ppt/slides/'+target)
with zipfile.ZipFile(a.deck) as z:
 need('no-true-3d-model-substitute',not any(x.lower().endswith(('.glb','.gltf')) for x in z.namelist()))
 for i in [1,2]:
  root=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'));items=descendants(root.find('p:cSld/p:spTree',n));rels={x.get('Id'):x.get('Target') for x in E.fromstring(z.read(f'ppt/slides/_rels/slide{i}.xml.rels'))};need(str(i)+'-48-editable-key-shapes',sum(k.startswith('key-') for k in items)==48)
  need(str(i)+'-morph-two-seconds',any(q.get('{'+n['p14']+'}dur')=='2000' for q in root.findall('.//p:transition',n)))
  for kind,angle in [('screen',90 if i==1 else 0),('keyboard',270 if i==1 else 290)]:
   group=items[kind+'-hinge'];x,y,w,h=group['bbox'];cam=group['node'].find('p:grpSpPr/a:scene3d/a:camera',n);rot=cam.find('a:rot',n)
   need(f'{i}-{kind}-native-perspective-camera',cam.get('prst')=='perspectiveRelaxedModerately' and cam.get('fov')=='2700000')
   need(f'{i}-{kind}-native-Y-camera-latitude',int(rot.get('lat'))%(360*60000)==angle*60000 and int(rot.get('lon','0'))==0 and int(rot.get('rev','0'))==0)
   need(f'{i}-{kind}-common-hinge-center',near(y+h/2,11.9*360000) and near(x+w/2,16.935*360000))
   visible=items[kind+'-face']['bbox'];helper=items[kind+'-axis-helper'];hx,hy,hw,hh=helper['bbox'];vx,vy,vw,vh=visible
   need(f'{i}-{kind}-equal-size-axis-compensation',near(hw,vw) and near(hh,vh) and near(hx,vx))
   need(f'{i}-{kind}-opposite-side-helper',near(hy+hh,vy) if kind=='keyboard' else near(vy+vh,hy))
   need(f'{i}-{kind}-helper-invisible',helper['node'].find('p:spPr/a:noFill',n) is not None)
   shell=items['screen-bezel' if kind=='screen' else 'keyboard-shell']['node'];three=shell.find('p:spPr/a:sp3d',n)
   need(f'{i}-{kind}-3pt-extrusion',three is not None and three.get('extrusionH')=='38100')
   bevel=three.find('a:bevelB',n);need(f'{i}-{kind}-5pt-bottom-bevel',bevel is not None and bevel.get('w')=='63500' and bevel.get('h')=='63500')
  canvas=items['screen-canvas'];blip=canvas['node'].find('p:spPr/a:blipFill/a:blip',n);need(str(i)+'-matching-original-screen-poster',z.read(part(rels[blip.get('{'+n['r']+'}embed')]))==(a.assets_dir/'first-frame.png').read_bytes())
  pics=root.findall('.//p:pic',n);videos=[x for x in pics if x.find('p:nvPicPr/p:nvPr/a:videoFile',n) is not None]
  need(str(i)+'-video-only-on-open-page',len(videos)==(1 if i==2 else 0))
  if i==2:
   v=videos[0];vid=v.find('p:nvPicPr/p:cNvPr',n).get('id');tr=v.find('p:spPr/a:xfrm',n);off=tr.find('a:off',n);ext=tr.find('a:ext',n);box=tuple(map(int,[off.get('x'),off.get('y'),ext.get('cx'),ext.get('cy')]));need('flat-video-aligned-to-open-canvas',all(near(x,y) for x,y in zip(box,canvas['bbox'])))
   need('video-is-independent-open-plane',v.find('p:spPr/a:scene3d',n) is None)
   native_media=v.find('.//p14:media',n);need('actual-original-video-embedded',native_media is not None and z.read(part(rels[native_media.get('{'+n['r']+'}embed')]))==(a.assets_dir/'original-screen.mp4').read_bytes())
   main=root.find('.//p:cTn[@nodeType="mainSeq"]',n);entries=main.findall('.//p:cTn[@presetClass="entr"]',n);plays=main.findall('.//p:cmd',n)
   need('fade-1-second-with-previous',len(entries)==1 and entries[0].get('presetID')=='10' and entries[0].get('nodeType')=='withEffect' and any(t.get('dur')=='1000' for t in entries[0].findall('.//p:cTn',n)))
   need('native-play-node-starts-with-previous',any(q.get('nodeType')=='withEffect' for q in main.findall('.//p:cTn[@presetClass="mediacall"]',n)))
   need('real-media-play-with-previous',any(q.get('cmd','').startswith('playFrom(') and q.find('p:cBhvr/p:tgtEl/p:spTgt',n).get('spid')==vid for q in plays))
rows=json.loads((a.native_dir/'3d-readback.json').read_text(encoding='utf-8-sig'));media=next(x for x in rows if 'reopenedVideo' in x);need('native-reopened-real-video',media['type']==16 and media['mediaType']==3 and media['lengthMs']==6000 and media['effects']==2 and media['loop']==-1)
metrics['backend']='hybrid';metrics['boundary']='Two-plane perspective animation uses matching posters during hinge motion. Live video is a separate flat layer on fully open page; not a claim of a live video texture bound to a rotating 3D surface.'
if not a.structure_only:
 report=json.loads((a.live_dir/'report.json').read_text(encoding='utf-8-sig'));need('real-native-input-run-passed',report['passed']);need('all-eight-real-states',len(report['events'])==8 and all(x['passed'] for x in report['events']));need('physical-open-close-reopen', [x['key'] for x in report['keys']]==[39,36,39])
 players=[x for x in report['players'] if isinstance(x.get('positionMs'),(int,float))];metrics['numericPlayerSamples']=len(players)
 for label,lo,hi in [('first',3.2,9.4),('second',14.2,22.2)]:
  samples=[x for x in players if lo<x['at']<hi];positions=[x['positionMs'] for x in samples];need(label+'-actual-playback-progress',len(samples)>8 and max(positions)-min(positions)>1500);need(label+'-native-loop-observed',any(x['positionMs']-y['positionMs']>3000 for x,y in zip(samples,samples[1:])))
 scale=report['scale'];ox,oy=report['offset'];box=[round(ox+8.435*scale),round(oy+1.8*scale),round(ox+25.435*scale),round(oy+11.3625*scale)];metrics['videoScreenRectPx']=box
 def image(name):return np.asarray(Image.open(a.live_dir/(name+'.png')).convert('RGB')).astype(int)
 first=image('open-live-first');late=image('open-live-late');x0,y0,x1,y1=box;inside=np.zeros(first.shape[:2],bool);inside[y0:y1,x0:x1]=True;expanded=inside.copy();expanded[y0-3:y1+3,x0-3:x1+3]=True
 scene=np.zeros(inside.shape,bool);scene[round(oy+1.1*scale):round(oy+16.7*scale),round(ox+5*scale):round(ox+28.8*scale)]=True
 diff=np.max(abs(first-late),2);metrics['screenMeanFrameDifference']=float(np.mean(abs(first[inside]-late[inside])));metrics['outsideScreenMaxDifference']=int(diff[scene&~expanded].max());need('actual-pixels-move-inside-screen',metrics['screenMeanFrameDifference']>2);need('live-video-does-not-change-bezel-or-keyboard',metrics['outsideScreenMaxDifference']==0)
 closed0=image('closed-before');closed1=image('closed-after-video');metrics['closedBeforeAfterMaxDifference']=int(np.max(abs(closed0[scene]-closed1[scene])));need('actual-close-restores-identical-scene',metrics['closedBeforeAfterMaxDifference']==0)
 duration=float(subprocess.run([a.ffprobe,'-v','error','-show_entries','format=duration','-of','default=noprint_wrappers=1:nokey=1',str(a.live_dir/'laptop-live.mp4')],check=True,capture_output=True,text=True).stdout.strip());metrics['movieSeconds']=duration;need('whole-recording-covers-all-eight-states',max(x['at'] for x in report['events'])+0.2<duration)
 if a.input_plan:
  plan=json.loads(a.input_plan.read_text(encoding='utf-8'));need('input-deck-unchanged',hashlib.sha256(a.deck.read_bytes()).hexdigest()==plan['inputSha256'])
result={'passed':all(checks.values()),'checks':checks,'metrics':metrics};a.output.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps({'passed':result['passed'],'checks':len(checks),'failed':[k for k,v in checks.items() if not v],'metrics':metrics},indent=2));raise SystemExit(0 if result['passed'] else 1)
