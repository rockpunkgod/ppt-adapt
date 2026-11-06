import argparse,json,zipfile,xml.etree.ElementTree as E,posixpath,subprocess,hashlib
from pathlib import Path
import numpy as np
from PIL import Image,ImageDraw
p=argparse.ArgumentParser();p.add_argument('--initial',required=True);p.add_argument('--edited',required=True);p.add_argument('--deck',required=True);p.add_argument('--resize',required=True);p.add_argument('--plan',required=True);p.add_argument('--live',required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--out',required=True);a=p.parse_args();o=Path(a.out)
if o.exists():raise FileExistsError('Refusing overwrite')
o.mkdir(parents=True);N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','s':'http://schemas.microsoft.com/office/powerpoint/2016/summaryzoom','p14':'http://schemas.microsoft.com/office/powerpoint/2010/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

def deck(path):
 with zipfile.ZipFile(path) as z:
  pr=E.fromstring(z.read('ppt/presentation.xml'));relations={x.get('Id'):posixpath.normpath('ppt/'+x.get('Target')) for x in E.fromstring(z.read('ppt/_rels/presentation.xml.rels'))};order=[(x.get('id'),relations[x.get('{'+N['r']+'}id')]) for x in pr.findall('p:sldIdLst/p:sldId',N)];sections={x.get('id'):{'name':x.get('name'),'ids':[e.get('id') for e in x.findall('p14:sldIdLst/p14:sldId',N)]} for x in pr.findall('.//p14:section',N)};summaries=[]
  for sid,part in order:
   root=E.fromstring(z.read(part))
   for s in root.findall('.//s:summaryZm',N):
    rows=[]
    for obj in s.findall('s:summaryZmObj',N):
     sec=sections[obj.get('sectionId')];xf=obj.find('.//a:xfrm',N);ex=xf.find('a:ext',N);of=xf.find('a:off',N);assert abs(int(ex.get('cx'))/int(ex.get('cy'))-int(pr.find('p:sldSz',N).get('cx'))/int(pr.find('p:sldSz',N).get('cy')))<.00001;rows.append({'sectionId':obj.get('sectionId'),'name':sec['name'],'zoomId':obj.find('s:zmPr',N).get('id'),'x':int(of.get('x')),'y':int(of.get('y')),'cx':int(ex.get('cx')),'cy':int(ex.get('cy')),'slideIds':sec['ids']})
    assert s.find('s:gridLayout',N) is not None;summaries.append({'slideId':sid,'items':rows})
  assert len(summaries)==1 and summaries[0]['slideId']==order[0][0];assert len(order)==7 and len(sections)==4
  return {'order':order,'sections':sections,'summary':summaries[0]}
initial=deck(a.initial);edited=deck(a.edited);final=deck(a.deck);assert len(initial['summary']['items'])==2 and len(edited['summary']['items'])==3;assert final['summary']==edited['summary'];assert [x['name'] for x in final['summary']['items']]==['Field Notes','Signal Lab','Next Route'];assert [x['sectionId'] for x in initial['summary']['items']]==[x['sectionId'] for x in final['summary']['items'][:2]]
reflow={phase:deck(Path(a.resize)/(phase+'.pptx'))['summary']['items'] for phase in ['initial','wide','narrow']}
for phase in ['initial','narrow']:assert reflow[phase][0]['x']==reflow[phase][1]['x'] and reflow[phase][0]['y']!=reflow[phase][1]['y']
assert reflow['wide'][0]['y']==reflow['wide'][1]['y'] and reflow['wide'][0]['x']!=reflow['wide'][1]['x']
live=Path(a.live);log=json.loads((live/'clicks.json').read_text(encoding='utf-8-sig'));plan=json.loads(Path(a.plan).read_text());assert log['passed'] and len(log['events'])==9;assert [x['actual'] for x in log['events']]==[2,3,1,4,5,1,6,7,1];assert hashlib.sha256(Path(a.deck).read_bytes()).hexdigest().upper()==log['deckShaBefore'];assert log['normalAppActivate'] and log['foregroundSet'] and log['ownWindowMoved']
rect=log['contentRect'];captures=[]
def content(path):
 im=Image.open(path).convert('RGB');return im.crop((rect[0],rect[1],rect[0]+rect[2],rect[1]+rect[3])).resize((1600,900),Image.Resampling.LANCZOS)
def roi_mask(im):
 ar=np.asarray(im);return np.sum(ar[round(7/19.05*900):round(13/19.05*900),round(12/33.87*1600):round(31/33.87*1600)].astype(float),axis=2)<420
for event in log['events']:
 name=event['name'];actual=content(live/(name+'.png'));expected=Image.open(live/('native-'+str(event['expected'])+'.png')).convert('RGB');aa=np.asarray(actual,float);ee=np.asarray(expected,float);flat=[]
 for x,y in [(3,12),(17,13),(31,3),(6,6),(6,13)]:
  px,py=round(x/33.87*1600),round(y/19.05*900);err=float(np.max(abs(np.mean(aa[py-2:py+3,px-2:px+3],axis=(0,1))-np.mean(ee[py-2:py+3,px-2:px+3],axis=(0,1)))));flat.append(err)
 assert max(flat)<=2,(name,flat)
 row={'name':name,'expected':event['expected'],'maxFlatRGBError':max(flat)}
 if event['expected']!=1:
  mask=roi_mask(actual);good=roi_mask(expected);other=event['expected']+1 if event['expected']%2==0 else event['expected']-1;wrong=roi_mask(Image.open(live/('native-'+str(other)+'.png')).convert('RGB'));iou=lambda x,y:float(np.sum(x&y)/max(1,np.sum(x|y)));giou,wiou=iou(mask,good),iou(mask,wrong);assert giou>.55 and giou>wiou+.15,(name,giou,wiou);row.update(textIoU=giou,wrongPageTextIoU=wiou)
 captures.append(row)
# Read every8fps frame and emit contact sheets, not a full-rate viewing claim.
frames=o/'frames';frames.mkdir();r=subprocess.run([a.ffmpeg,'-v','error','-i',str(live/'summary-live.mp4'),'-vf','fps=8,scale=320:-2',str(frames/'f-%03d.jpg')],capture_output=True);assert r.returncode==0,r.stderr
files=sorted(frames.glob('*.jpg'));assert len(files)==128,len(files)
for start in range(0,len(files),32):
 sheet=Image.new('RGB',(1280,8*214),'#172332');draw=ImageDraw.Draw(sheet)
 for k,file in enumerate(files[start:start+32]):
  im=Image.open(file);x=(k%4)*320;y=(k//4)*214;sheet.paste(im,(x,y+20));draw.text((x+5,y+3),f'{(start+k)/8:.3f}s',fill='white')
 sheet.save(o/f'contact-{start//32+1}.jpg',quality=92)
report={'passed':True,'trueSummaryNamespace':N['s'],'initialItems':2,'nativeEditedItems':3,'nativeReflowColumns':{'initial':1,'wide':2,'narrow':1},'cardAspectRatioPreserved':True,'physicalClicks':9,'branchOrder':[2,3,1,4,5,1,6,7,1],'desktopChecks':captures,'movieSampleFrames':len(files),'deckShaUnchanged':True,'scope':'UI-created Summary+native EditSummary and grid reflow, all three two-slide sections with real physical end return; no hand-grouped SectionZoom substitute'}
(o/'audit.json').write_text(json.dumps(report,indent=2));print(json.dumps({'passed':True,'captures':len(captures),'minimumTextIoU':min(x.get('textIoU',1) for x in captures),'frames':len(files)}))
