"""Read-only combined hover + true Slide Zoom + aperture + live input audit."""
import argparse,json,zipfile,xml.etree.ElementTree as E,subprocess,hashlib,posixpath
from pathlib import Path
from PIL import Image
p=argparse.ArgumentParser();p.add_argument('--deck',type=Path,required=True);p.add_argument('--live-dir',type=Path,required=True);p.add_argument('--ffprobe',required=True);p.add_argument('--input-plan',type=Path);p.add_argument('--structure-only',action='store_true');p.add_argument('--native-dir',type=Path);p.add_argument('--output',type=Path,required=True);a=p.parse_args()
if a.output.exists():p.error('Fresh output required')
n={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','mc':'http://schemas.openxmlformats.org/markup-compatibility/2006','z':'http://schemas.microsoft.com/office/powerpoint/2016/slidezoom','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'};c={}
def need(k,x):c[k]=bool(x)
with zipfile.ZipFile(a.deck) as z:
 pres=E.fromstring(z.read('ppt/presentation.xml'));slideids=[x.get('id') for x in pres.findall('p:sldIdLst/p:sldId',n)]
 for s in [1,2,3]:
  rels={x.get('Id'):x.get('Target') for x in E.fromstring(z.read(f'ppt/slides/_rels/slide{s}.xml.rels'))};r=E.fromstring(z.read(f'ppt/slides/slide{s}.xml'));tree=r.find('p:cSld/p:spTree',n)
  named={q.find('p:nvSpPr/p:cNvPr',n).get('name'):q for q in tree.findall('p:sp',n)};mask=named['!!return-mask'];path=mask.find('p:spPr/a:custGeom/a:pathLst/a:path',n)
  need(str(s)+'-true-two-apertures',path is not None and len(path.findall('a:moveTo',n))==3 and len(path.findall('a:close',n))==3)
  parts=[];current=[]
  for q in path:
   if q.tag.endswith('}moveTo') and current:parts.append(current);current=[]
   pt=q.find('a:pt',n)
   if pt is not None:current.append((int(pt.get('x')),int(pt.get('y'))))
  parts.append(current)
  def area(q):return sum(q[i][0]*q[(i+1)%len(q)][1]-q[(i+1)%len(q)][0]*q[i][1] for i in range(len(q)))
  need(str(s)+'-opposite-hole-winding',len(parts)==3 and area(parts[0])*area(parts[1])<0 and area(parts[0])*area(parts[2])<0)
  maskindex=list(tree).index(mask)
  need(str(s)+'-mask-is-above-cards-and-zoom',all(list(tree).index(q)<maskindex for q in tree if q.tag.endswith('}AlternateContent') or (q.tag.endswith('}sp') and q.find('p:nvSpPr/p:cNvPr',n).get('name','').startswith('!!card-'))))
  hover=mask.find('p:nvSpPr/p:cNvPr/a:hlinkHover',n)
  need(str(s)+'-background-return-hover',(hover is None) if s==1 else hover is not None and rels[hover.get('{'+n['r']+'}id')].endswith('slide1.xml'))
  for j in [1,2]:
   card=named['!!card-'+str(j)];h=card.find('p:nvSpPr/p:cNvPr/a:hlinkHover',n)
   need(f'{s}-{j}-self-hover-disabled-or-other-enabled',(h is None) if s==j+1 else h is not None and rels[h.get('{'+n['r']+'}id')].endswith('slide'+str(j+1)+'.xml'))
   block=next(x for x in tree.findall('mc:AlternateContent',n) if x.find(f".//p:cNvPr[@name='!!zoom-{j}']",n) is not None)
   obj=block.find('mc:Choice/p:graphicFrame/a:graphic/a:graphicData/z:sldZm/z:sldZmObj',n);need(f'{s}-{j}-real-slide-zoom',obj is not None and obj.get('sldId')==slideids[j+2])
   props=obj.find('z:zmPr',n);need(f'{s}-{j}-native-return-and-500ms',props.get('returnToParent','true') in ('1','true') and props.get('transitionDur')=='500')
   need(f'{s}-{j}-not-section-zoom','sectionZm' not in E.tostring(block,encoding='unicode'))
   if a.native_dir:
    blip=props.find('.//a:blip',n);rid=blip.get('{'+n['r']+'}embed');media=rels[rid].lstrip('/') if rels[rid].startswith('/') else posixpath.normpath('ppt/slides/'+rels[rid]);need(f'{s}-{j}-poster-is-native-target-export',z.read(media)==(a.native_dir/f'slide-{j+3}.png').read_bytes())
 if a.structure_only:
  result={'passed':all(c.values()),'checks':c,'scope':'Structure and native-poster bytes only; no fresh playback claimed'};a.output.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps({'passed':result['passed'],'checks':len(c),'failed':[k for k,v in c.items() if not v]},indent=2));raise SystemExit(0 if result['passed'] else 1)
 report=json.loads((a.live_dir/'report.json').read_text(encoding='utf-8-sig'));need('actual-physical-run-passed',report.get('passed') is True)
 events=report.get('events',[]);need('21-states-checked',len(events)==21)
 for event in events:need('live-'+event['name'],event['passed'] and event['samples']==[event['expected']]*6)
 need('pointer-coordinates-within-one-physical-pixel',all(max(abs(x-y) for x,y in zip(m['actual'],m['expected']))<=1 for m in report['moves']))
 need('all-pointer-endpoints-inside-real-slide',all(m['actualInsideSlide'] for m in report['moves']))
 need('continuous-physical-move-traces-recorded',all(len(m['trace'])>0 and m['input']=='small relative mouse MOVE only' for m in report['moves']))
 need('slide-aspect-matches-document',abs(report['slidePhysical'][2]-report['slidePhysical'][3]*report['documentAspect'])<1)
 duration=float(subprocess.run([a.ffprobe,'-v','error','-show_entries','format=duration','-of','default=noprint_wrappers=1:nokey=1',str(a.live_dir/'hover-zoom-live.mp4')],check=True,capture_output=True,text=True).stdout.strip())
 need('movie-covers-all-actual-states',max(e['at'] for e in events)+0.2<duration)
 scale=report['scale'];ox,oy=report['offset']
 def color(im,x,y):return im.getpixel((round(ox+x*scale),round(oy+y*scale)))
 def equal(rgb,want):return max(abs(x-y) for x,y in zip(rgb,want))<=3
 for event in events:
  im=Image.open(a.live_dir/(event['name']+'.png')).convert('RGB');state=event['expected'];label=event['name']
  if state<=3:
   for j,x in [(1,3.2),(2,18.5)]:need(label+'-card-'+str(j)+'-pixel',equal(color(im,x,10),(188,217,208) if state==j+1 else (213,226,219)))
   need(label+'-mask-hides-outside-poster-and-label',equal(color(im,8,17),(243,241,235)))
  else:
   for x,want in [(4,(217,177,144)),(13.7,(172,178,208)),(23.4,(168,205,191))]:need(label+'-detail-node-'+str(x),equal(color(im,x,8.5),want))
 if a.input_plan:
  plan=json.loads(a.input_plan.read_text(encoding='utf-8'));need('tested-input-file-unchanged',hashlib.sha256(a.deck.read_bytes()).hexdigest()==plan['inputSha256'])
 need('actual-desktop-recording',(a.live_dir/'hover-zoom-live.mp4').stat().st_size>10000)
result={'passed':all(c.values()),'checks':c,'movieSeconds':duration,'lastActualStateAt':max(e['at'] for e in events),'limits':'Only the two original branches, repeated twice, with six stable observations per state. Actual movie duration must cover the recorded events. Pixel and intermediate movie review supplement actual slide indices.'}
a.output.write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps({'passed':result['passed'],'checks':len(c),'failed':[k for k,v in c.items() if not v]},indent=2));raise SystemExit(0 if result['passed'] else 1)
