import argparse,json,subprocess,zipfile,posixpath,math,xml.etree.ElementTree as E
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--fixture',type=Path,required=True);p.add_argument('--deck',type=Path,required=True);p.add_argument('--native',type=Path,required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--screen-byte-compensation',type=int,default=0,choices=[0,1]);p.add_argument('--report',type=Path,required=True);a=p.parse_args()
if a.report.exists():raise FileExistsError(a.report)
def dec(f,fmt):return subprocess.run([a.ffmpeg,'-v','error','-i',str(f),'-f','rawvideo','-pix_fmt',fmt,'-'],capture_output=True,check=True).stdout
src=dec(a.fixture/'neutral-probe.png','rgba');ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'};R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed';checks={};rows=[]
with zipfile.ZipFile(a.deck) as z:
 for n,bg in [(2,(22,42,56)),(3,(215,173,89)),(4,(113,142,121))]:
  root=E.fromstring(z.read(f'ppt/slides/slide{n}.xml'));blip=root.find('.//p:pic/p:blipFill/a:blip',ns);colors=blip.findall('a:duotone/a:srgbClr',ns);screen=blip.find('a:fillOverlay',ns);checks[f's{n}NativeInvertScreenChain']=[c.get('val') for c in colors]==['FFFFFF','000000'] and screen is not None and screen.get('blend')=='screen' and screen.find('a:solidFill/a:srgbClr',ns).get('val')==''.join(f'{min(255,c+a.screen_byte_compensation):02X}' for c in bg)
  checks[f's{n}NotClaimedTrueAlpha']=not blip.findall('.//a:alpha',ns)
  rels={r.get('Id'):r.get('Target') for r in E.fromstring(z.read(f'ppt/slides/_rels/slide{n}.xml.rels'))};target=rels[blip.get(R)];part=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('ppt/slides',target));checks[f's{n}OriginalScanBytes']=z.read(part)==(a.fixture/'neutral-probe.png').read_bytes()
  d=dec(a.native/f'slide-{n}.png','rgb24');errors=[];paper=[];faint=[]
  def g(x,y):return src[(y*640+x)*4]
  for y in sorted(set(range(20,340,3))|{300}):
   for x in range(20,620,3):
    v=g(x,y);nb=[g(x+dx,y+dy) for dx in [-1,0,1] for dy in [-1,0,1]]
    if max(nb)-min(nb)>2:continue
    # Screen(invert(source), solid backdrop): 1-(1-I)*(1-B).
    expected=[round(255-(v*(255-c)/255)) for c in bg];j=((120+2*y+1)*1600+160+2*x+1)*3;err=[abs(d[j+c]-expected[c]) for c in range(3)];errors.extend(err)
    if v==255:paper.extend(err)
    if v==243:faint.extend(err)
  rms=lambda v:math.sqrt(sum(e*e for e in v)/len(v));checks[f's{n}NativeScreenFormula']=rms(errors)<=2;checks[f's{n}WhitePaperVisuallyGone']=rms(paper)<=0.5;checks[f's{n}FaintTraceRetained']=len(faint)>=300 and rms(faint)<=2;rows.append({'slide':n,'effectInputRGB':[min(255,c+a.screen_byte_compensation) for c in bg],'pageBackdropRGB':list(bg),'actualFlatPaperRGB':list(d[(200*1600+250)*3:(200*1600+250)*3+3]),'rms':rms(errors),'paperRMS':rms(paper),'faintSamples':len(faint)//3,'faintRMS':rms(faint)})
r={'passed':all(checks.values()),'checks':checks,'measurements':rows,'scope':'Native invert plus Screen over fixed solid backdrop; visual removal only, not image transparency, arbitrary photograph or automatically live background.'};a.report.write_text(json.dumps(r,indent=2),encoding='utf-8');print(json.dumps(r,indent=2))
