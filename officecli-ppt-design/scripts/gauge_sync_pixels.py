"""Independently infer arc/needle angle and digit glyphs from native PNG."""
import argparse,json,math,zipfile,xml.etree.ElementTree as E
from pathlib import Path
import numpy as np
from PIL import Image,ImageDraw,ImageFont
W,H=33.87,19.05;cx,cy=16.935/W*1600,10/H*900;radius=8.5/W*1600
font=ImageFont.truetype('C:/Windows/Fonts/arialbd.ttf',256)
def structure(directory,native):
 o=Path(directory);native=Path(native);models=json.loads((o/'model.json').read_text());ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','c':'http://schemas.openxmlformats.org/drawingml/2006/chart'};checks=[]
 for stem,static in [('gauge',False),('final-reference',True)]:
  with zipfile.ZipFile(o/(stem+'.pptx')) as z:
   parts=sorted(n for n in z.namelist() if '/charts/' in n and n.endswith('.xml'));assert len(parts)==2
   for i,(part,model) in enumerate(zip(parts,models),1):
    chart=E.fromstring(z.read(part));v=model['value'];vals=[float(q.text) for q in chart.findall('.//c:val/c:numLit/c:pt/c:v',ns)];assert vals==[v,200-v] and 1<=v<=99
    assert chart.find('.//c:doughnutChart/c:holeSize',ns).get('val')=='80' and chart.find('.//c:firstSliceAng',ns).get('val')==('270' if static else '90')
    dp=chart.findall('.//c:dPt',ns);assert len(dp)==2 and dp[1].find('c:spPr/a:noFill',ns) is not None and chart.find('c:externalData',ns) is None
    root=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'));group=root.find('p:cSld/p:spTree/p:grpSp',ns);gid=group.find('p:nvGrpSpPr/p:cNvPr',ns).get('id');x=group.find('p:grpSpPr/a:xfrm',ns);off=x.find('a:off',ns);ext=x.find('a:ext',ns)
    assert abs((int(off.get('x'))+int(ext.get('cx'))/2)/360000-16.935)<0.0001 and abs((int(off.get('y'))+int(ext.get('cy'))/2)/360000-10)<0.0001
    names={q.find('p:nvSpPr/p:cNvPr',ns).get('name'):q for q in root.findall('.//p:sp',ns)};envelope=names['needle-envelope'];assert envelope.find('p:spPr/a:noFill',ns) is not None and envelope.find('p:spPr/a:ln/a:noFill',ns) is not None
    for name,count in [('tens',10),('ones',11)]:
     ps=names[name].findall('p:txBody/a:p',ns);assert len(ps)==count
     assert [''.join(q.itertext()) for q in names[name].findall('.//a:t',ns)]==list(map(str,range(10)))+(["0"] if name=='ones' else [])
     assert all(q.find('a:pPr/a:lnSpc/a:spcPts',ns).get('val')=='4535' for q in ps)
    effects=root.findall('.//p:cTn[@presetClass]',ns)
    if static:assert not effects and abs(int(x.get('rot'))/60000-1.8*v)<0.0001
    else:
     assert len(effects)==5;rots=[q for q in effects if q.find('.//p:animRot',ns) is not None];assert len(rots)==2
     assert [int(q.find('.//p:animRot',ns).get('by'))/60000 for q in rots]==[180,1.8*v] and rots[1].find('.//p:spTgt',ns).get('spid')==gid
     motions=[q for q in effects if q.find('.//p:animMotion',ns) is not None];assert len(motions)==3
     assert [q.find('.//p:animMotion/p:cBhvr/p:cTn',ns).get('dur') for q in motions]==['2000','250','1000']
     assert [q.find('.//p:animMotion/p:cBhvr/p:cTn',ns).get('fill') for q in motions]==['hold','remove','hold']
     assert motions[1].get('repeatCount')=='4000' and motions[2].find('p:stCondLst/p:cond',ns).get('delay')=='1000'
    checks.append({'deck':stem,'slide':i,'value':v,'needlePivotCm':[16.935,10],'trueChartValues':vals,'effectCount':len(effects)})
 objects=json.loads((native/'objects.json').read_text(encoding='utf-8-sig'));assert len(objects)==4
 for i,row in enumerate(objects):
  assert row['chartShapeType']==3 and row['groupType']==6
  if i<2:
   es=row['effects'];assert len(es)==5 and [q['target'] for q in es]==['gauge-data','needle-pivot','tens','ones','ones'];assert all(abs(es[k]['duration']-2)<0.001 for k in [0,1,2]);assert es[3]['duration']==0.25 and es[3]['repeatCount']==4 and es[4]['duration']==1 and es[4]['delay']==1
  else:assert not row['effects']
 (native/'structure-audit.json').write_text(json.dumps({'passed':True,'checks':checks,'editDataClaimed':False,'nativeTiming':'2s main effects /1s last-ones delay separate from screen boundary evidence'},indent=2),encoding='utf-8')
 return checks
def normalize(image):
 im=image.convert('RGB');ar=np.asarray(im);column=ar[:,int(im.width*0.94),:].astype(float);bg=np.max(np.abs(column-np.array([246,248,252])),axis=1)<8;ys=np.where(bg)[0];assert len(ys)>im.height*0.6
 return im.crop((0,int(ys.min()),im.width,int(ys.max())+1)).resize((1600,900))
def glyph(data):
 ys,xs=np.where(data);assert len(xs)>50;return Image.fromarray((data[min(ys):max(ys)+1,min(xs):max(xs)+1]*255).astype('uint8')).resize((64,96),Image.Resampling.NEAREST),float((max(xs)-min(xs)+1)/(max(ys)-min(ys)+1))
def expected(digit):
 im=Image.new('L',(400,400));ImageDraw.Draw(im).text((20,20),str(digit),font=font,fill=255);return glyph(np.asarray(im)>127)
def analyze(image,v):
 im=np.asarray(normalize(image),dtype=float);red,green,blue=[im[:,:,k] for k in range(3)];teal=(red<100)&(green>130)&(blue>100)&(green>red+60)
 scan=[]
 for angle in np.arange(0,180.01,0.1):
  x=int(round(cx-radius*0.9*math.cos(math.radians(angle))));y=int(round(cy-radius*0.9*math.sin(math.radians(angle))))
  if teal[y,x]:scan.append(angle)
 end=max(scan) if scan else 0;assert abs(end-1.8*v)<0.6,(v,end)
 coral=(red>180)&(green>60)&(green<160)&(blue<140)&(red>green*1.4);ys,xs=np.where(coral);dist=np.hypot(xs-cx,ys-cy);far=dist>250;angles=np.degrees(np.arctan2(-(ys[far]-cy),-(xs[far]-cx)));needle=float(np.median(angles));assert abs(needle-1.8*v)<0.6,(v,needle)
 digits=[]
 for x,target in [(15.3,v//10),(16.45,v%10)]:
  x0=int(x/W*1600);x1=int((x+1.1)/W*1600);y0=int(11.2/H*900);y1=int((11.2+45.35/72*2.54)/H*900);crop=im[y0:y1,x0:x1];mask=(crop[:,:,0]<100)&(crop[:,:,1]<120)&(crop[:,:,2]<150);actual,ratio=glyph(mask);scores=[]
  for d in range(10):
   exp,expRatio=expected(d);aa=np.asarray(actual)>127;bb=np.asarray(exp)>127;iou=float(np.count_nonzero(aa&bb)/np.count_nonzero(aa|bb));scores.append((iou,d,expRatio))
  best=max(scores);assert best[1]==target and best[0]>0.85 and abs(ratio/best[2]-1)<0.08,(target,best,ratio)
  digits.append({'expected':target,'recognized':best[1],'glyphIoU':best[0],'aspectRatio':ratio})
 return {'value':v,'visibleArcDegrees':end,'needleDegrees':needle,'digits':digits}

if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--directory',required=True);p.add_argument('--native',required=True);a=p.parse_args();o=Path(a.directory);native=Path(a.native);rows=[]
 structure(o,native)
 with zipfile.ZipFile(o/'final-reference.pptx') as z:
  parts=sorted(n for n in z.namelist() if '/charts/' in n and n.endswith('.xml'))
  for i,part in enumerate(parts,1):
   r=E.fromstring(z.read(part));n={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'};vals=[float(v.text) for v in r.findall('.//c:ser/c:val/c:numLit/c:pt/c:v',n)];assert len(vals)==2 and sum(vals)==200 and vals[0]==int(vals[0]);row=analyze(Image.open(native/f'final-reference-{i}.png'),int(vals[0]));row['chartValues']=vals;rows.append(row)
 (native/'static-audit.json').write_text(json.dumps({'passed':True,'rows':rows},indent=2),encoding='utf-8');print(json.dumps(rows,indent=2))
